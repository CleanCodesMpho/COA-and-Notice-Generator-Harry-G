import os
import sys
import traceback
import datetime
from typing import Dict, Any, Optional, Tuple

from arcgis.gis import GIS
from docxtpl import DocxTemplate
from jinja2 import Environment


# =========================================================
# CONFIG FROM ENVIRONMENT VARIABLES
# =========================================================
AGOL_URL = os.getenv("AGOL_URL", "https://www.arcgis.com")
AGOL_USERNAME = os.getenv("AGOL_USERNAME")
AGOL_PASSWORD = os.getenv("AGOL_PASSWORD")

SURVEY_ITEM_ID = os.getenv("SURVEY_ITEM_ID")
LAYER_INDEX = int(os.getenv("LAYER_INDEX", "0"))

# Dedicated processing field in your hosted layer
REPORT_STATUS_FIELD = os.getenv("REPORT_STATUS_FIELD", "report_status")

# Template file locations inside your Render service
COA_TEMPLATE_PATH = os.getenv("COA_TEMPLATE_PATH", "./templates/HGDM-COA (formal).docx")
NOTICE_TEMPLATE_PATH = os.getenv("NOTICE_TEMPLATE_PATH", "./templates/NOTICE OF COMPLIANCE V2.docx")

# Temporary working folder on Render
TEMP_DIR = os.getenv("TEMP_DIR", "/tmp/reports")

# Optional: attach only once per record even if field was reset manually
SKIP_IF_ATTACHMENT_ALREADY_EXISTS = os.getenv("SKIP_IF_ATTACHMENT_ALREADY_EXISTS", "false").lower() == "true"

os.makedirs(TEMP_DIR, exist_ok=True)


# =========================================================
# BASIC HELPERS
# =========================================================
def now_stamp() -> str:
    return datetime.datetime.now().strftime("%Y%m%d_%H%M%S")


def out_docx_path(object_id: Any, label: str = "report") -> str:
    safe_oid = str(object_id).replace("/", "_").replace("\\", "_")
    return os.path.join(TEMP_DIR, f"{label}_{safe_oid}_{now_stamp()}.docx")


def require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def validate_paths() -> None:
    missing = []
    if not os.path.exists(COA_TEMPLATE_PATH):
        missing.append(f"COA template not found: {COA_TEMPLATE_PATH}")
    if not os.path.exists(NOTICE_TEMPLATE_PATH):
        missing.append(f"NOTICE template not found: {NOTICE_TEMPLATE_PATH}")

    if missing:
        raise FileNotFoundError(" | ".join(missing))


def normalize_context(attrs: Dict[str, Any]) -> Dict[str, Any]:
    """
    Prepare attributes for docxtpl rendering.
    Converts None to empty string so templates render more safely.
    """
    context = {}
    for k, v in attrs.items():
        if v is None:
            context[k] = ""
        else:
            context[k] = v
    return context


# =========================================================
# TEMPLATE RENDERING
# =========================================================
jinja_env = Environment(
    variable_start_string='${',
    variable_end_string='}',
    enable_async=False
)


def render_docx_template(template_path: str, context: Dict[str, Any], out_path: str) -> str:
    tpl = DocxTemplate(template_path)
    tpl.render(context, jinja_env)
    tpl.save(out_path)
    return out_path


# =========================================================
# BUSINESS RULES
# =========================================================
def choose_template(attrs: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    """
    Decide which template to use.
    Keeps your current logic, but fixes lowercase comparisons.
    """
    inspection_status = str(attrs.get("inspection_status") or "").strip().lower()
    compliance = str(attrs.get("compliance") or "").strip().lower()
    action_taken = str(attrs.get("action_taken") or "").strip().lower()

    # COA
    if inspection_status == "completed" and compliance == "compliant":
        return COA_TEMPLATE_PATH, "COA"

    # Notice
    if action_taken == "compliance notice":
        return NOTICE_TEMPLATE_PATH, "NOTICE"

    return None, None


def get_object_id(attrs: Dict[str, Any]) -> Any:
    for key in ("objectid", "OBJECTID", "ObjectID"):
        if key in attrs and attrs[key] is not None:
            return attrs[key]
    raise KeyError("Could not find OBJECTID field in feature attributes.")


# =========================================================
# AGOL CONNECTION
# =========================================================
def connect_gis() -> GIS:
    require_env("AGOL_USERNAME")
    require_env("AGOL_PASSWORD")
    require_env("SURVEY_ITEM_ID")

    gis = GIS(
        AGOL_URL,
        AGOL_USERNAME,
        AGOL_PASSWORD
    )
    print(f"Logged into AGOL as: {gis.users.me.username}")
    return gis


def get_layer_or_table(gis: GIS):
    survey_item = gis.content.get(SURVEY_ITEM_ID)
    if not survey_item:
        raise RuntimeError(f"Survey item not found: {SURVEY_ITEM_ID}")

    try:
        layer = survey_item.layers[LAYER_INDEX]
        print(f"Using layer: {layer.url}")
        return layer
    except Exception:
        pass

    try:
        table = survey_item.tables[LAYER_INDEX]
        print(f"Using table: {table.url}")
        return table
    except Exception as e:
        raise RuntimeError("Could not load layer/table from survey item.") from e


def verify_processing_field(layer) -> None:
    field_names = {f["name"] for f in layer.properties.fields}
    if REPORT_STATUS_FIELD not in field_names:
        raise RuntimeError(
            f"Processing field '{REPORT_STATUS_FIELD}' does not exist in the layer. "
            f"Create a text field with that exact name."
        )


# =========================================================
# FEATURE / ATTACHMENT HELPERS
# =========================================================
def has_existing_attachments(layer, object_id: Any) -> bool:
    try:
        attachments = layer.attachments.get_list(oid=object_id)
        return len(attachments) > 0
    except Exception:
        return False


def attach_file(layer, object_id: Any, file_path: str) -> None:
    result = layer.attachments.add(rel_objectid=object_id, file_path=file_path)
    print(f"Attachment result for OBJECTID {object_id}: {result}")


def update_report_status(layer, feature, new_status: str) -> None:
    feature.attributes[REPORT_STATUS_FIELD] = new_status
    result = layer.edit_features(updates=[feature])
    print(f"Status update result: {result}")


# =========================================================
# QUERY / PROCESSING
# =========================================================
def build_unprocessed_where() -> str:
    return f"{REPORT_STATUS_FIELD} IS NULL OR {REPORT_STATUS_FIELD} = ''"


def process_new_records(layer, limit: Optional[int] = None) -> int:
    where_clause = build_unprocessed_where()
    print(f"Querying unprocessed records with: {where_clause}")

    features_result = layer.query(where=where_clause, out_fields="*", return_geometry=False)
    features = features_result.features if features_result and features_result.features else []

    print(f"Found {len(features)} unprocessed record(s) to evaluate).")

    processed_count = 0

    for feature in features:
        if limit is not None and processed_count >= limit:
            break

        try:
            attrs = feature.attributes
            object_id = get_object_id(attrs)

            print("\n----------------------------------------")
            print(f"Processing OBJECTID: {object_id}")

            template_path, label = choose_template(attrs)
            if not template_path:
                print(f"No template matched OBJECTID {object_id}. Leaving unprocessed.")
                continue

            if SKIP_IF_ATTACHMENT_ALREADY_EXISTS and has_existing_attachments(layer, object_id):
                print(f"OBJECTID {object_id} already has attachment(s). Marking as generated.")
                update_report_status(layer, feature, "generated")
                processed_count += 1
                continue

            context = normalize_context(attrs)
            docx_path = out_docx_path(object_id, label=label)

            render_docx_template(template_path, context, docx_path)
            print(f"Generated DOCX: {docx_path}")

            attach_file(layer, object_id, docx_path)
            print(f"Attached DOCX to OBJECTID {object_id}")

            update_report_status(layer, feature, "generated")
            print(f"Marked OBJECTID {object_id} as generated")

            try:
                os.remove(docx_path)
            except Exception:
                pass

            processed_count += 1

        except Exception as e:
            print(f"Error while processing a record: {e}")
            print(traceback.format_exc())

    print("\n========================================")
    print(f"Completed. Successfully handled {processed_count} record(s).")
    return processed_count


# =========================================================
# MAIN
# =========================================================
def main():
    validate_paths()
    gis = connect_gis()
    layer = get_layer_or_table(gis)
    verify_processing_field(layer)

    limit_env = os.getenv("PROCESS_LIMIT")
    limit = int(limit_env) if limit_env else None

    process_new_records(layer, limit=limit)


if __name__ == "__main__":
    try:
        main()
        sys.exit(0)
    except Exception as e:
        print(f"Fatal error: {e}")
        print(traceback.format_exc())
        sys.exit(1)
