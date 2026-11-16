import json
import os
import re
import shutil
import tempfile
import time
import zipfile
from pathlib import Path

import pandas as pd
import requests

from utils.file_io import read_yaml_config
from utils.ioe_workflow import add_date_stamp_to_path


DEFAULT_AQUA_CREDENTIALS_PATH = Path("domain-knowledge/wafer-map/AQUA API Key.yaml")
DEFAULT_AQUA_OUTPUT_PATH = add_date_stamp_to_path(Path("outputs/AQUA class_7820.xlsx"))
DEFAULT_AQUA_CA_BUNDLE = Path("utils/IntelChain.pem")
DEFAULT_AQUA_HOSTS = {
    "AMR": "https://aqua-api.intel.com",
    "QA": "https://aqua-intg-api.intel.com",
    "DEV": "https://aqua-test-api.intel.com",
}
DEFAULT_CLASS_OPERATION = "7820"
DEFAULT_POLL_SECONDS = 5
DEFAULT_TIMEOUT_SECONDS = 900
AQUA_RESULT_SHEET = "class_7820"
PREVIEW_COLUMNS = [
    "Visual_ID",
    "Lot",
    "Sort_Lot",
    "Sort_Wafer_ID",
    "Sort_X",
    "Sort_Y",
    "Operation",
    "SubFlowStep",
    "Unit_Functional_Bin",
    "Unit_Interface_Bin",
    "Unit_Good_Bad_Flag",
    "Start_Date_Time",
    "End_Date_Time",
    "Program_Name",
    "Testing_Entity",
    "Package",
]


def save_aqua_result_workbook(aqua_df, output_path=DEFAULT_AQUA_OUTPUT_PATH, sheet_name="AQUA class_7820"):
    output_workbook = Path(output_path)
    output_workbook.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_workbook, engine="openpyxl") as writer:
        aqua_df.to_excel(writer, index=False, sheet_name=sheet_name)
    return str(output_workbook)


def load_cached_aqua_result_workbook(
    final_workbook_path,
    cache_workbook_path=DEFAULT_AQUA_OUTPUT_PATH,
    vid_sheet_name="VID",
    cache_sheet_name="AQUA class_7820",
    vid_limit=None,
):
    cache_workbook = Path(cache_workbook_path)
    if not cache_workbook.exists():
        return None

    expected_vids = _load_true_vids_from_workbook(final_workbook_path, sheet_name=vid_sheet_name, limit=vid_limit)

    sheet_candidates = [cache_sheet_name]
    if AQUA_RESULT_SHEET not in sheet_candidates:
        sheet_candidates.append(AQUA_RESULT_SHEET)

    aqua_df = None
    loaded_sheet = None
    for sheet_name in sheet_candidates:
        try:
            aqua_df = pd.read_excel(cache_workbook, sheet_name=sheet_name)
            loaded_sheet = sheet_name
            break
        except ValueError:
            continue

    if aqua_df is None:
        return None
    if "Visual_ID" not in aqua_df.columns:
        return None

    cached_vids = []
    for value in aqua_df["Visual_ID"].dropna().astype(str):
        cleaned = value.strip()
        if cleaned:
            cached_vids.append(cleaned)
    cached_vids = list(dict.fromkeys(cached_vids))

    if len(expected_vids) != len(cached_vids) or set(expected_vids) != set(cached_vids):
        return None

    return {
        "status": "completed",
        "source": "cache",
        "vid_count": len(expected_vids),
        "vid_sample": expected_vids[:5],
        "operation": str(DEFAULT_CLASS_OPERATION),
        "row_count": int(len(aqua_df)),
        "result_df": aqua_df,
        "sheet_name": loaded_sheet,
        "output_path": str(cache_workbook),
    }


class AquaExecutionError(RuntimeError):
    pass


def _load_true_vids(vid_csv_path, limit=None):
    vid_path = Path(vid_csv_path)
    if not vid_path.exists():
        raise FileNotFoundError(f"VID.csv not found: {vid_path}")

    df = pd.read_csv(vid_path)
    if "VISUAL_ID" not in df.columns:
        raise ValueError("VID.csv must contain a VISUAL_ID column.")

    vids = []
    for value in df["VISUAL_ID"].dropna().astype(str):
        cleaned = value.strip()
        if cleaned and "_" not in cleaned:
            vids.append(cleaned)

    deduped = list(dict.fromkeys(vids))
    if limit is not None:
        deduped = deduped[: int(limit)]
    if not deduped:
        raise ValueError("VID.csv does not contain any true VID values.")
    return deduped


def _load_true_vids_from_workbook(workbook_path, sheet_name="VID", limit=None):
    workbook = Path(workbook_path)
    if not workbook.exists():
        raise FileNotFoundError(f"Final workbook not found: {workbook}")

    df = pd.read_excel(workbook, sheet_name=sheet_name)
    if "VISUAL_ID" not in df.columns:
        raise ValueError(f"Workbook sheet {sheet_name} must contain a VISUAL_ID column.")

    vids = []
    for value in df["VISUAL_ID"].dropna().astype(str):
        cleaned = value.strip()
        if cleaned and "_" not in cleaned:
            vids.append(cleaned)

    deduped = list(dict.fromkeys(vids))
    if limit is not None:
        deduped = deduped[: int(limit)]
    if not deduped:
        raise ValueError(f"Workbook sheet {sheet_name} does not contain any true VID values.")
    return deduped


def _load_aqua_credentials(credentials_path):
    payload = read_yaml_config(str(credentials_path))
    aqua = payload.get("aqua") if isinstance(payload, dict) else None
    if not isinstance(aqua, dict):
        raise ValueError("AQUA credentials file must contain an 'aqua' mapping.")

    environment = str(aqua.get("environment") or "AMR").strip().upper()
    user_id = str(aqua.get("user_id") or "").strip()
    api_key = str(aqua.get("api_key") or "").strip()
    if environment not in DEFAULT_AQUA_HOSTS:
        raise ValueError(f"Unsupported AQUA environment: {environment}")
    if not user_id:
        raise ValueError("AQUA credentials file is missing aqua.user_id")
    if not api_key:
        raise ValueError("AQUA credentials file is missing aqua.api_key")

    return {
        "environment": environment,
        "user_id": user_id,
        "api_key": api_key,
        "base_url": DEFAULT_AQUA_HOSTS[environment],
    }


def build_aqua_class_report_config(vids, operation=DEFAULT_CLASS_OPERATION):
    visual_id_lines = ",\n".join(f"\t\t= {vid}" for vid in vids)
    return (
        "@ Report : IOE_VID_CLASS\n\n"
        "####################\n"
        "@ FilterSet : IOE_VID_CLASS_FilterSet\n"
        "####################\n\n"
        "@ Domain : PostAssembly\n\n"
        "@ Attributes : \n"
        "\t@@ VisualId :\n"
        f"{visual_id_lines}\n"
        "\t;\n\n"
        "####################\n"
        "@ Parameters : IOE_VID_CLASS_Parameters\n"
        "####################\n\n"
        "@ Operations : \n"
        f"\t@@ {operation}| : Data Domain = CLASS; Show Latest = Yes; \n\n"
        "@ Parameters Group : Bin\n"
        f"\t@@ \"0|FUNCTIONAL_BIN\" : Operations = {operation}; \n"
        f"\t@@ \"1|INTERFACE_BIN\" : Operations = {operation}; \n\n"
        "####################\n"
        "@ Report Format\n"
        "####################\n\n"
        "@ Data Tracking Level : Units\n"
        "@ Format : Split\n"
        "@ Report Output Format : Csv\n"
        "@ Pivot SSID : YES\n"
        "@ Stack Session Data : NO\n"
        "@ Drop Original Unparsed Column : NO\n"
    )


def _build_local_unc_path(path):
    resolved = Path(path).resolve()
    drive = resolved.drive.rstrip(":")
    if not drive:
        raise ValueError(f"Expected a drive-qualified Windows path, received: {resolved}")
    computer_name = os.environ.get("COMPUTERNAME")
    if not computer_name:
        raise ValueError("COMPUTERNAME environment variable is required to build the AQUA UNC path.")
    tail = str(resolved).replace(resolved.drive, "", 1).lstrip("\\/")
    return f"\\\\{computer_name}\\{drive.lower()}$\\{tail.replace('/', '\\')}"


def _create_session(api_key, verify_ssl=DEFAULT_AQUA_CA_BUNDLE):
    session = requests.Session()
    session.trust_env = False
    if verify_ssl is False:
        raise ValueError("TLS certificate verification cannot be disabled for AQUA requests.")

    if isinstance(verify_ssl, Path):
        if not verify_ssl.exists():
            raise FileNotFoundError(f"AQUA CA bundle not found: {verify_ssl}")
        session.verify = str(verify_ssl)
    else:
        session.verify = verify_ssl
    session.headers.update({"X-Api-Key": api_key})
    return session


def _parse_job_id(response):
    raw_text = response.text.strip()
    if raw_text.isdigit():
        return int(raw_text)

    try:
        payload = response.json()
    except json.JSONDecodeError as exc:
        raise AquaExecutionError(f"Unable to parse AQUA job id from response: {raw_text[:500]}") from exc

    for key in ["jobId", "job_id", "id"]:
        value = payload.get(key)
        if value is not None and str(value).isdigit():
            return int(value)

    raise AquaExecutionError(f"Unable to find AQUA job id in response: {payload}")


def _post_aqua_multipart(session, url, config_name, config_text, output_path=None, timeout=120):
    files = {
        "file": (config_name, config_text.encode("utf-8"), "text/plain"),
    }
    data = {"emailNotification": "no"}
    if output_path:
        data["outputpath"] = output_path

    response = session.post(url, files=files, data=data, timeout=timeout)
    if response.status_code >= 400:
        raise AquaExecutionError(f"AQUA request failed with HTTP {response.status_code}: {response.text[:1000]}")
    return _parse_job_id(response)


def _poll_job(session, base_url, job_id, timeout_seconds=DEFAULT_TIMEOUT_SECONDS, poll_seconds=DEFAULT_POLL_SECONDS):
    deadline = time.time() + timeout_seconds
    status_url = f"{base_url}/api/job/{job_id}/status/full"
    transient_network_retries = 0
    transient_http_retries = 0
    max_transient_retries = 5

    while time.time() < deadline:
        try:
            response = session.get(status_url, timeout=60)
            if response.status_code >= 500:
                transient_http_retries += 1
                if transient_http_retries <= max_transient_retries:
                    time.sleep(poll_seconds)
                    continue
            # AQUA status endpoint can intermittently return 400/404/429 even for valid jobs.
            # Retry these as transient polling errors, but still fail fast on auth/permission errors.
            if response.status_code in {400, 404, 429}:
                transient_http_retries += 1
                if transient_http_retries <= max_transient_retries:
                    time.sleep(poll_seconds)
                    continue
            response.raise_for_status()
        except (requests.exceptions.ConnectionError, requests.exceptions.Timeout):
            transient_network_retries += 1
            if transient_network_retries <= max_transient_retries:
                time.sleep(poll_seconds)
                continue
            raise

        payload = response.json()
        status = str(payload.get("status") or "").strip().lower()
        if status in {"completed", "complete", "success", "succeeded"}:
            return payload
        if status in {"fail", "failed", "error", "aborted", "killed"}:
            return payload
        time.sleep(poll_seconds)

    raise TimeoutError(f"AQUA job {job_id} did not finish within {timeout_seconds} seconds.")


def _extract_server_zip_path(message):
    if not message:
        return None
    match = re.search(r"(\\\\[^\s]+\.zip)", str(message))
    return match.group(1) if match else None


def _find_zip_file(folder):
    zip_files = sorted(Path(folder).glob("*.zip"))
    return zip_files[0] if zip_files else None


def _extract_zip_to_folder(zip_path, target_folder):
    target = Path(target_folder)
    target.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(zip_path, "r") as archive:
        archive.extractall(target)
    text_files = list(target.rglob("*.txt"))
    if not text_files:
        raise AquaExecutionError(f"No text file found inside AQUA ZIP: {zip_path}")
    return text_files[0]


def _load_result_table(text_path, operation=DEFAULT_CLASS_OPERATION):
    df = pd.read_csv(text_path, sep="\t")
    filtered = df.loc[df["Operation"].astype(str).eq(str(operation))].copy()
    selected_columns = [column for column in PREVIEW_COLUMNS if column in filtered.columns]
    if not selected_columns:
        raise AquaExecutionError("AQUA result did not contain the expected preview columns.")

    preview = filtered[selected_columns].sort_values(
        [column for column in ["Visual_ID", "Sort_Lot", "Sort_Wafer_ID", "Sort_X", "Sort_Y"] if column in selected_columns]
    )
    return preview


def fetch_aqua_class_7820_from_vid_csv(
    vid_csv_path,
    credentials_path=DEFAULT_AQUA_CREDENTIALS_PATH,
    output_path=DEFAULT_AQUA_OUTPUT_PATH,
    vid_limit=None,
    operation=DEFAULT_CLASS_OPERATION,
    timeout_seconds=DEFAULT_TIMEOUT_SECONDS,
    verify_ssl=DEFAULT_AQUA_CA_BUNDLE,
):
    vids = _load_true_vids(vid_csv_path, limit=vid_limit)
    credentials = _load_aqua_credentials(credentials_path)
    config_text = build_aqua_class_report_config(vids, operation=operation)
    base_url = credentials["base_url"]

    output_workbook = Path(output_path)
    output_workbook.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory(prefix="aqua_class_", dir=str(output_workbook.parent.resolve())) as temp_dir:
        temp_root = Path(temp_dir)
        unc_output_dir = _build_local_unc_path(temp_root)
        session = _create_session(credentials["api_key"], verify_ssl=verify_ssl)

        generate_url = f"{base_url}/api/miniquery/UnitsList/generate/user/{credentials['user_id']}"
        execute_url = f"{base_url}/api/miniquery/UnitsList/execute/user/{credentials['user_id']}"

        generate_job_id = _post_aqua_multipart(
            session=session,
            url=generate_url,
            config_name="aqua_vid_class.txt",
            config_text=config_text,
        )
        generate_status = _poll_job(
            session=session,
            base_url=base_url,
            job_id=generate_job_id,
            timeout_seconds=timeout_seconds,
        )
        if str(generate_status.get("status") or "").lower() not in {"completed", "complete", "success", "succeeded"}:
            raise AquaExecutionError(f"AQUA generate failed: {generate_status}")

        execute_job_id = _post_aqua_multipart(
            session=session,
            url=execute_url,
            config_name="aqua_vid_class.txt",
            config_text=config_text,
            output_path=unc_output_dir,
        )
        execute_status = _poll_job(
            session=session,
            base_url=base_url,
            job_id=execute_job_id,
            timeout_seconds=timeout_seconds,
        )

        zip_path = None
        local_zip = _find_zip_file(temp_root)
        if local_zip is not None:
            zip_path = local_zip
        else:
            server_zip_path = _extract_server_zip_path(execute_status.get("message"))
            if server_zip_path:
                zip_path = temp_root / Path(server_zip_path).name
                shutil.copy2(server_zip_path, zip_path)

        if zip_path is None:
            raise AquaExecutionError(f"Unable to locate AQUA ZIP output. Final execute status: {execute_status}")

        extracted_text = _extract_zip_to_folder(zip_path, temp_root / "extracted")
        preview_df = _load_result_table(extracted_text, operation=operation)
        with pd.ExcelWriter(output_workbook, engine="openpyxl") as writer:
            preview_df.to_excel(writer, index=False, sheet_name=AQUA_RESULT_SHEET)

        return {
            "status": "completed",
            "environment": credentials["environment"],
            "user_id": credentials["user_id"],
            "vid_count": len(vids),
            "vid_sample": vids[:5],
            "operation": str(operation),
            "generate_job_id": generate_job_id,
            "execute_job_id": execute_job_id,
            "output_path": str(output_workbook),
            "row_count": int(len(preview_df)),
            "result_df": preview_df,
            "sheet_name": AQUA_RESULT_SHEET,
        }


def fetch_aqua_class_7820_from_workbook_vid_sheet(
    final_workbook_path,
    vid_sheet_name="VID",
    credentials_path=DEFAULT_AQUA_CREDENTIALS_PATH,
    vid_limit=None,
    operation=DEFAULT_CLASS_OPERATION,
    timeout_seconds=DEFAULT_TIMEOUT_SECONDS,
    verify_ssl=DEFAULT_AQUA_CA_BUNDLE,
):
    vids = _load_true_vids_from_workbook(final_workbook_path, sheet_name=vid_sheet_name, limit=vid_limit)
    credentials = _load_aqua_credentials(credentials_path)
    config_text = build_aqua_class_report_config(vids, operation=operation)
    base_url = credentials["base_url"]

    final_workbook = Path(final_workbook_path)
    with tempfile.TemporaryDirectory(prefix="aqua_class_", dir=str(final_workbook.parent.resolve())) as temp_dir:
        temp_root = Path(temp_dir)
        unc_output_dir = _build_local_unc_path(temp_root)
        session = _create_session(credentials["api_key"], verify_ssl=verify_ssl)

        generate_url = f"{base_url}/api/miniquery/UnitsList/generate/user/{credentials['user_id']}"
        execute_url = f"{base_url}/api/miniquery/UnitsList/execute/user/{credentials['user_id']}"

        generate_job_id = _post_aqua_multipart(session=session, url=generate_url, config_name="aqua_vid_class.txt", config_text=config_text)
        generate_status = _poll_job(session=session, base_url=base_url, job_id=generate_job_id, timeout_seconds=timeout_seconds)
        if str(generate_status.get("status") or "").lower() not in {"completed", "complete", "success", "succeeded"}:
            raise AquaExecutionError(f"AQUA generate failed: {generate_status}")

        execute_job_id = _post_aqua_multipart(session=session, url=execute_url, config_name="aqua_vid_class.txt", config_text=config_text, output_path=unc_output_dir)
        execute_status = _poll_job(session=session, base_url=base_url, job_id=execute_job_id, timeout_seconds=timeout_seconds)

        zip_path = _find_zip_file(temp_root)
        if zip_path is None:
            server_zip_path = _extract_server_zip_path(execute_status.get("message"))
            if server_zip_path:
                zip_path = temp_root / Path(server_zip_path).name
                shutil.copy2(server_zip_path, zip_path)
        if zip_path is None:
            raise AquaExecutionError(f"Unable to locate AQUA ZIP output. Final execute status: {execute_status}")

        extracted_text = _extract_zip_to_folder(zip_path, temp_root / "extracted")
        preview_df = _load_result_table(extracted_text, operation=operation)
        return {
            "status": "completed",
            "environment": credentials["environment"],
            "user_id": credentials["user_id"],
            "vid_count": len(vids),
            "vid_sample": vids[:5],
            "operation": str(operation),
            "generate_job_id": generate_job_id,
            "execute_job_id": execute_job_id,
            "row_count": int(len(preview_df)),
            "result_df": preview_df,
            "sheet_name": AQUA_RESULT_SHEET,
        }