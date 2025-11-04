"""
drive_uploader.py
~~~~~~~~~~~~~~~~~~

This module provides a simple helper to upload files to Google Drive using
service account credentials. The service account JSON and destination folder ID
should be supplied via environment variables `GDRIVE_SERVICE_ACCOUNT_JSON`
and `GDRIVE_FOLDER_ID` respectively. If either is missing, uploads will be
skipped and a warning will be logged.

The `upload_files` function accepts a list of `(filepath, dest_name)` tuples and
uploads each file into the specified Drive folder. On success, a mapping from
file names to their Drive file IDs and web links is returned. If an upload
fails, the error is logged and processing continues for subsequent files.

Note: The Google Drive API has network side effects. When running in
environments without internet access or secrets configured, the upload will
silently skip and return an empty mapping.
"""

from __future__ import annotations

import json
import logging
import os
from pathlib import Path
from typing import Iterable, List, Tuple, Dict

try:
    from google.oauth2 import service_account
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
except ImportError:
    # If google-api-python-client is not installed, uploads will be skipped.
    service_account = None  # type: ignore
    build = None  # type: ignore
    MediaFileUpload = None  # type: ignore


def _get_drive_service() -> tuple[object | None, str | None]:
    """Create and return an authenticated Drive service or (None, error message).

    The function reads service account JSON from the `GDRIVE_SERVICE_ACCOUNT_JSON`
    environment variable. If not provided or invalid, returns (None, error msg).
    """
    json_str = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")
    folder_id = os.environ.get("GDRIVE_FOLDER_ID")
    if not json_str or not folder_id:
        return None, "Google Drive credentials or folder ID not configured"
    if service_account is None or build is None:
        return None, "google-api-python-client is not installed"
    try:
        info = json.loads(json_str)
        creds = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive.file"]
        )
        service = build("drive", "v3", credentials=creds)
        return service, None
    except Exception as exc:
        return None, f"Failed to create Drive service: {exc}"


def upload_files(files: Iterable[Tuple[str, str]]) -> Dict[str, Dict[str, str]]:
    """Upload multiple files to Google Drive.

    Args:
        files: An iterable of (local_path, dest_name) tuples. `local_path` is
            the path to the file on the local filesystem. `dest_name` is the
            desired name of the file in Drive. If dest_name already exists,
            Drive will create a new version automatically.

    Returns:
        A dict mapping dest_name to a dict with keys 'id' and 'webViewLink'.
        Only files that successfully uploaded will appear in the result.
    """
    logger = logging.getLogger(__name__)
    service, err = _get_drive_service()
    results: Dict[str, Dict[str, str]] = {}
    if service is None:
        logger.warning(f"Skipping Drive upload: {err}")
        return results
    folder_id = os.environ.get("GDRIVE_FOLDER_ID")
    for local_path, dest_name in files:
        try:
            media = MediaFileUpload(local_path, resumable=True)
            file_metadata = {"name": dest_name, "parents": [folder_id]}
            request = service.files().create(
                body=file_metadata, media_body=media, fields="id, webViewLink"
            )
            response = request.execute()
            file_id = response.get("id")
            link = response.get("webViewLink")
            results[dest_name] = {"id": file_id, "webViewLink": link}
            logger.info(f"Uploaded {dest_name} to Drive (ID: {file_id})")
        except Exception as exc:
            logger.error(f"Failed to upload {dest_name}: {exc}")
    return results


__all__ = ["upload_files"]