"""Download a small, openly licensed, real-world .docx smoke corpus.

The source list is intentionally curated from publishers' own pages. It is not
an exhaustive or random sample. Every fetched file is checked for the OOXML ZIP
signature before it is retained.
"""
from __future__ import annotations

import csv
import hashlib
import json
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FILES = ROOT / "files"

SOURCES = [
    {
        "id": "hmt_wga_annex_2023_24",
        "filename": "hmt_wga_annex_written_guidance_2024.docx",
        "title": "Annex - Written Guidance update Sept 2024",
        "publisher": "HM Treasury",
        "source_page_first_published": "2024-06-17",
        "source_page_last_updated": "2025-04-08",
        "source_page": "https://www.gov.uk/government/publications/whole-of-government-accounts-2023-to-2024-guidance-for-preparers",
        "download_url": "https://assets.publishing.service.gov.uk/media/66e80929f8082e9740881b40/Annex_Written_guidance_-_updates_September_2024.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "long guidance / narrative / tables",
    },
    {
        "id": "hmt_wga_2024_25_cg02",
        "filename": "hmt_wga_2024_25_cg02.docx",
        "title": "CG-02 template for notification of the agreement process with other WGA bodies",
        "publisher": "HM Treasury",
        "source_page_first_published": "2025-05-29",
        "source_page_last_updated": "2026-03-20",
        "source_page": "https://www.gov.uk/government/publications/whole-of-government-accounts-2024-to-2025-guidance-for-preparers",
        "download_url": "https://assets.publishing.service.gov.uk/media/6842ef8e0392ed9b784c0128/WGA_2024-25_form_CG-02_notification_of_the_agreement_process_with_other_WGA_bodies.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "operational form / template",
    },
    {
        "id": "hmt_wga_2024_25_cg03",
        "filename": "hmt_wga_2024_25_cg03.docx",
        "title": "CG-03 template for minor body exemption from completing WGA returns",
        "publisher": "HM Treasury",
        "source_page_first_published": "2025-05-29",
        "source_page_last_updated": "2026-03-20",
        "source_page": "https://www.gov.uk/government/publications/whole-of-government-accounts-2024-to-2025-guidance-for-preparers",
        "download_url": "https://assets.publishing.service.gov.uk/media/6842efaa57f3515d9611f063/WGA_2024-25_CG-03_template_for_minor_body_exemption_from_completing_WGA_returns.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "operational form / template",
    },
    {
        "id": "hmt_wga_2024_25_cg04_central",
        "filename": "hmt_wga_2024_25_cg04_central.docx",
        "title": "CG-04 management review checklist for central government",
        "publisher": "HM Treasury",
        "source_page_first_published": "2025-05-29",
        "source_page_last_updated": "2026-03-20",
        "source_page": "https://www.gov.uk/government/publications/whole-of-government-accounts-2024-to-2025-guidance-for-preparers",
        "download_url": "https://assets.publishing.service.gov.uk/media/6842f06457f3515d9611f065/WGA_2024-25_form_CG-04_management_review_checklist_for_central_government__1_.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "operational checklist / template",
    },
    {
        "id": "hmt_wga_2024_25_cg04_local",
        "filename": "hmt_wga_2024_25_cg04_local.docx",
        "title": "CG-04 management review checklist for local government",
        "publisher": "HM Treasury",
        "source_page_first_published": "2025-05-29",
        "source_page_last_updated": "2026-03-20",
        "source_page": "https://www.gov.uk/government/publications/whole-of-government-accounts-2024-to-2025-guidance-for-preparers",
        "download_url": "https://assets.publishing.service.gov.uk/media/6842f071e5a089417c8060de/WGA_2024-25__form_CG-04_management_review_checklist_for_local_government_.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "operational checklist / template",
    },
    {
        "id": "hmt_wga_2024_25_cg05",
        "filename": "hmt_wga_2024_25_cg05.docx",
        "title": "CG-05 template for the notification of audit completion",
        "publisher": "HM Treasury",
        "source_page_first_published": "2025-05-29",
        "source_page_last_updated": "2026-03-20",
        "source_page": "https://www.gov.uk/government/publications/whole-of-government-accounts-2024-to-2025-guidance-for-preparers",
        "download_url": "https://assets.publishing.service.gov.uk/media/69bd1455b920af63be1c7559/WGA_2024-25__form_CG-05_template_for_the_notification_of_audit_completion.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "audit form / template",
    },
    {
        "id": "dwp_atw_withdrawn_annex_b",
        "filename": "dwp_atw_withdrawn_annex_b.docx",
        "title": "Withdrawn Annex B: Access to Work MHSS Customer Consent To Share Information",
        "publisher": "Department for Work and Pensions",
        "source_page_first_published": "2024-08-20",
        "source_page_last_updated": "2026-08-20",
        "source_page": "https://www.gov.uk/government/publications/access-to-work-mhss-provider-guidance-for-contracts-effective-on-or-after-20-august-2024",
        "download_url": "https://assets.publishing.service.gov.uk/media/66c31ebfb75776507ecdf3af/withdrawn-annex-b-atw-mhss-consent.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "consent form / template",
    },
    {
        "id": "oag_style_guidance_2016",
        "filename": "oag_style_guidance_2016.docx",
        "title": "Style Guidance for Documents",
        "publisher": "Office of the Advocate General for Scotland",
        "source_page_first_published": "2016-04-21",
        "source_page": "https://www.gov.uk/government/publications/style-guidance-for-documents",
        "download_url": "https://assets.publishing.service.gov.uk/media/5a80d7e240f0b62305b8d766/REDACTED_09.02.16_Style__Guidence_for_documents.docx",
        "license": "Open Government Licence v3.0 (GOV.UK terms; subject to third-party exceptions)",
        "kind": "FOI-released prose / document",
    },
]


def fetch(source: dict, force: bool = False, previous: dict | None = None) -> dict:
    path = FILES / source["filename"]
    status = None
    content_type = None
    if path.exists() and not force:
        data = path.read_bytes()
    else:
        request = urllib.request.Request(
            source["download_url"],
            headers={"User-Agent": "docviews real-world OOXML evaluation/0.1"},
        )
        with urllib.request.urlopen(request, timeout=60) as response:
            data = response.read()
            status = response.status
            content_type = response.headers.get("Content-Type", "")
        if status != 200:
            raise RuntimeError(f"{source['id']}: HTTP {status}")
        if not data.startswith(b"PK\x03\x04"):
            raise RuntimeError(
                f"{source['id']}: expected OOXML ZIP, got {content_type!r}; "
                f"signature={data[:12]!r}"
            )
        path.write_bytes(data)
    if not data.startswith(b"PK\x03\x04"):
        raise RuntimeError(f"{source['id']}: local file is not a ZIP-based OOXML package")
    row = dict(source)
    row["govuk_terms_url"] = "https://www.gov.uk/help/terms-conditions"
    row["open_government_licence_url"] = "https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/"
    row["bytes"] = len(data)
    row["sha256"] = hashlib.sha256(data).hexdigest()
    row["local_file"] = path.name
    row["http_status"] = status
    row["response_content_type"] = content_type
    row["fetched_utc"] = datetime.now(timezone.utc).isoformat(timespec="seconds") if status else None
    if status is None and previous and previous.get("sha256") == row["sha256"]:
        for key in ("http_status", "response_content_type", "fetched_utc"):
            row[key] = previous.get(key)
    return row


def main() -> None:
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--force", action="store_true", help="re-download existing files")
    args = parser.parse_args()
    FILES.mkdir(parents=True, exist_ok=True)
    manifest_path = ROOT / "manifest.json"
    previous = json.loads(manifest_path.read_text()) if manifest_path.exists() else []
    previous_by_id = {row.get("id"): row for row in previous}
    manifest = []
    for i, source in enumerate(SOURCES, 1):
        try:
            row = fetch(source, force=args.force, previous=previous_by_id.get(source["id"]))
            manifest.append(row)
            print(f"{i}/{len(SOURCES)} OK {row['filename']} {row['bytes']} bytes sha256={row['sha256'][:12]}")
        except Exception as exc:
            print(f"{i}/{len(SOURCES)} FAILED {source['filename']}: {type(exc).__name__}: {exc}")
            raise
    (ROOT / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
    columns = list(manifest[0]) if manifest else []
    with (ROOT / "manifest.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(manifest)
    print(f"wrote {ROOT / 'manifest.json'}")
    print(f"wrote {ROOT / 'manifest.csv'}")


if __name__ == "__main__":
    main()
