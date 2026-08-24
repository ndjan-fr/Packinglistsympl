#!/usr/bin/env python3
"""
Packing List Automation — sympl.fr
Runs hourly via GitHub Actions. No browser needed.

Required environment variables (GitHub Secrets):
  NETLIFY_TOKEN        — Netlify personal access token
  NETLIFY_SITE_ID      — f9550cb4-bc9c-4105-a009-f54324835a11
  CALENDLY_TOKEN       — Calendly personal access token
  SYMPL_EMAIL          — mdjan@sympl.fr
  SYMPL_PASSWORD       — mD411062++!!
  SHEETS_WEBHOOK_URL   — Google Apps Script webhook URL
"""

import os
import sys
import json
import logging
from datetime import datetime, timezone, timedelta

import requests
from openpyxl import Workbook

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ── Config ─────────────────────────────────────────────────────────────────────
NETLIFY_TOKEN   = os.environ["NETLIFY_TOKEN"]
NETLIFY_SITE_ID = os.environ.get("NETLIFY_SITE_ID", "f9550cb4-bc9c-4105-a009-f54324835a11")
NETLIFY_FORM    = "packing-list"
CALENDLY_TOKEN  = os.environ["CALENDLY_TOKEN"]
SYMPL_EMAIL     = os.environ["SYMPL_EMAIL"]
SYMPL_PASSWORD  = os.environ["SYMPL_PASSWORD"]
SYMPL_BASE      = "https://live.sympl.fr"
WINDOW_HOURS    = 24
SHEETS_WEBHOOK_URL = os.environ.get("SHEETS_WEBHOOK_URL", "")
SHEETS_SECRET      = "sympl-packing-list-2024"

# Classique template column headers (required by sympl.fr)
CLASSIQUE_HEADERS = [
    "Code-barres", "Quantité", "Nom", "Reference", "Couleur", "Taille",
    "Type de code barre", "Assurance", "Composition", "Valeur",
    "Pays d'origine", "Code SH", "Imei", "Lot",
    "Date d'expiration (au format JJ/MM/AAAA)", "Quantité buffer",
]

# Company name → companyId lookup table (partial — extended as needed)
COMPANY_TABLE = {
    "23 HEURES 59 EDITIONS": 1284, "AGAIN": 1911, "AKOR": 1275, "ASPHALTE": 358,
    "ATELIER MATERI": 1760, "BIBI": 1859, "BTLC": 1747, "ECOJOKO": 1615,
    "FEMPO": 1829, "FORLIFE": 1639, "FRENCH PALS": 1814, "GEEK STORE": 481,
    "KALIOS": 1838, "KUMIKO MATCHA": 1213, "LE DOSE CLUB": 1864,
    "MAISON MATINE": 1830, "PANAME COLLECTIONS": 307, "RED ART GAMES": 1576,
    "SAEVE": 1388, "SHATELY": 396, "SISTERS REPUBLIC": 1544, "SYMPL": 1,
    "THE7CGROUP": 1999, "THE ENTHUSIASTS": 1813, "URBAN DIET": 1378,
    "VERRE&IMAGE": 63,
    # Add more as needed — full list in the skill SKILL.md
}


# ── Netlify ─────────────────────────────────────────────────────────────────────
def netlify_headers():
    return {"Authorization": f"Bearer {NETLIFY_TOKEN}"}

def get_form_id():
    r = requests.get(
        f"https://api.netlify.com/api/v1/sites/{NETLIFY_SITE_ID}/forms",
        headers=netlify_headers(), timeout=30
    )
    r.raise_for_status()
    for form in r.json():
        if form["name"] == NETLIFY_FORM:
            return form["id"]
    raise ValueError(f"Form '{NETLIFY_FORM}' not found")

def get_recent_submissions(form_id):
    r = requests.get(
        f"https://api.netlify.com/api/v1/forms/{form_id}/submissions",
        headers=netlify_headers(), params={"per_page": 50}, timeout=30
    )
    r.raise_for_status()
    cutoff = datetime.now(timezone.utc) - timedelta(hours=WINDOW_HOURS)
    recent = []
    for sub in r.json():
        created = datetime.fromisoformat(sub["created_at"].replace("Z", "+00:00"))
        if created >= cutoff:
            recent.append(sub)
    log.info(f"Netlify: {len(r.json())} total submissions, {len(recent)} within {WINDOW_HOURS}h")
    return recent

def delete_submission(submission_id):
    r = requests.delete(
        f"https://api.netlify.com/api/v1/submissions/{submission_id}",
        headers=netlify_headers(), timeout=30
    )
    r.raise_for_status()
    log.info(f"Deleted Netlify submission {submission_id}")


# ── Calendly ────────────────────────────────────────────────────────────────────
def get_calendly_user():
    r = requests.get(
        "https://api.calendly.com/users/me",
        headers={"Authorization": f"Bearer {CALENDLY_TOKEN}"}, timeout=30
    )
    r.raise_for_status()
    return r.json()["resource"]["uri"]

def find_calendly_event(email, reference):
    """Find upcoming Réception de marchandises event matching email or reference."""
    user_uri = get_calendly_user()
    now = datetime.now(timezone.utc).isoformat()
    r = requests.get(
        "https://api.calendly.com/scheduled_events",
        headers={"Authorization": f"Bearer {CALENDLY_TOKEN}"},
        params={
            "user": user_uri,
            "min_start_time": now,
            "status": "active",
            "count": 100,
            "sort": "start_time:asc",
        },
        timeout=30
    )
    r.raise_for_status()
    events = r.json().get("collection", [])
    for event in events:
        if "livraison" not in event.get("name", "").lower():
            continue
        # Fetch invitees to match by email
        inv_r = requests.get(
            f"https://api.calendly.com/scheduled_events/{event['uri'].split('/')[-1]}/invitees",
            headers={"Authorization": f"Bearer {CALENDLY_TOKEN}"}, timeout=30
        )
        if inv_r.status_code != 200:
            continue
        for invitee in inv_r.json().get("collection", []):
            q_text = " ".join(
                str(a.get("value", "")) for a in invitee.get("questions_and_answers", [])
            )
            if invitee.get("email", "").lower() == email.lower() or reference in q_text:
                start = event["start_time"]
                dt = datetime.fromisoformat(start.replace("Z", "+00:00"))
                return dt.strftime("%d/%m/%Y")
    return None


# ── sympl.fr session ────────────────────────────────────────────────────────────
def sympl_login():
    s = requests.Session()
    s.headers.update({"User-Agent": "Mozilla/5.0 (compatible; packing-list-bot/1.0)"})
    # Get login page to get CSRF token
    r = s.get(f"{SYMPL_BASE}/login", timeout=30)
    r.raise_for_status()
    from html.parser import HTMLParser
    class TokenParser(HTMLParser):
        token = None
        def handle_starttag(self, tag, attrs):
            if tag == "input":
                d = dict(attrs)
                if d.get("name") == "_token":
                    self.token = d.get("value")
    p = TokenParser()
    p.feed(r.text)
    if not p.token:
        raise ValueError("Could not extract CSRF token from login page")
    # Login
    r2 = s.post(f"{SYMPL_BASE}/login", data={
        "_token": p.token,
        "email": SYMPL_EMAIL,
        "password": SYMPL_PASSWORD,
    }, allow_redirects=True, timeout=30)
    if "/login" in r2.url:
        raise ValueError("sympl.fr login failed — check credentials")
    log.info("sympl.fr login successful")
    return s

def check_reception_exists(session, company_id, reference):
    r = session.get(
        f"{SYMPL_BASE}/admin/stock/receptions",
        params={"reference": reference}, timeout=30
    )
    return "Aucune réception n'a été trouvée" not in r.text

def get_company_id(company_name, email_domain):
    name_upper = company_name.upper().strip()
    # Direct match
    if name_upper in COMPANY_TABLE:
        return COMPANY_TABLE[name_upper]
    # Partial match
    for k, v in COMPANY_TABLE.items():
        if k in name_upper or name_upper in k:
            return v
    # Domain match
    for k, v in COMPANY_TABLE.items():
        if email_domain.lower() in k.lower():
            return v
    return None

def create_placeholder_xlsx():
    wb = Workbook()
    ws = wb.active
    ws.title = "Feuil1"
    ws.append(CLASSIQUE_HEADERS)
    ws.append([""] * len(CLASSIQUE_HEADERS))
    import tempfile, os
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    wb.save(tmp.name)
    tmp.close()
    return tmp.name

def create_reception(session, company_id, data, delivery_date, packing_list_url):
    # Get create page for CSRF token
    r = session.get(
        f"{SYMPL_BASE}/admin/stock/receptions/create",
        params={"companyId": company_id}, timeout=30
    )
    r.raise_for_status()
    from html.parser import HTMLParser
    class TokenParser(HTMLParser):
        token = None
        def handle_starttag(self, tag, attrs):
            if tag == "input":
                d = dict(attrs)
                if d.get("name") == "_token":
                    self.token = d.get("value")
    p = TokenParser()
    p.feed(r.text)
    if not p.token:
        raise ValueError("Could not extract CSRF token from reception create page")

    # Build placeholder xlsx file
    xlsx_path = create_placeholder_xlsx()

    description = data.get("commentaires") or (
        f"Soumis via portail client par {data.get('contact')} ({data.get('email')}). "
        f"Fichier PL: {packing_list_url}"
    )

    form_data = {
        "_token": p.token,
        "company_id": str(company_id),
        "location_id": "4",  # MARLY1
        "reference": data.get("reference", ""),
        "description": description,
        "carrier_name": data.get("transporteur", ""),
        "tracking_number": data.get("numero_suivi", ""),
        "expected_delivery_date": delivery_date,
        "expected_number_of_pallets": str(data.get("nombre_palettes", 1)),
        "expected_number_of_packages_in_pallet": str(data.get("cartons_palette", 0)),
        "expected_number_of_packages_in_bulk": str(data.get("cartons_vrac", 0)),
        "packing_list_extractor_code": "SYMPL",
    }

    with open(xlsx_path, "rb") as f:
        files = {"packing_list": ("packing_list.xlsx", f,
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")}
        r2 = session.post(
            f"{SYMPL_BASE}/admin/stock/receptions/create",
            params={"companyId": company_id},
            data=form_data,
            files=files,
            allow_redirects=True,
            timeout=60
        )

    os.unlink(xlsx_path)

    if "réception a bien été créée" in r2.text or "/receptions/" in r2.url:
        reception_id = r2.url.split("/receptions/")[-1].split("?")[0] if "/receptions/" in r2.url else "?"
        log.info(f"Reception created: #{reception_id}")
        return reception_id
    else:
        # Check for error
        from html.parser import HTMLParser
        class ErrParser(HTMLParser):
            errors = []
            def handle_starttag(self, tag, attrs):
                d = dict(attrs)
                if "alert" in d.get("class",""):
                    self._in_alert = True
            def handle_data(self, data):
                if data.strip():
                    self.errors.append(data.strip())
        ep = ErrParser()
        ep.feed(r2.text)
        raise ValueError(f"Reception creation failed. URL: {r2.url}. Errors: {ep.errors[:3]}")


# ── Google Sheets logging ───────────────────────────────────────────────────────
def log_to_sheets(data, reception_id):
    """Send a row to the Google Sheets webhook (Apps Script doPost)."""
    if not SHEETS_WEBHOOK_URL:
        log.warning("SHEETS_WEBHOOK_URL not set — skipping Sheets logging")
        return
    payload = {
        "secret": SHEETS_SECRET,
        "date_soumission":      datetime.now(timezone.utc).strftime("%d/%m/%Y %H:%M"),
        "societe":              data.get("societe", ""),
        "contact":              data.get("contact", ""),
        "email":                data.get("email", ""),
        "telephone":            data.get("telephone", ""),
        "reference":            data.get("reference", ""),
        "transporteur":         data.get("transporteur", ""),
        "numero_suivi":         data.get("numero_suivi", ""),
        "nombre_palettes":      data.get("nombre_palettes", ""),
        "cartons_palette":      data.get("cartons_palette", ""),
        "cartons_vrac":         data.get("cartons_vrac", ""),
        "format":               data.get("format", "classique"),
        "commentaires":         data.get("commentaires", ""),
        "fichier_packing_list": data.get("fichier_packing_list", ""),
        "reception_id":         str(reception_id),
    }
    try:
        r = requests.post(SHEETS_WEBHOOK_URL, json=payload, timeout=15)
        result = r.json()
        if result.get("status") == "ok":
            log.info(f"Logged to Google Sheets (reception #{reception_id})")
        else:
            log.warning(f"Sheets webhook returned: {result}")
    except Exception as e:
        log.warning(f"Could not log to Google Sheets: {e}")


# ── Main ────────────────────────────────────────────────────────────────────────
def main():
    log.info("=== Packing list automation starting ===")

    # 1. Get Netlify submissions
    form_id = get_form_id()
    submissions = get_recent_submissions(form_id)

    if not submissions:
        log.info("No new submissions in the last 24h. Nothing to do.")
        return

    # 2. Login to sympl.fr once
    sympl = sympl_login()

    processed = 0
    for sub in submissions:
        d = sub.get("data", {})
        company_name = d.get("societe", "")
        reference    = d.get("reference", "")
        email        = d.get("email", "")
        email_domain = email.split("@")[-1].split(".")[0] if "@" in email else ""

        log.info(f"Processing: {company_name} — ref {reference}")

        # 3. Anti-duplicate check
        company_id = get_company_id(company_name, email_domain)
        if company_id is None:
            log.error(f"Company not found for '{company_name}' ({email}). Manual action required.")
            continue

        if check_reception_exists(sympl, company_id, reference):
            log.warning(f"Reception already exists for {company_name} ref {reference}. Skipping.")
            continue

        # 4. Get Calendly delivery date
        delivery_date = find_calendly_event(email, reference)
        if delivery_date:
            log.info(f"Calendly date found: {delivery_date}")
        else:
            # Fallback: use form date
            delivery_date = d.get("date_livraison", "")
            log.warning(f"No Calendly event found for {email}. Using form date: {delivery_date}")

        # 5. Create reception
        try:
            reception_id = create_reception(
                sympl, company_id, d, delivery_date,
                d.get("fichier_packing_list", "")
            )
            log.info(f"Reception #{reception_id} created for {company_name} ref {reference}")

            # 6. Log to Google Sheets
            log_to_sheets(d, reception_id)

            # 7. Delete Netlify submission
            delete_submission(sub["id"])
            processed += 1

        except Exception as e:
            log.error(f"Failed to create reception for {company_name}: {e}")

    log.info(f"=== Done: {processed}/{len(submissions)} submissions processed ===")


if __name__ == "__main__":
    main()
