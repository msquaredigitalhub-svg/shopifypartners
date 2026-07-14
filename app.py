"""
Flask web app for the Shopify Partner Directory Scraper.

Run with:
    pip install flask requests beautifulsoup4 openpyxl
    python app.py

Then open http://localhost:5000
"""

import json
import math
import os
import threading
from functools import wraps

from flask import Flask, Response, jsonify, render_template, request, send_file, session, redirect

from shopify_partner_scraper import (
    COUNTRIES,
    INDUSTRIES,
    LANGUAGES,
    OUTPUT_FILE,
    PARTNER_TIERS,
    PARTNERS_PER_PAGE,
    check_listing_count,
    get_download_path,
    cleanup_download,
    run_scrape,
)

app = Flask(__name__)
app.secret_key = os.urandom(32).hex()

PASSWORD = "Shopify@123$"
SESSION_KEY = "authenticated"

_scraping_lock = threading.Lock()


# ---------------------------------------------------------------------------
# Authentication helpers
# ---------------------------------------------------------------------------

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get(SESSION_KEY):
            return redirect("/login")
        return f(*args, **kwargs)
    return decorated


# ---------------------------------------------------------------------------
# Login / Logout
# ---------------------------------------------------------------------------

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "GET":
        return render_template("login.html", error=None)

    password = request.form.get("password", "")
    if password == PASSWORD:
        session[SESSION_KEY] = True
        return redirect("/")

    return render_template("login.html", error="Incorrect password."), 401


@app.route("/logout")
def logout():
    session.pop(SESSION_KEY, None)
    return redirect("/login")


# ---------------------------------------------------------------------------
# API: filter options
# ---------------------------------------------------------------------------

@app.route("/api/filters")
@login_required
def api_filters():
    return jsonify({
        "industries": [{"handle": k, "name": v["name"]} for k, v in INDUSTRIES.items()],
        "tiers": [{"handle": k, "name": v["name"]} for k, v in PARTNER_TIERS.items()],
        "languages": [{"code": k, "name": v["name"]} for k, v in LANGUAGES.items()],
        "countries": [{"code": k, "name": v["name"]} for k, v in COUNTRIES.items()],
    })


# ---------------------------------------------------------------------------
# Extract filters from JSON body
# ---------------------------------------------------------------------------

def _parse_filters(data: dict) -> dict:
    return {
        "country_codes": data.get("countries", []),
        "industry": data.get("industry", ""),
        "partner_tiers": data.get("partnerTiers", []),
        "language": data.get("language", ""),
    }


# ---------------------------------------------------------------------------
# API: check listing count for filters
# ---------------------------------------------------------------------------

@app.route("/api/check-count", methods=["POST"])
@login_required
def api_check_count():
    data = request.get_json(silent=True) or {}
    country_codes = data.get("countries", [])

    if not isinstance(country_codes, list):
        return jsonify({"error": "countries must be a list"}), 400

    for cc in country_codes:
        if cc not in COUNTRIES:
            return jsonify({"error": f"Invalid country code: {cc}"}), 400

    filters = _parse_filters(data)
    total = check_listing_count(**filters)

    if total == 0:
        return jsonify({"error": "Could not retrieve listing count."}), 502

    estimated_pages = max(1, math.ceil(total / PARTNERS_PER_PAGE))

    return jsonify({
        "total": total,
        "estimated_pages": estimated_pages,
    })


# ---------------------------------------------------------------------------
# API: run scrape (Server-Sent Events)
# ---------------------------------------------------------------------------

@app.route("/api/scrape", methods=["POST"])
@login_required
def api_scrape():
    data = request.get_json(silent=True) or {}
    country_codes = data.get("countries", [])
    limit = data.get("limit")

    if not isinstance(country_codes, list):
        return jsonify({"error": "countries must be a list"}), 400

    for cc in country_codes:
        if cc not in COUNTRIES:
            return jsonify({"error": f"Invalid country code: {cc}"}), 400

    if limit is not None:
        try:
            limit = int(limit)
        except (ValueError, TypeError):
            return jsonify({"error": "Limit must be a number"}), 400

    if not _scraping_lock.acquire(blocking=False):
        return jsonify({"error": "A scrape is already in progress. Please wait for it to finish."}), 409

    filters = _parse_filters(data)

    def generate():
        try:
            output_path = OUTPUT_FILE
            for event in run_scrape(output_path=output_path, limit=limit, **filters):
                yield f"data: {json.dumps(event)}\n\n".encode()
        finally:
            _scraping_lock.release()

    resp = Response(generate(), mimetype="text/event-stream", headers={
        "Cache-Control": "no-cache",
        "X-Accel-Buffering": "no",
    })
    resp.direct_passthrough = True
    return resp


# ---------------------------------------------------------------------------
# API: download the completed Excel file
# ---------------------------------------------------------------------------

@app.route("/api/download/<download_id>")
@login_required
def api_download(download_id):
    filepath = get_download_path(download_id)
    if not filepath or not os.path.exists(filepath):
        return jsonify({"error": "Download not found or expired."}), 404

    response = send_file(
        filepath,
        as_attachment=True,
        download_name="shopify_partners.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    cleanup_download(download_id)
    return response


# ---------------------------------------------------------------------------
# Web page
# ---------------------------------------------------------------------------

@app.route("/")
@login_required
def index():
    return render_template("index.html",
                           countries=COUNTRIES,
                           industries=INDUSTRIES,
                           tiers=PARTNER_TIERS,
                           languages=LANGUAGES)


if __name__ == "__main__":
    print("Starting Shopify Partners Extractor web interface...")
    print("Open http://localhost:5000 in your browser.")
    app.run(debug=True, threaded=True)
