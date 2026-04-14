"""Flask web server for the mobile/browser overlay."""

import logging
import socket
from dataclasses import asdict

from flask import Flask, jsonify, render_template, request

from scanning_tool.state_context import app_state
from scanning_tool.config import resource_path
## Removed import of MULTIPLIER_CODES (now replaced by dynamic scan signature data)

logger = logging.getLogger("scanning_tool")


def get_local_ip() -> str:
    """Best-effort detection of the primary local network IP address."""
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            ip_address = sock.getsockname()[0]
            if ip_address:
                return ip_address
    except Exception as exc:
        logger.debug(f"Unable to determine local IP automatically: {exc}")
    return "127.0.0.1"


def create_app() -> Flask:
    """Create and configure the Flask application."""
    template_folder = resource_path("templates")
    app = Flask(__name__, template_folder=template_folder)

    @app.route("/")
    def index():
        return render_template("overlay.html")

    @app.route("/status")
    def status():
        """Return the latest scan information for the overlay UI."""
        selected_region = request.args.get("region", "STANTON").upper()
        result = app_state.scan_state.last_result
        info = getattr(result, "info", None)

        table = None
        if info:
            deposit_key = (info.get("key") or info.get("name") or "").upper()
            region_tables = app_state.service_state.deposit_tables.get(selected_region, {})
            table = region_tables.get(deposit_key)
            category = str(info.get("category", "")).lower()
            if not table or category not in {"rock deposits", "gems"}:
                table = None

        response = {
            "region": app_state.settings.capture.cap_region,
            "label_color": app_state.settings.overlay.label_color,
            "last": asdict(result),
            "alignment": asdict(app_state.scan_state.last_alignment_info),
            "selected_region": selected_region,
            "info": info,
            "code": result.code,
            "code_raw": result.code_raw,
            "confidence": float(result.confidence),
            "raw_text": result.raw_text,
            "table": table,
        }

        return jsonify(response)

    return app
