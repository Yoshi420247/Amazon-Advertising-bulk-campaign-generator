"""
Vercel Serverless Function: Generate bulk campaign XLSX.

POST /api/generate
Body (JSON):
  {
    "asins": ["B08NV6CLGF", "B09ABC1234"],
    "competitor_asins": ["B07ABC1234"],        // optional
    "tier_assignment": {"B08NV6CLGF": "365+"}, // optional
    "daily_budget": 50.0,                      // optional
    "default_bid": 0.75,                       // optional
    "bidding_strategy": "Dynamic bids - down only", // optional
    "campaign_prefix": "Clearance"             // optional
  }

Returns: XLSX file download

Requires env vars:
  SUPABASE_URL, SUPABASE_KEY (optional - for persistence)
  LISTINGS_REPORT_PATH (optional - defaults to samples/sample_active_listings.txt)
"""

import io
import json
import os
import sys
import tempfile
from http.server import BaseHTTPRequestHandler
from pathlib import Path

# Add project root to path so we can import the generator
project_root = str(Path(__file__).parent.parent)
if project_root not in sys.path:
    sys.path.insert(0, project_root)

from bulk_campaign_generator import (
    CampaignConfig,
    DEFAULT_COLUMNS,
    BulkCampaignGenerator,
    build_asin_sku_mapping,
    generate_bulk_upload,
    read_active_listings_report,
    read_template_headers,
    write_bulk_xlsx,
)


class handler(BaseHTTPRequestHandler):
    def do_POST(self):
        try:
            # Parse request body
            content_length = int(self.headers.get("Content-Length", 0))
            body = self.rfile.read(content_length)
            data = json.loads(body) if body else {}

            # Validate required fields
            asins = data.get("asins", [])
            if not asins:
                self._error(400, "Missing required field: asins")
                return

            if isinstance(asins, str):
                asins = [a.strip() for a in asins.split(",") if a.strip()]

            # Build config from request
            config_kwargs = {}
            if "daily_budget" in data:
                config_kwargs["daily_budget_365_plus"] = float(data["daily_budget"])
                config_kwargs["daily_budget_211_330"] = float(data["daily_budget"])
            if "daily_budget_365" in data:
                config_kwargs["daily_budget_365_plus"] = float(data["daily_budget_365"])
            if "daily_budget_211" in data:
                config_kwargs["daily_budget_211_330"] = float(data["daily_budget_211"])
            if "default_bid" in data:
                config_kwargs["default_bid"] = float(data["default_bid"])
            if "bidding_strategy" in data:
                config_kwargs["bidding_strategy"] = data["bidding_strategy"]
            if "campaign_prefix" in data:
                config_kwargs["campaign_prefix"] = data["campaign_prefix"]

            config = CampaignConfig(**config_kwargs)

            # Optional fields
            competitor_asins = data.get("competitor_asins")
            if isinstance(competitor_asins, str):
                competitor_asins = [a.strip() for a in competitor_asins.split(",") if a.strip()]

            tier_assignment = data.get("tier_assignment")

            # Find listings report
            listings_path = os.environ.get(
                "LISTINGS_REPORT_PATH",
                os.path.join(project_root, "samples", "sample_active_listings.txt")
            )

            # Generate to temp file
            with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
                output_path = tmp.name

            try:
                result = generate_bulk_upload(
                    target_asins=asins,
                    listings_report_path=listings_path,
                    output_path=output_path,
                    config=config,
                    competitor_asins=competitor_asins,
                    tier_assignment=tier_assignment,
                )

                if not result.is_valid:
                    self._error(422, f"Validation failed: {'; '.join(result.errors)}")
                    return

                # Read the generated file
                with open(output_path, "rb") as f:
                    xlsx_data = f.read()

                # Return XLSX
                self.send_response(200)
                self.send_header("Content-Type",
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                self.send_header("Content-Disposition",
                                 "attachment; filename=campaigns.xlsx")
                self.send_header("Content-Length", str(len(xlsx_data)))
                self.send_header("Access-Control-Allow-Origin", "*")
                self.end_headers()
                self.wfile.write(xlsx_data)

            finally:
                if os.path.exists(output_path):
                    os.unlink(output_path)

        except json.JSONDecodeError:
            self._error(400, "Invalid JSON in request body")
        except Exception as e:
            self._error(500, str(e))

    def do_OPTIONS(self):
        """Handle CORS preflight."""
        self.send_response(200)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def _error(self, code, message):
        self.send_response(code)
        self.send_header("Content-Type", "application/json")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps({"error": message}).encode())
