"""
Supabase integration for the Amazon Ads Bulk Campaign Generator.

Provides persistent storage for:
- ASIN→SKU listings (from Active Listings Reports)
- Generation run history
- Upload result tracking
- XLSX file storage

This module is entirely optional. The generator works without it.
Enable by setting SUPABASE_URL and SUPABASE_KEY environment variables.
"""

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional

try:
    from supabase import create_client, Client
    HAS_SUPABASE = True
except ImportError:
    HAS_SUPABASE = False
    Client = None


class SupabaseStore:
    """Optional Supabase backend for persistent state."""

    BUCKET = "campaign-files"

    def __init__(self, url: Optional[str] = None, key: Optional[str] = None):
        """
        Initialize Supabase client.

        Args:
            url: Supabase project URL (or set SUPABASE_URL env var)
            key: Supabase anon/service key (or set SUPABASE_KEY env var)

        Raises:
            RuntimeError: If supabase-py is not installed
            ValueError: If URL or key is missing
        """
        if not HAS_SUPABASE:
            raise RuntimeError(
                "supabase-py is not installed. "
                "Run: pip install supabase"
            )

        self.url = url or os.environ.get("SUPABASE_URL", "")
        self.key = key or os.environ.get("SUPABASE_KEY", "")

        if not self.url or not self.key:
            raise ValueError(
                "Supabase credentials not configured. Set SUPABASE_URL and "
                "SUPABASE_KEY environment variables, or pass them directly."
            )

        self.client: Client = create_client(self.url, self.key)

    # =========================================================================
    # Listings
    # =========================================================================

    def sync_listings(self, listings_data: List[dict]) -> dict:
        """
        Upsert listings from an Active Listings Report parse.

        Each item in listings_data should have at minimum:
            {"asin": "B08...", "seller_sku": "my-sku-123"}

        Optional fields: item_name, fulfillment_channel, price, quantity

        Args:
            listings_data: List of listing dicts

        Returns:
            Dict with counts: {"upserted": N, "errors": N}
        """
        now = datetime.now(timezone.utc).isoformat()
        upserted = 0
        errors = 0

        for item in listings_data:
            row = {
                "asin": item["asin"],
                "seller_sku": item["seller_sku"],
                "last_seen_at": now,
                "status": "active",
            }

            # Add optional fields if present
            for field in ("item_name", "fulfillment_channel", "price", "quantity"):
                if field in item and item[field]:
                    row[field] = item[field]

            try:
                self.client.table("listings").upsert(
                    row,
                    on_conflict="asin,seller_sku"
                ).execute()
                upserted += 1
            except Exception as e:
                print(f"  Warning: Failed to upsert listing {item['asin']}: {e}")
                errors += 1

        return {"upserted": upserted, "errors": errors}

    def get_sku_for_asin(self, asin: str) -> Optional[str]:
        """
        Look up the seller SKU for an ASIN from stored listings.

        Uses the most recently seen SKU if multiple exist.

        Args:
            asin: Amazon ASIN

        Returns:
            Seller SKU string, or None if not found
        """
        result = (
            self.client.table("listings")
            .select("seller_sku")
            .eq("asin", asin)
            .eq("status", "active")
            .order("last_seen_at", desc=True)
            .limit(1)
            .execute()
        )

        if result.data:
            return result.data[0]["seller_sku"]
        return None

    def get_all_active_listings(self) -> List[dict]:
        """
        Fetch all active listings from Supabase.

        Returns:
            List of listing dicts
        """
        result = (
            self.client.table("listings")
            .select("*")
            .eq("status", "active")
            .order("last_seen_at", desc=True)
            .execute()
        )
        return result.data or []

    def build_asin_sku_map(self) -> Dict[str, str]:
        """
        Build ASIN→SKU mapping from all active listings in Supabase.

        Returns:
            Dict mapping ASIN to seller_sku
        """
        listings = self.get_all_active_listings()
        mapping = {}
        for row in listings:
            asin = row["asin"]
            # First seen SKU wins (already ordered by last_seen_at desc)
            if asin not in mapping:
                mapping[asin] = row["seller_sku"]
        return mapping

    def mark_listings_inactive(self, asins: List[str]) -> int:
        """
        Mark listings as inactive (e.g., when they disappear from a report).

        Args:
            asins: List of ASINs to mark inactive

        Returns:
            Number of rows updated
        """
        if not asins:
            return 0

        result = (
            self.client.table("listings")
            .update({"status": "inactive"})
            .in_("asin", asins)
            .execute()
        )
        return len(result.data) if result.data else 0

    # =========================================================================
    # Generation Runs
    # =========================================================================

    def record_generation_run(
        self,
        target_asins: List[str],
        matched_skus: Dict[str, str],
        missing_asins: List[str],
        config: dict,
        row_count: int,
        campaign_count: int,
        tier_assignment: Optional[dict] = None,
        competitor_asins: Optional[List[str]] = None,
        file_name: Optional[str] = None,
        notes: Optional[str] = None,
    ) -> int:
        """
        Record a generation run in the database.

        Args:
            target_asins: ASINs that were targeted
            matched_skus: Dict of ASIN→SKU that were found
            missing_asins: ASINs not found in listings
            config: CampaignConfig as dict
            row_count: Total rows in the generated file
            campaign_count: Number of campaigns created
            tier_assignment: ASIN→tier mapping used
            competitor_asins: Competitor ASINs used
            file_name: Name of the generated file
            notes: Optional notes

        Returns:
            The generation run ID
        """
        row = {
            "target_asins": target_asins,
            "matched_skus": matched_skus,
            "missing_asins": missing_asins,
            "config": config,
            "row_count": row_count,
            "campaign_count": campaign_count,
            "file_name": file_name,
            "notes": notes,
        }

        if tier_assignment is not None:
            row["tier_assignment"] = tier_assignment
        if competitor_asins is not None:
            row["competitor_asins"] = competitor_asins

        result = self.client.table("generation_runs").insert(row).execute()
        return result.data[0]["id"]

    def get_generation_history(self, limit: int = 20) -> List[dict]:
        """
        Fetch recent generation runs.

        Args:
            limit: Max number of runs to return

        Returns:
            List of generation run dicts, most recent first
        """
        result = (
            self.client.table("generation_runs")
            .select("*")
            .order("created_at", desc=True)
            .limit(limit)
            .execute()
        )
        return result.data or []

    def get_generation_run(self, run_id: int) -> Optional[dict]:
        """
        Fetch a specific generation run by ID.

        Args:
            run_id: The generation run ID

        Returns:
            Generation run dict, or None
        """
        result = (
            self.client.table("generation_runs")
            .select("*")
            .eq("id", run_id)
            .limit(1)
            .execute()
        )
        return result.data[0] if result.data else None

    # =========================================================================
    # Upload Results
    # =========================================================================

    def record_upload_result(
        self,
        generation_run_id: int,
        status: str = "pending",
        total_rows: Optional[int] = None,
        processed_rows: Optional[int] = None,
        failed_rows: Optional[int] = None,
        error_summary: Optional[str] = None,
        error_details: Optional[list] = None,
        notes: Optional[str] = None,
    ) -> int:
        """
        Record the result of uploading a bulk file to Amazon Ads.

        Args:
            generation_run_id: ID of the generation run
            status: One of "pending", "success", "partial", "failed"
            total_rows: Total rows submitted
            processed_rows: Rows successfully processed
            failed_rows: Rows that failed
            error_summary: Human-readable summary
            error_details: Structured error data
            notes: Additional notes

        Returns:
            The upload result ID
        """
        row = {
            "generation_run_id": generation_run_id,
            "status": status,
        }

        if total_rows is not None:
            row["total_rows"] = total_rows
        if processed_rows is not None:
            row["processed_rows"] = processed_rows
        if failed_rows is not None:
            row["failed_rows"] = failed_rows
        if error_summary is not None:
            row["error_summary"] = error_summary
        if error_details is not None:
            row["error_details"] = error_details
        if notes is not None:
            row["notes"] = notes

        result = self.client.table("upload_results").insert(row).execute()
        return result.data[0]["id"]

    def update_upload_result(self, upload_id: int, **kwargs) -> None:
        """
        Update an existing upload result.

        Args:
            upload_id: The upload result ID
            **kwargs: Fields to update
        """
        self.client.table("upload_results").update(kwargs).eq("id", upload_id).execute()

    def get_upload_results(self, generation_run_id: int) -> List[dict]:
        """
        Fetch upload results for a generation run.

        Args:
            generation_run_id: The generation run ID

        Returns:
            List of upload result dicts
        """
        result = (
            self.client.table("upload_results")
            .select("*")
            .eq("generation_run_id", generation_run_id)
            .order("uploaded_at", desc=True)
            .execute()
        )
        return result.data or []

    # =========================================================================
    # File Storage
    # =========================================================================

    def upload_file(self, local_path: str, storage_path: Optional[str] = None) -> str:
        """
        Upload a generated XLSX file to Supabase Storage.

        Args:
            local_path: Path to the local file
            storage_path: Remote path in the bucket (defaults to filename)

        Returns:
            The storage path
        """
        path = Path(local_path)
        if storage_path is None:
            timestamp = datetime.now(timezone.utc).strftime("%Y%m%d_%H%M%S")
            storage_path = f"{timestamp}_{path.name}"

        with open(local_path, "rb") as f:
            self.client.storage.from_(self.BUCKET).upload(
                storage_path,
                f,
                file_options={
                    "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                }
            )

        return storage_path

    def download_file(self, storage_path: str, local_path: str) -> None:
        """
        Download a file from Supabase Storage.

        Args:
            storage_path: Path in the storage bucket
            local_path: Local path to save to
        """
        data = self.client.storage.from_(self.BUCKET).download(storage_path)
        with open(local_path, "wb") as f:
            f.write(data)

    def get_file_url(self, storage_path: str, expires_in: int = 3600) -> str:
        """
        Get a signed URL for a stored file.

        Args:
            storage_path: Path in the storage bucket
            expires_in: URL expiration in seconds (default 1 hour)

        Returns:
            Signed URL string
        """
        result = self.client.storage.from_(self.BUCKET).create_signed_url(
            storage_path, expires_in
        )
        return result["signedURL"]


def get_store() -> Optional[SupabaseStore]:
    """
    Try to create a SupabaseStore from environment variables.

    Returns:
        SupabaseStore if configured, None otherwise.
        Never raises - just returns None if anything is missing.
    """
    url = os.environ.get("SUPABASE_URL", "")
    key = os.environ.get("SUPABASE_KEY", "")

    if not url or not key:
        return None

    if not HAS_SUPABASE:
        print("Warning: SUPABASE_URL is set but supabase-py is not installed. "
              "Run: pip install supabase")
        return None

    try:
        return SupabaseStore(url, key)
    except Exception as e:
        print(f"Warning: Could not connect to Supabase: {e}")
        return None
