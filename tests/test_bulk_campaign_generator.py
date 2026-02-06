"""
Unit tests for bulk_campaign_generator.

Covers:
- ASIN→SKU mapping (vectorized and edge cases)
- Deterministic output ordering (AUTO_TARGETING_GROUPS)
- Validation rules (numeric ranges, date format, placement %, enums)
- Config validation and schema enforcement
- Row counts per campaign type
- Active Listings Report parsing (delimiter detection, required column validation)
- Output directory creation
- ASIN format validation helper
"""

import io
import json
import os
import tempfile
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import pytest

# Ensure the project root is importable
import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from bulk_campaign_generator import (
    AUTO_TARGETING_GROUPS,
    BulkCampaignGenerator,
    CampaignConfig,
    DEFAULT_COLUMNS,
    ValidationResult,
    _validate_asin_format,
    _validate_date_format,
    _validate_positive_float,
    build_asin_sku_mapping,
    build_listing_info_map,
    generate_bulk_upload,
    read_active_listings_report,
    validate_all_rows,
    validate_bidding_adjustment_row,
    validate_campaign_row,
    write_bulk_xlsx,
)


# =========================================================================
# Fixtures
# =========================================================================

@pytest.fixture
def sample_listings_tsv(tmp_path):
    """Create a minimal tab-delimited Active Listings Report."""
    content = (
        "seller-sku\tasin1\titem-name\tprice\tquantity\n"
        "SKU-001\tB08NV6CLGF\tMylar Bags 3x4.5 100 Pack\t12.99\t500\n"
        "SKU-002\tB09ABC1234\tMylar Bags 5x7 200 Pack\t19.99\t300\n"
        "SKU-003\tB0DEADBEEF\tMylar Bags 8x12 500 Pack\t29.99\t100\n"
    )
    p = tmp_path / "listings.txt"
    p.write_text(content)
    return str(p)


@pytest.fixture
def sample_listings_csv(tmp_path):
    """Create a comma-delimited Active Listings Report."""
    content = (
        "seller-sku,asin1,item-name,price,quantity\n"
        "SKU-001,B08NV6CLGF,Mylar Bags 3x4.5 100 Pack,12.99,500\n"
        "SKU-002,B09ABC1234,Mylar Bags 5x7 200 Pack,19.99,300\n"
    )
    p = tmp_path / "listings.csv"
    p.write_text(content)
    return str(p)


@pytest.fixture
def sample_asins():
    return ["B08NV6CLGF", "B09ABC1234"]


@pytest.fixture
def default_config():
    return CampaignConfig()


# =========================================================================
# AUTO_TARGETING_GROUPS deterministic ordering
# =========================================================================

class TestDeterministicOrdering:
    def test_auto_targeting_groups_is_tuple(self):
        """AUTO_TARGETING_GROUPS must be a tuple (not a set) for deterministic iteration."""
        assert isinstance(AUTO_TARGETING_GROUPS, tuple)

    def test_auto_targeting_groups_order_is_stable(self):
        """The ordering should be consistent across repeated access."""
        expected = ("close-match", "loose-match", "substitutes", "complements")
        assert AUTO_TARGETING_GROUPS == expected

    def test_generated_rows_have_stable_targeting_order(self, default_config):
        """Product targeting rows in auto campaigns should follow the tuple order."""
        gen = BulkCampaignGenerator(default_config, DEFAULT_COLUMNS)
        gen.generate_clearance_campaigns("365+", ["SKU-A", "SKU-B"])
        rows = gen.get_rows()

        pt_rows = [r for r in rows
                    if r["Entity"] == "Product Targeting"
                    and r["Product Targeting Expression"] in AUTO_TARGETING_GROUPS]
        expressions = [r["Product Targeting Expression"] for r in pt_rows]
        assert expressions == list(AUTO_TARGETING_GROUPS)


# =========================================================================
# ASIN→SKU Mapping
# =========================================================================

class TestAsinSkuMapping:
    def test_basic_mapping(self, sample_listings_tsv):
        df = read_active_listings_report(sample_listings_tsv)
        mapping = build_asin_sku_mapping(df)
        assert mapping["B08NV6CLGF"] == "SKU-001"
        assert mapping["B09ABC1234"] == "SKU-002"
        assert len(mapping) == 3

    def test_csv_fallback(self, sample_listings_csv):
        df = read_active_listings_report(sample_listings_csv)
        mapping = build_asin_sku_mapping(df)
        assert mapping["B08NV6CLGF"] == "SKU-001"

    def test_explicit_delimiter(self, sample_listings_tsv):
        df = read_active_listings_report(sample_listings_tsv, delimiter='\t')
        mapping = build_asin_sku_mapping(df)
        assert len(mapping) == 3

    def test_missing_asin_column_raises(self, tmp_path):
        p = tmp_path / "bad.csv"
        p.write_text("col_a,col_b\n1,2\n")
        with pytest.raises(ValueError, match="missing required columns"):
            read_active_listings_report(str(p))

    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            read_active_listings_report("/nonexistent/file.txt")


# =========================================================================
# Listing Info Map
# =========================================================================

class TestListingInfoMap:
    def test_build_listing_info_map(self, sample_listings_tsv):
        df = read_active_listings_report(sample_listings_tsv)
        info_map = build_listing_info_map(df)
        assert "B08NV6CLGF" in info_map
        assert info_map["B08NV6CLGF"]["item-name"] == "Mylar Bags 3x4.5 100 Pack"
        assert info_map["B08NV6CLGF"]["price"] == "12.99"


# =========================================================================
# Validation
# =========================================================================

class TestValidation:
    def test_positive_float_valid(self):
        assert _validate_positive_float(10.5, "Budget") == []
        assert _validate_positive_float(0, "Bid") == []  # zero is allowed
        assert _validate_positive_float('', "Bid") == []  # blank is ok

    def test_positive_float_negative(self):
        errs = _validate_positive_float(-5, "Budget")
        assert len(errs) == 1
        assert "negative" in errs[0].lower()

    def test_positive_float_non_numeric(self):
        errs = _validate_positive_float("abc", "Bid")
        assert len(errs) == 1
        assert "number" in errs[0].lower()

    def test_date_format_valid(self):
        assert _validate_date_format("20250115", "Start") == []

    def test_date_format_invalid_format(self):
        errs = _validate_date_format("2025-01-15", "Start")
        assert len(errs) == 1
        assert "YYYYMMDD" in errs[0]

    def test_date_format_invalid_date(self):
        errs = _validate_date_format("20251332", "Start")
        assert len(errs) == 1
        assert "not a valid date" in errs[0]

    def test_campaign_row_validates_budget_range(self):
        row = {
            "Entity": "Campaign",
            "Campaign ID": -1001,
            "Campaign Name": "Test",
            "Campaign Daily Budget": -10,
            "Campaign Targeting Type": "MANUAL",
            "Campaign Status": "enabled",
            "Campaign Start Date": "20250115",
            "Campaign Bidding Strategy": "Dynamic bids - down only",
        }
        result = validate_campaign_row(row)
        assert not result.is_valid
        assert any("negative" in e.lower() for e in result.errors)

    def test_campaign_row_validates_date(self):
        row = {
            "Entity": "Campaign",
            "Campaign ID": -1001,
            "Campaign Name": "Test",
            "Campaign Daily Budget": 50.0,
            "Campaign Targeting Type": "MANUAL",
            "Campaign Status": "enabled",
            "Campaign Start Date": "not-a-date",
            "Campaign Bidding Strategy": "Dynamic bids - down only",
        }
        result = validate_campaign_row(row)
        assert not result.is_valid
        assert any("YYYYMMDD" in e for e in result.errors)

    def test_bidding_adjustment_validates_percentage_range(self):
        row = {
            "Entity": "Bidding Adjustment",
            "Campaign ID": -1001,
            "Placement": "placementTop",
            "Percentage": 1000,
        }
        result = validate_bidding_adjustment_row(row)
        assert not result.is_valid
        assert any("0-900" in e for e in result.errors)

    def test_bidding_adjustment_valid_percentage(self):
        row = {
            "Entity": "Bidding Adjustment",
            "Campaign ID": -1001,
            "Placement": "placementTop",
            "Percentage": 80,
        }
        result = validate_bidding_adjustment_row(row)
        assert result.is_valid

    def test_validate_all_rows_passes_for_valid_data(self, default_config):
        gen = BulkCampaignGenerator(default_config, DEFAULT_COLUMNS)
        gen.generate_clearance_campaigns("365+", ["SKU-A"])
        result = validate_all_rows(gen.get_rows())
        assert result.is_valid, f"Validation failed: {result.errors}"


# =========================================================================
# Config Validation
# =========================================================================

class TestCampaignConfig:
    def test_valid_config(self):
        config = CampaignConfig()
        assert config.daily_budget_365_plus == 50.0

    def test_negative_budget_raises(self):
        with pytest.raises(ValueError, match="daily_budget_365_plus must be positive"):
            CampaignConfig(daily_budget_365_plus=-10)

    def test_invalid_bidding_strategy_raises(self):
        with pytest.raises(ValueError, match="Invalid bidding_strategy"):
            CampaignConfig(bidding_strategy="Invalid Strategy")

    def test_placement_out_of_range_raises(self):
        with pytest.raises(ValueError, match="placement_top_percentage"):
            CampaignConfig(placement_top_percentage=1000)

    def test_invalid_start_date_raises(self):
        with pytest.raises(ValueError, match="start_date"):
            CampaignConfig(start_date="bad-date")

    def test_from_dict_warns_on_unknown_keys(self, caplog):
        import logging
        with caplog.at_level(logging.WARNING):
            config = CampaignConfig.from_dict({
                "daily_budget_365_plus": 55.0,
                "typo_key": "oops",
            })
        assert config.daily_budget_365_plus == 55.0
        assert "Unknown keys" in caplog.text
        assert "typo_key" in caplog.text

    def test_from_dict_ignores_unknown_keys(self):
        config = CampaignConfig.from_dict({
            "daily_budget_365_plus": 42.0,
            "unknown_field": 999,
        })
        assert config.daily_budget_365_plus == 42.0


# =========================================================================
# Row Counts (golden-file style)
# =========================================================================

class TestRowCounts:
    def test_single_sku_row_count(self, default_config):
        """Verify expected row counts for a single SKU with no competitors."""
        gen = BulkCampaignGenerator(default_config, DEFAULT_COLUMNS)
        gen.generate_clearance_campaigns("365+", ["SKU-A"],
                                          competitor_asins=["B00COMP001"])

        rows = gen.get_rows()
        campaigns = [r for r in rows if r["Entity"] == "Campaign"]
        bidding_adj = [r for r in rows if r["Entity"] == "Bidding Adjustment"]
        ad_groups = [r for r in rows if r["Entity"] == "Ad Group"]
        product_ads = [r for r in rows if r["Entity"] == "Product Ad"]
        keywords = [r for r in rows if r["Entity"] == "Keyword"]
        pt = [r for r in rows if r["Entity"] == "Product Targeting"]

        # 3 campaigns: Manual KW, Auto, Product Targeting
        assert len(campaigns) == 3
        # 2 placements per campaign = 6
        assert len(bidding_adj) == 6
        # Manual KW: 1 per SKU, Auto: 1, PT: 1 = 3
        assert len(ad_groups) == 3
        # 1 SKU ad per ad group = 3
        assert len(product_ads) == 3

        # Keywords: 5 generic * 3 match types + 5 pack counts = 20 per SKU
        # (size keywords depend on listing info, not added here)
        generic_kw_count = len(default_config.generic_keywords) * 3
        pack_kw_count = len(default_config.pack_counts)
        assert len(keywords) == generic_kw_count + pack_kw_count

        # Product targeting: 4 auto groups + 1 competitor = 5
        assert len(pt) == 4 + 1

    def test_two_sku_row_count(self, default_config):
        """Two SKUs should double per-SKU entities but not campaigns."""
        gen = BulkCampaignGenerator(default_config, DEFAULT_COLUMNS)
        gen.generate_clearance_campaigns("365+", ["SKU-A", "SKU-B"],
                                          competitor_asins=["B00COMP001"])
        rows = gen.get_rows()

        campaigns = [r for r in rows if r["Entity"] == "Campaign"]
        assert len(campaigns) == 3  # still 3 campaigns

        product_ads = [r for r in rows if r["Entity"] == "Product Ad"]
        # Manual KW: 2 (one per SKU ad group) + Auto: 2 + PT: 2 = 6
        assert len(product_ads) == 6


# =========================================================================
# ASIN Format Validation
# =========================================================================

class TestAsinFormatValidation:
    def test_valid_asin(self):
        assert _validate_asin_format("B08NV6CLGF") is True

    def test_too_short(self):
        assert _validate_asin_format("B08NV") is False

    def test_lowercase(self):
        assert _validate_asin_format("b08nv6clgf") is False

    def test_special_chars(self):
        assert _validate_asin_format("B08NV6-LGF") is False


# =========================================================================
# XLSX Writer
# =========================================================================

class TestXlsxWriter:
    def test_write_creates_file(self, tmp_path, default_config):
        gen = BulkCampaignGenerator(default_config, DEFAULT_COLUMNS)
        gen.generate_clearance_campaigns("365+", ["SKU-A"])
        output = str(tmp_path / "test_output.xlsx")
        write_bulk_xlsx(gen.get_rows(), DEFAULT_COLUMNS, output)
        assert Path(output).exists()

    def test_write_creates_parent_dirs(self, tmp_path, default_config):
        gen = BulkCampaignGenerator(default_config, DEFAULT_COLUMNS)
        gen.generate_clearance_campaigns("365+", ["SKU-A"])
        output = str(tmp_path / "subdir" / "nested" / "output.xlsx")
        write_bulk_xlsx(gen.get_rows(), DEFAULT_COLUMNS, output)
        assert Path(output).exists()


# =========================================================================
# End-to-end (no Supabase)
# =========================================================================

class TestEndToEnd:
    def test_generate_bulk_upload(self, sample_listings_tsv, sample_asins,
                                  tmp_path):
        output = str(tmp_path / "bulk_upload.xlsx")
        result = generate_bulk_upload(
            target_asins=sample_asins,
            listings_report_path=sample_listings_tsv,
            output_path=output,
            competitor_asins=["B00COMP001"],
            config=CampaignConfig(),
            store=None,
        )
        assert result.is_valid, f"Validation failed: {result.errors}"
        assert Path(output).exists()

    def test_generate_with_missing_asin_still_succeeds(
        self, sample_listings_tsv, tmp_path
    ):
        output = str(tmp_path / "bulk_upload.xlsx")
        result = generate_bulk_upload(
            target_asins=["B08NV6CLGF", "B00NONEXIST"],
            listings_report_path=sample_listings_tsv,
            output_path=output,
            config=CampaignConfig(),
            store=None,
        )
        # Should succeed with at least one valid ASIN
        assert result.is_valid

    def test_generate_with_all_missing_asins_fails(
        self, sample_listings_tsv, tmp_path
    ):
        output = str(tmp_path / "bulk_upload.xlsx")
        result = generate_bulk_upload(
            target_asins=["B00NONEXIST"],
            listings_report_path=sample_listings_tsv,
            output_path=output,
            config=CampaignConfig(),
            store=None,
        )
        assert not result.is_valid
