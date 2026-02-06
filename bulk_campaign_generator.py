#!/usr/bin/env python3
"""
Amazon Sponsored Products Bulk Campaign Generator

Generates upload-ready XLSX files for Amazon Ads bulk operations using:
1. Target ASINs (clearance inventory)
2. Active Listings Report from Seller Central (ASIN→SKU mapping)
3. Amazon Ads bulk operations template (column headers source)

Output: Upload-ready .xlsx with only "Sponsored Products Campaigns" sheet.
"""

import argparse
import csv
import json
import logging
import re
import sys
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook, load_workbook

logger = logging.getLogger(__name__)

from supabase_client import SupabaseStore, get_store


# =============================================================================
# CONSTANTS - Allowed values (exact strings from Amazon)
# =============================================================================

TARGETING_TYPES = {"AUTO", "MANUAL"}

BIDDING_STRATEGIES = {
    "Dynamic bids - down only",
    "Dynamic bids - up and down",
    "Fixed bid",
}

PLACEMENTS = {
    "placementTop",
    "placementProductPage",
    "placementRestOfSearch",
    "placementAmazonBusiness",
}

MATCH_TYPES = {"exact", "phrase", "broad"}

STATES = {"enabled", "paused", "archived"}

AUTO_TARGETING_GROUPS = ("close-match", "loose-match", "substitutes", "complements")

# Sheet name required by Amazon
SHEET_NAME = "Sponsored Products Campaigns"

# Default column headers (fallback if no template provided)
DEFAULT_COLUMNS = [
    "Product",
    "Entity",
    "Operation",
    "Campaign ID",
    "Campaign Name",
    "Campaign Name (Informational only)",
    "Campaign Daily Budget",
    "Portfolio ID",
    "Campaign Start Date",
    "Campaign End Date",
    "Campaign Targeting Type",
    "Campaign Status",
    "Campaign Bidding Strategy",
    "Ad Group ID",
    "Ad Group Name",
    "Ad Group Name (Informational only)",
    "Ad Group Default Bid",
    "Ad Group Status",
    "SKU",
    "ASIN",
    "ASIN (Informational only)",
    "Ad Status",
    "Keyword ID",
    "Keyword Text",
    "Match Type",
    "Keyword Status",
    "Bid",
    "Keyword Bid (Informational only)",
    "Product Targeting ID",
    "Product Targeting Expression",
    "Product Targeting Expression Type",
    "Product Targeting Status",
    "Product Targeting Bid",
    "Product Targeting Bid (Informational only)",
    "Placement",
    "Percentage",
]


# =============================================================================
# DATA CLASSES
# =============================================================================

_KNOWN_CONFIG_KEYS = {
    "daily_budget_365_plus", "daily_budget_211_330", "default_bid",
    "bidding_strategy", "placement_top_percentage",
    "placement_product_page_percentage", "start_date", "campaign_prefix",
    "generic_keywords", "size_tokens", "pack_counts",
}


@dataclass
class CampaignConfig:
    """Configuration for campaign generation."""
    # Campaign settings
    daily_budget_365_plus: float = 50.0
    daily_budget_211_330: float = 30.0
    default_bid: float = 0.75

    # Bidding strategy
    bidding_strategy: str = "Dynamic bids - down only"

    # Placement adjustments (percentage, not decimal)
    placement_top_percentage: int = 80
    placement_product_page_percentage: int = 50

    # Date settings
    start_date: Optional[str] = None  # YYYYMMDD format, defaults to today

    # Campaign naming
    campaign_prefix: str = "Clearance"

    # Keywords to use for manual campaigns
    generic_keywords: list = field(default_factory=lambda: [
        "mylar bags",
        "mylar bags for food storage",
        "food storage bags",
        "oxygen absorber bags",
        "zipper mylar bags",
    ])

    # Size tokens for keyword generation
    size_tokens: list = field(default_factory=lambda: [
        "3x4.5",
        "3.6x5",
        "4x6",
        "5x7",
        "6x9",
        "8x12",
        "10x14",
    ])

    # Pack count keywords
    pack_counts: list = field(default_factory=lambda: [
        "100 pack",
        "200 pack",
        "500 pack",
        "1000 pack",
        "1000 count",
    ])

    def __post_init__(self):
        """Validate config values after initialization."""
        errors = []

        # Validate budgets are positive
        if self.daily_budget_365_plus <= 0:
            errors.append(
                f"daily_budget_365_plus must be positive "
                f"(got {self.daily_budget_365_plus})")
        if self.daily_budget_211_330 <= 0:
            errors.append(
                f"daily_budget_211_330 must be positive "
                f"(got {self.daily_budget_211_330})")

        # Validate default bid is positive
        if self.default_bid <= 0:
            errors.append(
                f"default_bid must be positive (got {self.default_bid})")

        # Validate bidding strategy
        if self.bidding_strategy not in BIDDING_STRATEGIES:
            errors.append(
                f"Invalid bidding_strategy: {self.bidding_strategy!r}. "
                f"Must be one of: {', '.join(sorted(BIDDING_STRATEGIES))}")

        # Validate placement percentages (Amazon allows 0-900)
        for name, val in [
            ("placement_top_percentage", self.placement_top_percentage),
            ("placement_product_page_percentage",
             self.placement_product_page_percentage),
        ]:
            if not isinstance(val, int) or val < 0 or val > 900:
                errors.append(
                    f"{name} must be an integer 0-900 (got {val})")

        # Validate start_date format if provided
        if self.start_date is not None:
            sd = str(self.start_date)
            if len(sd) != 8 or not sd.isdigit():
                errors.append(
                    f"start_date must be YYYYMMDD format (got {sd!r})")
            else:
                try:
                    datetime.strptime(sd, '%Y%m%d')
                except ValueError:
                    errors.append(
                        f"start_date is not a valid date (got {sd!r})")

        if errors:
            raise ValueError(
                "Invalid campaign configuration:\n  - "
                + "\n  - ".join(errors))

    @classmethod
    def from_dict(cls, data: dict) -> "CampaignConfig":
        """
        Create a CampaignConfig from a dict, warning on unknown keys.

        This is preferred over ``CampaignConfig(**data)`` when loading
        from user-provided JSON, because it catches typos in key names.
        """
        unknown = set(data.keys()) - _KNOWN_CONFIG_KEYS
        if unknown:
            logger.warning(
                "Unknown keys in config (possibly typos): %s",
                ", ".join(sorted(unknown)),
            )
        # Only pass known keys to avoid TypeError on unknown kwargs
        filtered = {k: v for k, v in data.items() if k in _KNOWN_CONFIG_KEYS}
        return cls(**filtered)


@dataclass
class ValidationResult:
    """Result of validation."""
    is_valid: bool
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


# =============================================================================
# ACTIVE LISTINGS REPORT PARSER
# =============================================================================

REQUIRED_LISTING_COLUMNS = {
    "asin": ['asin1', 'ASIN1', 'asin', 'ASIN'],
    "sku": ['seller-sku', 'Seller SKU', 'sku', 'SKU'],
}


def read_active_listings_report(
    filepath: str,
    delimiter: Optional[str] = None,
) -> pd.DataFrame:
    """
    Read Seller Central Active Listings Report.

    Delimiter resolution order:
    1. Explicit ``delimiter`` argument (e.g. from ``--delimiter`` CLI flag).
    2. Auto-detect via ``csv.Sniffer`` on the first 8 KB.
    3. Fallback chain: tab-delimited, then comma-delimited CSV.

    After parsing, the function validates that columns needed for ASIN and
    SKU mapping are present in the resulting DataFrame.

    Args:
        filepath: Path to the Active Listings Report file
        delimiter: Optional explicit delimiter character

    Returns:
        DataFrame with listing data

    Raises:
        FileNotFoundError: If the file does not exist
        ValueError: If the file cannot be parsed or required columns are missing
    """
    path = Path(filepath)

    if not path.exists():
        raise FileNotFoundError(f"Active Listings Report not found: {filepath}")

    df = None
    parse_errors = []

    # 1. Explicit delimiter
    if delimiter is not None:
        try:
            df = pd.read_csv(filepath, sep=delimiter, dtype=str,
                             keep_default_na=False)
        except Exception as e:
            raise ValueError(
                f"Could not parse Active Listings Report with "
                f"delimiter {delimiter!r}: {e}"
            )
    else:
        # 2. Auto-detect via csv.Sniffer
        try:
            with open(filepath, 'r', newline='') as f:
                sample = f.read(8192)
            detected = csv.Sniffer().sniff(sample)
            df = pd.read_csv(filepath, sep=detected.delimiter, dtype=str,
                             keep_default_na=False)
            if len(df.columns) <= 1:
                df = None
                parse_errors.append(
                    f"Sniffer detected delimiter {detected.delimiter!r} "
                    f"but resulted in only 1 column"
                )
        except Exception as e:
            parse_errors.append(f"csv.Sniffer failed: {e}")

        # 3. Fallback: tab then comma
        if df is None:
            try:
                df = pd.read_csv(filepath, sep='\t', dtype=str,
                                 keep_default_na=False)
                if len(df.columns) <= 1:
                    parse_errors.append(
                        "Tab-delimited parse yielded only 1 column")
                    df = None
            except Exception as e:
                parse_errors.append(f"Tab-delimited parse failed: {e}")

        if df is None:
            try:
                df = pd.read_csv(filepath, dtype=str,
                                 keep_default_na=False)
            except Exception as e:
                parse_errors.append(f"CSV parse failed: {e}")

    if df is None:
        raise ValueError(
            f"Could not parse Active Listings Report '{filepath}'. "
            f"Attempts: {'; '.join(parse_errors)}"
        )

    if parse_errors:
        logger.debug("Listing report parse notes: %s", "; ".join(parse_errors))

    # Validate that required columns are present
    missing_cols = []
    for role, candidates in REQUIRED_LISTING_COLUMNS.items():
        if not any(c in df.columns for c in candidates):
            missing_cols.append(
                f"{role} (expected one of: {', '.join(candidates)})"
            )

    if missing_cols:
        raise ValueError(
            f"Active Listings Report is missing required columns: "
            f"{'; '.join(missing_cols)}. "
            f"Detected columns: {list(df.columns)}"
        )

    logger.info("Parsed Active Listings Report: %d rows, %d columns "
                "(delimiter=%r)", len(df), len(df.columns),
                delimiter or 'auto')

    return df


def _find_column(df: pd.DataFrame, candidates: list, role: str) -> str:
    """Find the first matching column name from a list of candidates."""
    for col in candidates:
        if col in df.columns:
            return col
    raise ValueError(
        f"Could not find {role} column in Active Listings Report. "
        f"Expected one of: {candidates}"
    )


def build_asin_sku_mapping(listings_df: pd.DataFrame) -> dict:
    """
    Build ASIN -> SKU mapping from Active Listings Report.

    Uses vectorized pandas operations for better performance on large files.

    Args:
        listings_df: DataFrame from Active Listings Report

    Returns:
        Dict mapping ASIN to seller-sku
    """
    asin_col = _find_column(listings_df,
                            ['asin1', 'ASIN1', 'asin', 'ASIN'], 'ASIN')
    sku_col = _find_column(listings_df,
                           ['seller-sku', 'Seller SKU', 'sku', 'SKU'], 'SKU')

    # Vectorized: strip whitespace and drop empty rows
    asins = listings_df[asin_col].astype(str).str.strip()
    skus = listings_df[sku_col].astype(str).str.strip()
    mask = (asins != '') & (skus != '')

    return dict(zip(asins[mask], skus[mask]))


def build_listing_info_map(listings_df: pd.DataFrame) -> dict:
    """
    Build a lookup dict mapping ASIN -> listing info for all rows.

    Pre-builds the entire map once so per-ASIN lookups avoid repeated
    DataFrame filtering.

    Returns:
        Dict mapping ASIN to info dict with keys like 'item-name', 'price', etc.
    """
    try:
        asin_col = _find_column(listings_df,
                                ['asin1', 'ASIN1', 'asin', 'ASIN'], 'ASIN')
    except ValueError:
        return {}

    field_mapping = {
        'item-name': ['item-name', 'Item Name', 'title', 'Title'],
        'fulfillment-channel': ['fulfillment-channel', 'Fulfillment Channel'],
        'price': ['price', 'Price'],
        'quantity': ['quantity', 'Quantity'],
    }

    # Resolve which actual columns exist
    resolved_fields = {}
    for key, candidates in field_mapping.items():
        for col in candidates:
            if col in listings_df.columns:
                resolved_fields[key] = col
                break

    info_map = {}
    for _, row in listings_df.iterrows():
        asin = str(row[asin_col]).strip()
        if not asin or asin in info_map:
            continue
        info = {}
        for key, col in resolved_fields.items():
            val = str(row[col]).strip()
            if val:
                info[key] = val
        info_map[asin] = info

    return info_map


def get_listing_info(listings_df: pd.DataFrame, asin: str) -> dict:
    """
    Get additional listing info for an ASIN.

    For bulk lookups, prefer ``build_listing_info_map()`` which avoids
    per-ASIN DataFrame filtering.

    Args:
        listings_df: DataFrame from Active Listings Report
        asin: ASIN to look up

    Returns:
        Dict with listing info (item-name, fulfillment-channel, price, etc.)
    """
    try:
        asin_col = _find_column(listings_df,
                                ['asin1', 'ASIN1', 'asin', 'ASIN'], 'ASIN')
    except ValueError:
        return {}

    row = listings_df[listings_df[asin_col] == asin]
    if row.empty:
        return {}

    row = row.iloc[0]
    info = {}

    field_mapping = {
        'item-name': ['item-name', 'Item Name', 'title', 'Title'],
        'fulfillment-channel': ['fulfillment-channel', 'Fulfillment Channel'],
        'price': ['price', 'Price'],
        'quantity': ['quantity', 'Quantity'],
    }

    for key, possible_cols in field_mapping.items():
        for col in possible_cols:
            if col in row.index and str(row[col]).strip():
                info[key] = str(row[col]).strip()
                break

    return info


# =============================================================================
# TEMPLATE READER
# =============================================================================

def read_template_headers(template_path: Optional[str] = None) -> list:
    """
    Read column headers from Amazon Ads bulk operations template.

    Args:
        template_path: Path to template XLSX file (optional)

    Returns:
        List of column headers
    """
    if template_path is None:
        return DEFAULT_COLUMNS.copy()

    path = Path(template_path)
    if not path.exists():
        logger.warning("Template not found at %s, using default columns",
                       template_path)
        return DEFAULT_COLUMNS.copy()

    try:
        wb = load_workbook(template_path, read_only=True)

        # Try to find the correct sheet
        sheet = None
        for name in [SHEET_NAME, "Sheet1", wb.sheetnames[0]]:
            if name in wb.sheetnames:
                sheet = wb[name]
                break

        if sheet is None:
            sheet = wb.active

        # Read headers from first row
        headers = []
        for cell in sheet[1]:
            if cell.value:
                headers.append(str(cell.value))
            else:
                break

        wb.close()

        if headers:
            return headers
        else:
            logger.warning("No headers found in template, using default columns")
            return DEFAULT_COLUMNS.copy()

    except Exception as e:
        logger.warning("Could not read template (%s), using default columns", e)
        return DEFAULT_COLUMNS.copy()


# =============================================================================
# VALIDATION
# =============================================================================

def _validate_positive_float(value, field_name: str) -> list:
    """Validate that a value is a positive number. Returns list of errors."""
    if value == '' or value is None:
        return []
    try:
        num = float(value)
        if num < 0:
            return [f"{field_name} must not be negative (got {num})"]
        return []
    except (ValueError, TypeError):
        return [f"{field_name} must be a number (got {value!r})"]


def _validate_date_format(value, field_name: str) -> list:
    """Validate YYYYMMDD date format. Returns list of errors."""
    if value == '' or value is None:
        return []
    val_str = str(value)
    if len(val_str) != 8 or not val_str.isdigit():
        return [f"{field_name} must be in YYYYMMDD format (got {val_str!r})"]
    try:
        datetime.strptime(val_str, '%Y%m%d')
    except ValueError:
        return [f"{field_name} is not a valid date (got {val_str!r})"]
    return []


def validate_campaign_row(row: dict) -> ValidationResult:
    """Validate a Campaign entity row."""
    errors = []
    warnings = []

    required = ['Campaign ID', 'Campaign Name', 'Campaign Daily Budget',
                'Campaign Targeting Type', 'Campaign Status', 'Campaign Start Date',
                'Campaign Bidding Strategy']

    for fld in required:
        val = row.get(fld, '')
        if val == '' or val is None:
            errors.append(f"Campaign missing required field: {fld}")

    # Validate enums
    targeting = row.get('Campaign Targeting Type', '')
    if targeting and targeting not in TARGETING_TYPES:
        errors.append(f"Invalid Targeting Type: {targeting}")

    strategy = row.get('Campaign Bidding Strategy', '')
    if strategy and strategy not in BIDDING_STRATEGIES:
        errors.append(f"Invalid Bidding Strategy: {strategy}")

    status = row.get('Campaign Status', '')
    if status and status not in STATES:
        errors.append(f"Invalid Campaign Status: {status}")

    # Validate numeric range: daily budget must be positive
    errors.extend(_validate_positive_float(
        row.get('Campaign Daily Budget', ''), 'Campaign Daily Budget'))

    # Validate date format
    errors.extend(_validate_date_format(
        row.get('Campaign Start Date', ''), 'Campaign Start Date'))
    errors.extend(_validate_date_format(
        row.get('Campaign End Date', ''), 'Campaign End Date'))

    return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)


def validate_bidding_adjustment_row(row: dict) -> ValidationResult:
    """Validate a Bidding Adjustment entity row."""
    errors = []
    warnings = []

    required = ['Campaign ID', 'Placement', 'Percentage']

    for fld in required:
        val = row.get(fld, '')
        if val == '' or val is None:
            errors.append(f"Bidding Adjustment missing required field: {fld}")

    placement = row.get('Placement', '')
    if placement and placement not in PLACEMENTS:
        errors.append(f"Invalid Placement: {placement}")

    # Validate percentage range (Amazon allows 0-900%)
    pct = row.get('Percentage', '')
    if pct != '' and pct is not None:
        try:
            pct_val = int(pct)
            if pct_val < 0 or pct_val > 900:
                errors.append(
                    f"Placement Percentage must be 0-900 (got {pct_val})")
        except (ValueError, TypeError):
            errors.append(f"Placement Percentage must be an integer (got {pct!r})")

    return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)


def validate_ad_group_row(row: dict) -> ValidationResult:
    """Validate an Ad Group entity row."""
    errors = []
    warnings = []

    required = ['Campaign ID', 'Ad Group ID', 'Ad Group Name',
                'Ad Group Default Bid', 'Ad Group Status']

    for fld in required:
        val = row.get(fld, '')
        if val == '' or val is None:
            errors.append(f"Ad Group missing required field: {fld}")

    # Validate default bid is a positive number
    errors.extend(_validate_positive_float(
        row.get('Ad Group Default Bid', ''), 'Ad Group Default Bid'))

    status = row.get('Ad Group Status', '')
    if status and status not in STATES:
        errors.append(f"Invalid Ad Group Status: {status}")

    return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)


def validate_product_ad_row(row: dict) -> ValidationResult:
    """Validate a Product Ad entity row."""
    errors = []

    required = ['Campaign ID', 'Ad Group ID', 'SKU', 'Ad Status']

    for field in required:
        val = row.get(field, '')
        if val == '' or val is None:
            errors.append(f"Product Ad missing required field: {field}")

    return ValidationResult(is_valid=len(errors) == 0, errors=errors)


def validate_keyword_row(row: dict) -> ValidationResult:
    """Validate a Keyword entity row."""
    errors = []
    warnings = []

    required = ['Campaign ID', 'Ad Group ID', 'Keyword Text', 'Match Type', 'Keyword Status']

    for fld in required:
        val = row.get(fld, '')
        if val == '' or val is None:
            errors.append(f"Keyword missing required field: {fld}")

    match_type = row.get('Match Type', '')
    if match_type and match_type not in MATCH_TYPES:
        errors.append(f"Invalid Match Type: {match_type}")

    # Validate bid is positive if present
    errors.extend(_validate_positive_float(row.get('Bid', ''), 'Keyword Bid'))

    return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)


def validate_product_targeting_row(row: dict) -> ValidationResult:
    """Validate a Product Targeting entity row."""
    errors = []
    warnings = []

    required = ['Campaign ID', 'Ad Group ID', 'Product Targeting Expression',
                'Product Targeting Status']

    for fld in required:
        val = row.get(fld, '')
        if val == '' or val is None:
            errors.append(f"Product Targeting missing required field: {fld}")

    # Validate bid is positive if present
    errors.extend(_validate_positive_float(
        row.get('Product Targeting Bid', ''), 'Product Targeting Bid'))

    return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)


def validate_all_rows(rows: list) -> ValidationResult:
    """
    Validate all rows in the bulk sheet.

    Args:
        rows: List of row dicts

    Returns:
        Combined ValidationResult
    """
    all_errors = []
    all_warnings = []

    validators = {
        'Campaign': validate_campaign_row,
        'Bidding Adjustment': validate_bidding_adjustment_row,
        'Ad Group': validate_ad_group_row,
        'Product Ad': validate_product_ad_row,
        'Keyword': validate_keyword_row,
        'Product Targeting': validate_product_targeting_row,
    }

    for i, row in enumerate(rows):
        entity = row.get('Entity', '')
        if entity in validators:
            result = validators[entity](row)
            if not result.is_valid:
                for error in result.errors:
                    all_errors.append(f"Row {i+2}: {error}")
            all_warnings.extend(result.warnings)

    return ValidationResult(
        is_valid=len(all_errors) == 0,
        errors=all_errors,
        warnings=all_warnings
    )


# =============================================================================
# CAMPAIGN DATA GENERATOR
# =============================================================================

class BulkCampaignGenerator:
    """Generates Amazon Ads bulk campaign data."""

    def __init__(self, config: CampaignConfig, headers: list):
        self.config = config
        self.headers = headers
        self.rows = []

        # ID counters (negative integers as per Amazon convention)
        self._campaign_id_counter = -1000
        self._ad_group_id_counter = -2000

        # Start date
        if config.start_date:
            self.start_date = config.start_date
        else:
            self.start_date = datetime.now().strftime('%Y%m%d')

    def _next_campaign_id(self) -> int:
        self._campaign_id_counter -= 1
        return self._campaign_id_counter

    def _next_ad_group_id(self) -> int:
        self._ad_group_id_counter -= 1
        return self._ad_group_id_counter

    def _create_empty_row(self) -> dict:
        """Create an empty row with all headers."""
        return {h: '' for h in self.headers}

    def _add_row(self, **kwargs) -> dict:
        """Add a row with specified values, leaving others blank."""
        row = self._create_empty_row()
        for key, value in kwargs.items():
            if key in row:
                # Never write 0 for blank fields - critical!
                if value is None or value == '':
                    row[key] = ''
                else:
                    row[key] = value
        self.rows.append(row)
        return row

    def add_campaign(self, name: str, daily_budget: float,
                     targeting_type: str, campaign_id: Optional[int] = None) -> int:
        """
        Add a Campaign row.

        Returns:
            Campaign ID (negative integer)
        """
        if campaign_id is None:
            campaign_id = self._next_campaign_id()

        self._add_row(
            Product="Sponsored Products",
            Entity="Campaign",
            Operation="Create",
            **{"Campaign ID": campaign_id},
            **{"Campaign Name": name},
            **{"Campaign Daily Budget": daily_budget},
            **{"Campaign Start Date": self.start_date},
            **{"Campaign Targeting Type": targeting_type},
            **{"Campaign Status": "enabled"},
            **{"Campaign Bidding Strategy": self.config.bidding_strategy},
        )

        return campaign_id

    def add_bidding_adjustment(self, campaign_id: int, placement: str,
                               percentage: int) -> None:
        """Add a Bidding Adjustment row."""
        self._add_row(
            Product="Sponsored Products",
            Entity="Bidding Adjustment",
            Operation="Create",
            **{"Campaign ID": campaign_id},
            Placement=placement,
            Percentage=percentage,
        )

    def add_ad_group(self, campaign_id: int, name: str,
                     default_bid: float, ad_group_id: Optional[int] = None) -> int:
        """
        Add an Ad Group row.

        Returns:
            Ad Group ID (negative integer)
        """
        if ad_group_id is None:
            ad_group_id = self._next_ad_group_id()

        self._add_row(
            Product="Sponsored Products",
            Entity="Ad Group",
            Operation="Create",
            **{"Campaign ID": campaign_id},
            **{"Ad Group ID": ad_group_id},
            **{"Ad Group Name": name},
            **{"Ad Group Default Bid": default_bid},
            **{"Ad Group Status": "enabled"},
        )

        return ad_group_id

    def add_product_ad(self, campaign_id: int, ad_group_id: int, sku: str) -> None:
        """Add a Product Ad row."""
        self._add_row(
            Product="Sponsored Products",
            Entity="Product Ad",
            Operation="Create",
            **{"Campaign ID": campaign_id},
            **{"Ad Group ID": ad_group_id},
            SKU=sku,
            **{"Ad Status": "enabled"},
        )

    def add_keyword(self, campaign_id: int, ad_group_id: int,
                    keyword_text: str, match_type: str, bid: Optional[float] = None) -> None:
        """Add a Keyword row."""
        row_data = {
            "Product": "Sponsored Products",
            "Entity": "Keyword",
            "Operation": "Create",
            "Campaign ID": campaign_id,
            "Ad Group ID": ad_group_id,
            "Keyword Text": keyword_text,
            "Match Type": match_type,
            "Keyword Status": "enabled",
        }

        if bid is not None:
            row_data["Bid"] = bid

        self._add_row(**row_data)

    def add_product_targeting(self, campaign_id: int, ad_group_id: int,
                              expression: str, bid: Optional[float] = None) -> None:
        """Add a Product Targeting row."""
        row_data = {
            "Product": "Sponsored Products",
            "Entity": "Product Targeting",
            "Operation": "Create",
            "Campaign ID": campaign_id,
            "Ad Group ID": ad_group_id,
            "Product Targeting Expression": expression,
            "Product Targeting Status": "enabled",
        }

        if bid is not None:
            row_data["Product Targeting Bid"] = bid

        self._add_row(**row_data)

    def add_campaign_with_placements(self, name: str, daily_budget: float,
                                     targeting_type: str) -> int:
        """
        Add a Campaign row together with its standard placement adjustments.

        This is a convenience wrapper that creates the campaign and appends
        the two default bidding-adjustment rows (Top-of-Search and Product
        Page) so the pattern is not repeated for every campaign type.

        Returns:
            Campaign ID (negative integer)
        """
        campaign_id = self.add_campaign(
            name=name, daily_budget=daily_budget,
            targeting_type=targeting_type,
        )
        self.add_bidding_adjustment(
            campaign_id, "placementTop",
            self.config.placement_top_percentage,
        )
        self.add_bidding_adjustment(
            campaign_id, "placementProductPage",
            self.config.placement_product_page_percentage,
        )
        return campaign_id

    def generate_clearance_campaigns(self, tier: str, skus: list,
                                     competitor_asins: Optional[list] = None,
                                     listing_info: Optional[dict] = None) -> None:
        """
        Generate the full set of clearance campaigns for a tier.

        Args:
            tier: Tier name (e.g., "365+" or "211-330")
            skus: List of SKUs for this tier
            competitor_asins: Optional list of competitor ASINs for product targeting
            listing_info: Optional dict mapping SKU to listing info
        """
        if not skus:
            return

        prefix = self.config.campaign_prefix

        # Determine budget based on tier
        if "365" in tier:
            budget = self.config.daily_budget_365_plus
        else:
            budget = self.config.daily_budget_211_330

        # =================================================================
        # CAMPAIGN 1: Manual Keywords
        # =================================================================
        manual_kw_campaign_id = self.add_campaign_with_placements(
            name=f"{prefix} - {tier} - Manual Keywords",
            daily_budget=budget,
            targeting_type="MANUAL",
        )

        # Create ad groups and keywords for each SKU
        for sku in skus:
            ad_group_id = self.add_ad_group(
                manual_kw_campaign_id,
                name=f"{sku} - Keywords",
                default_bid=self.config.default_bid
            )

            # Add product ad
            self.add_product_ad(manual_kw_campaign_id, ad_group_id, sku)

            # Add generic keywords
            for kw in self.config.generic_keywords:
                for match_type in ['exact', 'phrase', 'broad']:
                    self.add_keyword(
                        manual_kw_campaign_id, ad_group_id,
                        kw, match_type
                    )

            # Add size-specific keywords if listing info available
            if listing_info and sku in listing_info:
                info = listing_info[sku]
                item_name = info.get('item-name', '')

                # Extract size from item name or use size tokens
                for size in self.config.size_tokens:
                    if size.lower() in item_name.lower():
                        self.add_keyword(
                            manual_kw_campaign_id, ad_group_id,
                            f"mylar bags {size}", "exact"
                        )
                        self.add_keyword(
                            manual_kw_campaign_id, ad_group_id,
                            f"{size} mylar bags", "exact"
                        )

            # Add pack count keywords
            for pack in self.config.pack_counts:
                self.add_keyword(
                    manual_kw_campaign_id, ad_group_id,
                    f"mylar bags {pack}", "phrase"
                )

        # =================================================================
        # CAMPAIGN 2: Auto
        # =================================================================
        auto_campaign_id = self.add_campaign_with_placements(
            name=f"{prefix} - {tier} - Auto",
            daily_budget=budget,
            targeting_type="AUTO",
        )

        # Single ad group for auto campaign with all SKUs
        auto_ad_group_id = self.add_ad_group(
            auto_campaign_id,
            name=f"{tier} - Auto Targeting",
            default_bid=self.config.default_bid
        )

        # Add all SKUs as product ads
        for sku in skus:
            self.add_product_ad(auto_campaign_id, auto_ad_group_id, sku)

        # Add auto targeting group targets
        for target_group in AUTO_TARGETING_GROUPS:
            self.add_product_targeting(
                auto_campaign_id, auto_ad_group_id,
                target_group
            )

        # =================================================================
        # CAMPAIGN 3: Manual Product Targeting
        # =================================================================
        manual_pt_campaign_id = self.add_campaign_with_placements(
            name=f"{prefix} - {tier} - Product Targeting",
            daily_budget=budget,
            targeting_type="MANUAL",
        )

        # Create ad group for product targeting
        pt_ad_group_id = self.add_ad_group(
            manual_pt_campaign_id,
            name=f"{tier} - Product Targeting",
            default_bid=self.config.default_bid
        )

        # Add all SKUs as product ads
        for sku in skus:
            self.add_product_ad(manual_pt_campaign_id, pt_ad_group_id, sku)

        # Add competitor ASIN targets
        if competitor_asins:
            for comp_asin in competitor_asins:
                self.add_product_targeting(
                    manual_pt_campaign_id, pt_ad_group_id,
                    f'asin="{comp_asin}"'
                )
        else:
            logger.warning(
                "No competitor ASINs provided for %s Product Targeting "
                "campaign - only SKU ads will be created (no targets)", tier
            )

    def get_rows(self) -> list:
        """Get all generated rows."""
        return self.rows


# =============================================================================
# XLSX WRITER
# =============================================================================

def write_bulk_xlsx(rows: list, headers: list, output_path: str) -> None:
    """
    Write bulk campaign data to XLSX file.

    Args:
        rows: List of row dicts
        headers: Column headers
        output_path: Output file path
    """
    wb = Workbook()

    # Remove default sheet and create our sheet
    default_sheet = wb.active
    wb.remove(default_sheet)

    ws = wb.create_sheet(title=SHEET_NAME)

    # Write headers
    for col_idx, header in enumerate(headers, 1):
        ws.cell(row=1, column=col_idx, value=header)

    # Write data rows
    for row_idx, row_data in enumerate(rows, 2):
        for col_idx, header in enumerate(headers, 1):
            value = row_data.get(header, '')
            # Ensure blanks are truly blank (critical!)
            if value == '' or value is None:
                ws.cell(row=row_idx, column=col_idx, value=None)
            else:
                ws.cell(row=row_idx, column=col_idx, value=value)

    # Save
    # Ensure output directory exists
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    wb.save(output_path)
    logger.info("Wrote %d rows to %s", len(rows), output_path)


# =============================================================================
# MAIN WORKFLOW
# =============================================================================

def _sync_listings_to_supabase(store: SupabaseStore, listings_df: pd.DataFrame) -> None:
    """
    Sync parsed listings report data into Supabase for persistence.

    Args:
        store: Active SupabaseStore instance
        listings_df: DataFrame from Active Listings Report
    """
    # Find column names
    asin_col = None
    for col in ['asin1', 'ASIN1', 'asin', 'ASIN']:
        if col in listings_df.columns:
            asin_col = col
            break

    sku_col = None
    for col in ['seller-sku', 'Seller SKU', 'sku', 'SKU']:
        if col in listings_df.columns:
            sku_col = col
            break

    if not asin_col or not sku_col:
        logger.warning("Could not find ASIN/SKU columns for Supabase sync")
        return

    # Build listing dicts
    field_map = {
        'item_name': ['item-name', 'Item Name', 'title', 'Title'],
        'fulfillment_channel': ['fulfillment-channel', 'Fulfillment Channel'],
        'price': ['price', 'Price'],
        'quantity': ['quantity', 'Quantity'],
    }

    listings_data = []
    for _, row in listings_df.iterrows():
        asin = str(row[asin_col]).strip()
        sku = str(row[sku_col]).strip()
        if not asin or not sku:
            continue

        item = {"asin": asin, "seller_sku": sku}
        for field_key, possible_cols in field_map.items():
            for col in possible_cols:
                if col in row.index and str(row[col]).strip():
                    val = str(row[col]).strip()
                    if field_key == "price":
                        try:
                            item[field_key] = float(val)
                        except ValueError:
                            pass
                    elif field_key == "quantity":
                        try:
                            item[field_key] = int(val)
                        except ValueError:
                            pass
                    else:
                        item[field_key] = val
                    break

        listings_data.append(item)

    if listings_data:
        result = store.sync_listings(listings_data)
        err_count = result['errors']
        err_msg = f", {err_count} errors" if err_count else ""
        logger.info("Supabase: synced %d listings%s",
                     result['upserted'], err_msg)


def _count_campaigns(rows: list) -> int:
    """Count Campaign entity rows."""
    return sum(1 for r in rows if r.get('Entity') == 'Campaign')


def generate_bulk_upload(
    target_asins: list,
    listings_report_path: str,
    output_path: str,
    template_path: Optional[str] = None,
    competitor_asins: Optional[list] = None,
    config: Optional[CampaignConfig] = None,
    tier_assignment: Optional[dict] = None,
    store: Optional[SupabaseStore] = None,
    delimiter: Optional[str] = None,
) -> ValidationResult:
    """
    Main function to generate the bulk upload file.

    Args:
        target_asins: List of ASINs to create campaigns for
        listings_report_path: Path to Active Listings Report
        output_path: Output XLSX file path
        template_path: Optional path to Amazon Ads bulk template
        competitor_asins: Optional list of competitor ASINs
        config: Optional CampaignConfig (uses defaults if not provided)
        tier_assignment: Optional dict mapping ASIN to tier ("365+" or "211-330")
        store: Optional SupabaseStore for persistence
        delimiter: Optional explicit delimiter for listings report

    Returns:
        ValidationResult with any errors/warnings
    """
    if config is None:
        config = CampaignConfig()

    logger.info("=" * 60)
    logger.info("Amazon Sponsored Products Bulk Campaign Generator")
    logger.info("=" * 60)

    if store:
        logger.info("[Supabase] Connected - persistence enabled")

    # Step 1: Read Active Listings Report
    logger.info("[1/6] Reading Active Listings Report: %s",
                listings_report_path)
    listings_df = read_active_listings_report(listings_report_path,
                                              delimiter=delimiter)
    asin_sku_map = build_asin_sku_mapping(listings_df)
    all_listing_info = build_listing_info_map(listings_df)
    logger.info("Found %d ASIN->SKU mappings", len(asin_sku_map))

    # Step 1b: Sync listings to Supabase if connected
    if store:
        logger.info("[2/6] Syncing listings to Supabase")
        _sync_listings_to_supabase(store, listings_df)
    else:
        logger.info("[2/6] Supabase not configured, skipping sync")

    # Step 2: Validate target ASINs
    logger.info("[3/6] Validating %d target ASINs", len(target_asins))
    missing_asins = []
    valid_asin_skus = {}
    listing_info_map = {}

    for asin in target_asins:
        if asin in asin_sku_map:
            sku = asin_sku_map[asin]
            valid_asin_skus[asin] = sku
            listing_info_map[sku] = all_listing_info.get(asin, {})
            logger.info("  + %s -> %s", asin, sku)
        else:
            # Fallback: try Supabase if the ASIN isn't in the local file
            if store:
                sku = store.get_sku_for_asin(asin)
                if sku:
                    valid_asin_skus[asin] = sku
                    logger.info("  + %s -> %s  (from Supabase)", asin, sku)
                    continue

            missing_asins.append(asin)
            logger.warning("  x %s - NOT FOUND", asin)

    if missing_asins:
        sample = missing_asins[:5]
        sample_str = ", ".join(sample)
        extra = (f" (and {len(missing_asins) - 5} more)"
                 if len(missing_asins) > 5 else "")
        logger.warning(
            "%d of %d ASINs not found in listings report%s: %s%s. "
            "Check that your Active Listings Report is current.",
            len(missing_asins), len(target_asins),
            " or Supabase" if store else "",
            sample_str, extra,
        )

    if not valid_asin_skus:
        return ValidationResult(
            is_valid=False,
            errors=["No valid ASINs found in Active Listings Report"]
        )

    # Step 3: Read template headers
    logger.info("[4/6] Reading template headers")
    headers = read_template_headers(template_path)
    logger.info("Using %d columns", len(headers))

    # Step 4: Generate campaign data
    logger.info("[5/6] Generating campaign data")
    generator = BulkCampaignGenerator(config, headers)

    # Group SKUs by tier
    tier_365_skus = []
    tier_211_skus = []

    for asin, sku in valid_asin_skus.items():
        if tier_assignment and asin in tier_assignment:
            tier = tier_assignment[asin]
        else:
            # Default all to 365+ if no tier assignment
            tier = "365+"

        if "365" in tier:
            tier_365_skus.append(sku)
        else:
            tier_211_skus.append(sku)

    # Generate campaigns for each tier
    if tier_365_skus:
        logger.info("Generating 365+ tier campaigns for %d SKUs",
                     len(tier_365_skus))
        generator.generate_clearance_campaigns(
            "365+", tier_365_skus,
            competitor_asins=competitor_asins,
            listing_info=listing_info_map
        )

    if tier_211_skus:
        logger.info("Generating 211-330 tier campaigns for %d SKUs",
                     len(tier_211_skus))
        generator.generate_clearance_campaigns(
            "211-330", tier_211_skus,
            competitor_asins=competitor_asins,
            listing_info=listing_info_map
        )

    rows = generator.get_rows()
    logger.info("Generated %d total rows", len(rows))

    # Step 5: Validate and write
    logger.info("[6/6] Validating and writing output")
    validation = validate_all_rows(rows)

    if not validation.is_valid:
        logger.error("Validation FAILED:")
        for error in validation.errors:
            logger.error("  %s", error)
        return validation

    if validation.warnings:
        for warning in validation.warnings:
            logger.warning("  %s", warning)

    write_bulk_xlsx(rows, headers, output_path)

    # Step 6: Record run and upload file to Supabase
    if store:
        logger.info("[Supabase] Recording generation run...")
        campaign_count = _count_campaigns(rows)

        run_id = store.record_generation_run(
            target_asins=target_asins,
            matched_skus=valid_asin_skus,
            missing_asins=missing_asins,
            config=asdict(config),
            row_count=len(rows),
            campaign_count=campaign_count,
            tier_assignment=tier_assignment,
            competitor_asins=competitor_asins,
            file_name=Path(output_path).name,
        )
        logger.info("[Supabase] Saved as run #%s", run_id)

        try:
            storage_path = store.upload_file(output_path)
            # Update the run with the storage path
            store.client.table("generation_runs").update(
                {"file_storage_path": storage_path}
            ).eq("id", run_id).execute()
            logger.info("[Supabase] File uploaded to storage: %s",
                        storage_path)
        except Exception as e:
            logger.warning("[Supabase] File upload skipped: %s", e)

    logger.info("=" * 60)
    logger.info("SUCCESS! Upload file ready.")
    logger.info("=" * 60)

    return validation


# =============================================================================
# CLI
# =============================================================================

def _validate_asin_format(asin: str) -> bool:
    """Check that an ASIN looks valid (10 alphanumeric characters)."""
    return bool(re.match(r'^[A-Z0-9]{10}$', asin))


def main():
    parser = argparse.ArgumentParser(
        description="Generate Amazon Sponsored Products bulk upload file"
    )

    parser.add_argument(
        "--asins",
        required=True,
        help="Comma-separated list of target ASINs or path to file with one ASIN per line"
    )

    parser.add_argument(
        "--listings-report",
        required=True,
        help="Path to Seller Central Active Listings Report"
    )

    parser.add_argument(
        "--output",
        default="bulk_upload.xlsx",
        help="Output XLSX file path (default: bulk_upload.xlsx)"
    )

    parser.add_argument(
        "--template",
        help="Path to Amazon Ads bulk operations template XLSX"
    )

    parser.add_argument(
        "--competitor-asins",
        help="Comma-separated list of competitor ASINs for product targeting"
    )

    parser.add_argument(
        "--daily-budget-365",
        type=float,
        default=50.0,
        help="Daily budget for 365+ tier campaigns (default: 50.0)"
    )

    parser.add_argument(
        "--daily-budget-211",
        type=float,
        default=30.0,
        help="Daily budget for 211-330 tier campaigns (default: 30.0)"
    )

    parser.add_argument(
        "--default-bid",
        type=float,
        default=0.75,
        help="Default bid for ad groups (default: 0.75)"
    )

    parser.add_argument(
        "--bidding-strategy",
        default="Dynamic bids - down only",
        choices=list(BIDDING_STRATEGIES),
        help="Campaign bidding strategy"
    )

    parser.add_argument(
        "--config",
        help="Path to JSON config file for advanced settings"
    )

    parser.add_argument(
        "--tier-file",
        help="Path to JSON file mapping ASIN to tier (e.g., {\"B08NV6CLGF\": \"365+\"})"
    )

    parser.add_argument(
        "--no-supabase",
        action="store_true",
        help="Disable Supabase integration even if credentials are configured"
    )

    parser.add_argument(
        "--delimiter",
        help="Explicit delimiter for the Active Listings Report "
             "(e.g., '\\t' for tab, ',' for comma). Auto-detected if omitted."
    )

    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable verbose (DEBUG-level) logging output"
    )

    args = parser.parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(
        level=log_level,
        format="%(levelname)s: %(message)s",
    )

    # Parse ASINs
    if Path(args.asins).exists():
        with open(args.asins) as f:
            target_asins = [line.strip() for line in f if line.strip()]
    else:
        target_asins = [a.strip() for a in args.asins.split(',') if a.strip()]

    # Validate ASIN format early
    invalid_format = [a for a in target_asins if not _validate_asin_format(a)]
    if invalid_format:
        sample = invalid_format[:5]
        logger.warning(
            "%d ASIN(s) have unexpected format (expected 10 alphanumeric "
            "chars): %s. They will still be looked up, but may not match.",
            len(invalid_format), ", ".join(sample),
        )

    # Parse competitor ASINs
    competitor_asins = None
    if args.competitor_asins:
        competitor_asins = [a.strip() for a in args.competitor_asins.split(',')
                           if a.strip()]

    # Build config
    if args.config and Path(args.config).exists():
        with open(args.config) as f:
            config_dict = json.load(f)
        config = CampaignConfig.from_dict(config_dict)
    else:
        config = CampaignConfig(
            daily_budget_365_plus=args.daily_budget_365,
            daily_budget_211_330=args.daily_budget_211,
            default_bid=args.default_bid,
            bidding_strategy=args.bidding_strategy,
        )

    # Parse delimiter (handle escape sequences like \t)
    delimiter = None
    if args.delimiter:
        delimiter = args.delimiter.encode().decode('unicode_escape')

    # Parse tier assignment
    tier_assignment = None
    if args.tier_file and Path(args.tier_file).exists():
        with open(args.tier_file) as f:
            tier_assignment = json.load(f)

    # Initialize Supabase (optional - silently skipped if not configured)
    store = None
    if not args.no_supabase:
        store = get_store()

    # Run generation
    result = generate_bulk_upload(
        target_asins=target_asins,
        listings_report_path=args.listings_report,
        output_path=args.output,
        template_path=args.template,
        competitor_asins=competitor_asins,
        config=config,
        tier_assignment=tier_assignment,
        store=store,
        delimiter=delimiter,
    )

    if not result.is_valid:
        sys.exit(1)


if __name__ == "__main__":
    main()
