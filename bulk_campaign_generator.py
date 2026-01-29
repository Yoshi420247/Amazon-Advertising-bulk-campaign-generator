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
import json
import sys
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook, load_workbook


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

AUTO_TARGETING_GROUPS = {"close-match", "loose-match", "substitutes", "complements"}

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


@dataclass
class ValidationResult:
    """Result of validation."""
    is_valid: bool
    errors: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


# =============================================================================
# ACTIVE LISTINGS REPORT PARSER
# =============================================================================

def read_active_listings_report(filepath: str) -> pd.DataFrame:
    """
    Read Seller Central Active Listings Report.

    Attempts tab-delimited first, then falls back to CSV.

    Args:
        filepath: Path to the Active Listings Report file

    Returns:
        DataFrame with listing data
    """
    path = Path(filepath)

    if not path.exists():
        raise FileNotFoundError(f"Active Listings Report not found: {filepath}")

    # Try tab-delimited first
    try:
        df = pd.read_csv(filepath, sep='\t', dtype=str, keep_default_na=False)
        if len(df.columns) > 1:
            return df
    except Exception:
        pass

    # Fall back to CSV
    try:
        df = pd.read_csv(filepath, dtype=str, keep_default_na=False)
        return df
    except Exception as e:
        raise ValueError(f"Could not parse Active Listings Report: {e}")


def build_asin_sku_mapping(listings_df: pd.DataFrame) -> dict:
    """
    Build ASIN → SKU mapping from Active Listings Report.

    Args:
        listings_df: DataFrame from Active Listings Report

    Returns:
        Dict mapping ASIN to seller-sku
    """
    # Find the ASIN column (may be 'asin1' or 'asin')
    asin_col = None
    for col in ['asin1', 'ASIN1', 'asin', 'ASIN']:
        if col in listings_df.columns:
            asin_col = col
            break

    if asin_col is None:
        raise ValueError("Could not find ASIN column in Active Listings Report")

    # Find SKU column
    sku_col = None
    for col in ['seller-sku', 'Seller SKU', 'sku', 'SKU']:
        if col in listings_df.columns:
            sku_col = col
            break

    if sku_col is None:
        raise ValueError("Could not find SKU column in Active Listings Report")

    # Build mapping
    mapping = {}
    for _, row in listings_df.iterrows():
        asin = str(row[asin_col]).strip()
        sku = str(row[sku_col]).strip()
        if asin and sku:
            mapping[asin] = sku

    return mapping


def get_listing_info(listings_df: pd.DataFrame, asin: str) -> dict:
    """
    Get additional listing info for an ASIN.

    Args:
        listings_df: DataFrame from Active Listings Report
        asin: ASIN to look up

    Returns:
        Dict with listing info (item-name, fulfillment-channel, price, etc.)
    """
    # Find the ASIN column
    asin_col = None
    for col in ['asin1', 'ASIN1', 'asin', 'ASIN']:
        if col in listings_df.columns:
            asin_col = col
            break

    if asin_col is None:
        return {}

    row = listings_df[listings_df[asin_col] == asin]
    if row.empty:
        return {}

    row = row.iloc[0]
    info = {}

    # Extract common fields
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
        print(f"Warning: Template not found at {template_path}, using default columns")
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
            print("Warning: No headers found in template, using default columns")
            return DEFAULT_COLUMNS.copy()

    except Exception as e:
        print(f"Warning: Could not read template ({e}), using default columns")
        return DEFAULT_COLUMNS.copy()


# =============================================================================
# VALIDATION
# =============================================================================

def validate_campaign_row(row: dict) -> ValidationResult:
    """Validate a Campaign entity row."""
    errors = []
    warnings = []

    required = ['Campaign ID', 'Campaign Name', 'Campaign Daily Budget',
                'Campaign Targeting Type', 'Campaign Status', 'Campaign Start Date',
                'Campaign Bidding Strategy']

    for field in required:
        val = row.get(field, '')
        if val == '' or val is None:
            errors.append(f"Campaign missing required field: {field}")

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

    return ValidationResult(is_valid=len(errors) == 0, errors=errors, warnings=warnings)


def validate_bidding_adjustment_row(row: dict) -> ValidationResult:
    """Validate a Bidding Adjustment entity row."""
    errors = []

    required = ['Campaign ID', 'Placement', 'Percentage']

    for field in required:
        val = row.get(field, '')
        if val == '' or val is None:
            errors.append(f"Bidding Adjustment missing required field: {field}")

    placement = row.get('Placement', '')
    if placement and placement not in PLACEMENTS:
        errors.append(f"Invalid Placement: {placement}")

    return ValidationResult(is_valid=len(errors) == 0, errors=errors)


def validate_ad_group_row(row: dict) -> ValidationResult:
    """Validate an Ad Group entity row."""
    errors = []

    required = ['Campaign ID', 'Ad Group ID', 'Ad Group Name',
                'Ad Group Default Bid', 'Ad Group Status']

    for field in required:
        val = row.get(field, '')
        if val == '' or val is None:
            errors.append(f"Ad Group missing required field: {field}")

    return ValidationResult(is_valid=len(errors) == 0, errors=errors)


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

    required = ['Campaign ID', 'Ad Group ID', 'Keyword Text', 'Match Type', 'Keyword Status']

    for field in required:
        val = row.get(field, '')
        if val == '' or val is None:
            errors.append(f"Keyword missing required field: {field}")

    match_type = row.get('Match Type', '')
    if match_type and match_type not in MATCH_TYPES:
        errors.append(f"Invalid Match Type: {match_type}")

    return ValidationResult(is_valid=len(errors) == 0, errors=errors)


def validate_product_targeting_row(row: dict) -> ValidationResult:
    """Validate a Product Targeting entity row."""
    errors = []

    required = ['Campaign ID', 'Ad Group ID', 'Product Targeting Expression',
                'Product Targeting Status']

    for field in required:
        val = row.get(field, '')
        if val == '' or val is None:
            errors.append(f"Product Targeting missing required field: {field}")

    return ValidationResult(is_valid=len(errors) == 0, errors=errors)


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
        manual_kw_campaign_id = self.add_campaign(
            name=f"{prefix} - {tier} - Manual Keywords",
            daily_budget=budget,
            targeting_type="MANUAL"
        )

        # Add placement adjustments
        self.add_bidding_adjustment(
            manual_kw_campaign_id, "placementTop",
            self.config.placement_top_percentage
        )
        self.add_bidding_adjustment(
            manual_kw_campaign_id, "placementProductPage",
            self.config.placement_product_page_percentage
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
        auto_campaign_id = self.add_campaign(
            name=f"{prefix} - {tier} - Auto",
            daily_budget=budget,
            targeting_type="AUTO"
        )

        # Add placement adjustments
        self.add_bidding_adjustment(
            auto_campaign_id, "placementTop",
            self.config.placement_top_percentage
        )
        self.add_bidding_adjustment(
            auto_campaign_id, "placementProductPage",
            self.config.placement_product_page_percentage
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
        manual_pt_campaign_id = self.add_campaign(
            name=f"{prefix} - {tier} - Product Targeting",
            daily_budget=budget,
            targeting_type="MANUAL"
        )

        # Add placement adjustments
        self.add_bidding_adjustment(
            manual_pt_campaign_id, "placementTop",
            self.config.placement_top_percentage
        )
        self.add_bidding_adjustment(
            manual_pt_campaign_id, "placementProductPage",
            self.config.placement_product_page_percentage
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
    wb.save(output_path)
    print(f"Wrote {len(rows)} rows to {output_path}")


# =============================================================================
# MAIN WORKFLOW
# =============================================================================

def generate_bulk_upload(
    target_asins: list,
    listings_report_path: str,
    output_path: str,
    template_path: Optional[str] = None,
    competitor_asins: Optional[list] = None,
    config: Optional[CampaignConfig] = None,
    tier_assignment: Optional[dict] = None
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

    Returns:
        ValidationResult with any errors/warnings
    """
    if config is None:
        config = CampaignConfig()

    print("=" * 60)
    print("Amazon Sponsored Products Bulk Campaign Generator")
    print("=" * 60)

    # Step 1: Read Active Listings Report
    print(f"\n[1/5] Reading Active Listings Report: {listings_report_path}")
    listings_df = read_active_listings_report(listings_report_path)
    asin_sku_map = build_asin_sku_mapping(listings_df)
    print(f"      Found {len(asin_sku_map)} ASIN→SKU mappings")

    # Step 2: Validate target ASINs
    print(f"\n[2/5] Validating {len(target_asins)} target ASINs")
    missing_asins = []
    valid_asin_skus = {}
    listing_info_map = {}

    for asin in target_asins:
        if asin in asin_sku_map:
            sku = asin_sku_map[asin]
            valid_asin_skus[asin] = sku
            listing_info_map[sku] = get_listing_info(listings_df, asin)
            print(f"      ✓ {asin} → {sku}")
        else:
            missing_asins.append(asin)
            print(f"      ✗ {asin} - NOT FOUND")

    if missing_asins:
        print(f"\nWARNING: {len(missing_asins)} ASINs not found in listings report:")
        for asin in missing_asins:
            print(f"  - {asin}")

    if not valid_asin_skus:
        return ValidationResult(
            is_valid=False,
            errors=["No valid ASINs found in Active Listings Report"]
        )

    # Step 3: Read template headers
    print(f"\n[3/5] Reading template headers")
    headers = read_template_headers(template_path)
    print(f"      Using {len(headers)} columns")

    # Step 4: Generate campaign data
    print(f"\n[4/5] Generating campaign data")
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
        print(f"      Generating 365+ tier campaigns for {len(tier_365_skus)} SKUs")
        generator.generate_clearance_campaigns(
            "365+", tier_365_skus,
            competitor_asins=competitor_asins,
            listing_info=listing_info_map
        )

    if tier_211_skus:
        print(f"      Generating 211-330 tier campaigns for {len(tier_211_skus)} SKUs")
        generator.generate_clearance_campaigns(
            "211-330", tier_211_skus,
            competitor_asins=competitor_asins,
            listing_info=listing_info_map
        )

    rows = generator.get_rows()
    print(f"      Generated {len(rows)} total rows")

    # Step 5: Validate and write
    print(f"\n[5/5] Validating and writing output")
    validation = validate_all_rows(rows)

    if not validation.is_valid:
        print("\nValidation FAILED:")
        for error in validation.errors:
            print(f"  ERROR: {error}")
        return validation

    if validation.warnings:
        print("\nWarnings:")
        for warning in validation.warnings:
            print(f"  WARNING: {warning}")

    write_bulk_xlsx(rows, headers, output_path)

    print("\n" + "=" * 60)
    print("SUCCESS! Upload file ready.")
    print("=" * 60)

    return validation


# =============================================================================
# CLI
# =============================================================================

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

    args = parser.parse_args()

    # Parse ASINs
    if Path(args.asins).exists():
        with open(args.asins) as f:
            target_asins = [line.strip() for line in f if line.strip()]
    else:
        target_asins = [a.strip() for a in args.asins.split(',') if a.strip()]

    # Parse competitor ASINs
    competitor_asins = None
    if args.competitor_asins:
        competitor_asins = [a.strip() for a in args.competitor_asins.split(',') if a.strip()]

    # Build config
    if args.config and Path(args.config).exists():
        with open(args.config) as f:
            config_dict = json.load(f)
        config = CampaignConfig(**config_dict)
    else:
        config = CampaignConfig(
            daily_budget_365_plus=args.daily_budget_365,
            daily_budget_211_330=args.daily_budget_211,
            default_bid=args.default_bid,
            bidding_strategy=args.bidding_strategy,
        )

    # Parse tier assignment
    tier_assignment = None
    if args.tier_file and Path(args.tier_file).exists():
        with open(args.tier_file) as f:
            tier_assignment = json.load(f)

    # Run generation
    result = generate_bulk_upload(
        target_asins=target_asins,
        listings_report_path=args.listings_report,
        output_path=args.output,
        template_path=args.template,
        competitor_asins=competitor_asins,
        config=config,
        tier_assignment=tier_assignment,
    )

    if not result.is_valid:
        sys.exit(1)


if __name__ == "__main__":
    main()
