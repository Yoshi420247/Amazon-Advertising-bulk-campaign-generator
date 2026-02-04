# Amazon Sponsored Products Bulk Campaign Generator

Generate upload-ready XLSX files for Amazon Ads Sponsored Products bulk operations.

## Overview

This tool automates the creation of Amazon Sponsored Products campaigns by:
1. Reading target ASINs (e.g., clearance inventory)
2. Looking up Seller SKUs from your Active Listings Report
3. Generating a properly formatted bulk upload spreadsheet

## Requirements

- Python 3.8+
- pandas
- openpyxl
- supabase (optional, for persistence)

Install dependencies:
```bash
pip install -r requirements.txt

# Optional: enable Supabase integration
pip install supabase
```

## Quick Start

```bash
python bulk_campaign_generator.py \
    --asins "B08NV6CLGF,B09ABC1234,B09DEF5678" \
    --listings-report path/to/active_listings_report.txt \
    --output clearance_campaigns.xlsx
```

## Usage

### Basic Usage

```bash
python bulk_campaign_generator.py \
    --asins samples/sample_asins.txt \
    --listings-report samples/sample_active_listings.txt \
    --output output/clearance_campaigns.xlsx
```

### With All Options

```bash
python bulk_campaign_generator.py \
    --asins samples/sample_asins.txt \
    --listings-report samples/sample_active_listings.txt \
    --output output/clearance_campaigns.xlsx \
    --template path/to/amazon_template.xlsx \
    --competitor-asins "B07ABC1234,B07DEF5678" \
    --daily-budget-365 50.0 \
    --daily-budget-211 30.0 \
    --default-bid 0.75 \
    --bidding-strategy "Dynamic bids - down only" \
    --tier-file samples/sample_tier_assignment.json \
    --config samples/sample_config.json
```

### Command Line Arguments

| Argument | Required | Description |
|----------|----------|-------------|
| `--asins` | Yes | Comma-separated ASINs or path to file (one per line) |
| `--listings-report` | Yes | Path to Seller Central Active Listings Report |
| `--output` | No | Output XLSX file (default: `bulk_upload.xlsx`) |
| `--template` | No | Amazon Ads bulk template for column headers |
| `--competitor-asins` | No | Competitor ASINs for product targeting |
| `--daily-budget-365` | No | Budget for 365+ tier (default: 50.0) |
| `--daily-budget-211` | No | Budget for 211-330 tier (default: 30.0) |
| `--default-bid` | No | Default ad group bid (default: 0.75) |
| `--bidding-strategy` | No | Campaign bidding strategy |
| `--tier-file` | No | JSON file mapping ASIN to tier |
| `--config` | No | JSON config file for all settings |
| `--no-supabase` | No | Disable Supabase even if credentials are set |

## Input Files

### Active Listings Report (Required)

Download from Seller Central → Reports → Business Reports → Active Listings Report

The file must contain:
- `seller-sku` or `SKU` column
- `asin1` or `ASIN` column

Supported formats: Tab-delimited (.txt) or CSV (.csv)

### Target ASINs

Either:
- Comma-separated string: `"B08NV6CLGF,B09ABC1234"`
- Text file with one ASIN per line

### Tier Assignment (Optional)

JSON file mapping ASINs to inventory tiers:
```json
{
  "B08NV6CLGF": "365+",
  "B09ABC1234": "211-330"
}
```

### Configuration File (Optional)

JSON file with campaign settings:
```json
{
  "daily_budget_365_plus": 50.0,
  "daily_budget_211_330": 30.0,
  "default_bid": 0.75,
  "bidding_strategy": "Dynamic bids - down only",
  "placement_top_percentage": 80,
  "placement_product_page_percentage": 50,
  "campaign_prefix": "Clearance",
  "generic_keywords": ["mylar bags", "food storage bags"],
  "size_tokens": ["3x4.5", "4x6", "5x7"],
  "pack_counts": ["100 pack", "1000 pack"]
}
```

## Output

The tool generates an XLSX file with a single sheet named `Sponsored Products Campaigns` containing:

- **Campaign** rows with negative IDs
- **Bidding Adjustment** rows for placement multipliers
- **Ad Group** rows with negative IDs
- **Product Ad** rows using Seller SKUs
- **Keyword** rows for manual campaigns
- **Product Targeting** rows for auto and competitor targeting

### Campaign Structure

For each tier (365+ and 211-330), the tool creates:

1. **Manual Keywords Campaign** - Ad groups per SKU with size/pack keywords
2. **Auto Campaign** - Single ad group with all SKUs
3. **Product Targeting Campaign** - Competitor ASIN targeting

## Validation

The tool validates:
- All target ASINs exist in the Active Listings Report
- Required fields are populated for each entity type
- Enum values match Amazon's allowed values
- No `0` values in fields that must be blank

## Uploading to Amazon Ads

1. Go to Amazon Ads Console
2. Navigate to Sponsored Products → Bulk Operations
3. Upload the generated `.xlsx` file
4. Review and confirm the operations

### Troubleshooting Upload Errors

- **Fix parent rows first**: If a Campaign fails, all child rows show "Not Processed"
- **Don't reupload processing reports**: They contain extra columns and invalid defaults
- **Check SKU column**: Ensure SKUs weren't reformatted by Excel

## Supabase Integration (Optional)

When configured, Supabase adds persistent storage and tracking:

- **Listings table** - ASIN/SKU mappings survive between runs. If an ASIN is missing from your local listings report, the generator falls back to the Supabase copy.
- **Generation history** - Every run is recorded with its config, target ASINs, and results. Prevents accidental duplicate campaigns.
- **Upload tracking** - Record whether your Amazon Ads upload succeeded or had errors, linked to the generation run.
- **File storage** - Generated XLSX files are uploaded to a private Supabase Storage bucket for access from anywhere.

### Setup

1. Create a Supabase project at [supabase.com](https://supabase.com)

2. Run the schema migration in the SQL Editor (Dashboard -> SQL Editor -> New Query):
   ```
   -- paste contents of supabase_schema.sql
   ```

3. Install the Python client:
   ```bash
   pip install supabase
   ```

4. Set environment variables (copy `.env.example` to `.env`):
   ```bash
   cp .env.example .env
   # Edit .env with your project URL and API key from:
   # Dashboard -> Settings -> API
   ```

   Or export directly:
   ```bash
   export SUPABASE_URL=https://your-project-id.supabase.co
   export SUPABASE_KEY=your-anon-key
   ```

5. Run the generator normally. It detects Supabase automatically:
   ```bash
   python bulk_campaign_generator.py \
       --asins samples/sample_asins.txt \
       --listings-report samples/sample_active_listings.txt \
       --output clearance_campaigns.xlsx
   ```

   You'll see `[Supabase] Connected - persistence enabled` in the output.

### What gets stored

| Table | Purpose |
|-------|---------|
| `listings` | ASIN, SKU, item name, price, fulfillment channel, active/inactive status |
| `generation_runs` | Timestamp, ASINs, config, row count, file reference |
| `upload_results` | Linked to a run: status, processed/failed row counts, error details |
| `campaign-files` bucket | The actual XLSX files |

### Disabling Supabase

If credentials are configured but you want a local-only run:
```bash
python bulk_campaign_generator.py --no-supabase ...
```

Without `SUPABASE_URL` and `SUPABASE_KEY` set, the generator runs fully offline with no warnings.

## License

MIT
