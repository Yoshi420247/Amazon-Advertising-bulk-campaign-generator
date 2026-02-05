# CLAUDE.md - Project Guide for AI Assistants

## Project Overview

Amazon Sponsored Products Bulk Campaign Generator - a Python CLI tool that generates upload-ready XLSX files for Amazon Ads bulk operations. It reads target ASINs and Active Listings Reports, then creates properly formatted spreadsheets with campaigns, ad groups, keywords, and product targeting rows.

## Tech Stack

- **Language:** Python 3.8+
- **Core deps:** pandas, openpyxl
- **Optional:** supabase-py (for persistent storage)

## Project Structure

```
bulk_campaign_generator.py    # Main application (CLI + all generation logic)
supabase_client.py            # Optional Supabase integration
supabase_schema.sql           # Database schema for Supabase
requirements.txt              # Python dependencies
samples/                      # Example input files
tests/                        # Test suite (pytest)
api/                          # Vercel serverless functions
  generate.py                 # POST /api/generate - returns XLSX
  health.py                   # GET /api/health - health check
  requirements.txt            # Python deps for Vercel runtime
public/                       # Vercel static files
  index.html                  # Web UI for campaign generation
vercel.json                   # Vercel deployment config
.github/workflows/
  ci.yml                      # Tests on push/PR
  generate.yml                # Manual campaign generation workflow
```

## Key Commands

```bash
# Install dependencies
pip install -r requirements.txt

# Run the generator
python bulk_campaign_generator.py \
    --asins samples/sample_asins.txt \
    --listings-report samples/sample_active_listings.txt \
    --output output/campaigns.xlsx \
    --tier-file samples/sample_tier_assignment.json \
    --competitor-asins samples/sample_competitor_asins.txt \
    --no-supabase

# Run tests
python -m pytest tests/ -v
```

## Architecture

### Campaign Generation Flow

1. Parse Active Listings Report (tab-delimited or CSV) to build ASIN -> SKU mapping
2. Validate target ASINs against the mapping
3. Group SKUs by inventory tier (365+ days or 211-330 days)
4. For each tier, generate 3 campaigns:
   - **Manual Keywords** - Per-SKU ad groups with generic, size, and pack count keywords
   - **Auto** - Single ad group with all SKUs, auto-targeting groups
   - **Product Targeting** - Competitor ASIN targeting
5. Validate all rows against Amazon's expected format
6. Write to XLSX with sheet name "Sponsored Products Campaigns"

### Key Design Decisions

- **Negative IDs**: Amazon requires negative integers for new entity IDs in bulk uploads. Campaign IDs start at -1000, ad group IDs at -2000.
- **Blank cells**: Amazon rejects rows with `0` in fields that should be blank. The code explicitly writes `None` for empty cells.
- **Product Targeting Expression Type**: Must be set to `"auto"` for auto-targeting groups (close-match, loose-match, etc.) and `"manual"` for ASIN-based targeting.
- **Column names with spaces**: Amazon's column headers contain spaces (e.g., "Campaign Daily Budget"). The `_add_row` method uses `**kwargs` to handle these.

## Input File Formats

- **Active Listings Report**: Tab-delimited (.txt) or CSV. Must have ASIN and SKU columns.
- **Target ASINs**: File with one ASIN per line, or comma-separated string.
- **Competitor ASINs**: Same format as target ASINs (file or comma-separated).
- **Tier Assignment**: JSON mapping ASIN to tier string (e.g., `{"B08NV6CLGF": "365+"}`).
- **Config**: JSON with CampaignConfig fields (overrides all CLI budget/bid args when used).

## Common Issues & Solutions

### Amazon Upload Failures
- Always check that `Product Targeting Expression Type` is set (`auto` or `manual`)
- Never write `0` to fields that should be blank
- Campaign names must be unique
- SKUs must match exactly what's in Seller Central

### Config File vs CLI Args
- When `--config` is provided, it takes full precedence over CLI args like `--daily-budget-365`
- The tool now warns when both are specified

### Multiple SKUs per ASIN
- If the same ASIN maps to multiple SKUs in the listings report, the last one wins
- A warning is printed when this happens

## Testing

Tests are in `tests/test_bulk_campaign_generator.py` using pytest. Coverage includes:
- Column finding helper (`_find_column`)
- Active Listings Report parsing (tab-delimited and CSV)
- All validation functions (campaign, bidding adjustment, keyword, product targeting)
- Campaign generator (IDs, row creation, full structure generation)
- XLSX writer (blank cell handling, directory creation)
- End-to-end generation with sample data

## Supabase Integration

Optional. When `SUPABASE_URL` and `SUPABASE_KEY` env vars are set:
- Listings are persisted for cross-run ASIN lookups
- Each generation run is recorded with its config
- Upload results can be tracked
- Generated XLSX files are stored in a private bucket

Disable with `--no-supabase` even when credentials are configured.

## Deployment

### Vercel (Web UI + API)

The `api/` directory contains Python serverless functions. The `public/` directory has the web UI.

- `POST /api/generate` - Accepts JSON with ASINs, config, returns XLSX
- `GET /api/health` - Health check
- `GET /` - Web form UI

Deploy by connecting the GitHub repo to Vercel. Set `SUPABASE_URL` and `SUPABASE_KEY` as environment variables in Vercel dashboard for persistence.

### GitHub Actions

- `ci.yml` - Runs pytest on push/PR to main
- `generate.yml` - `workflow_dispatch` workflow: fill in ASINs/config in GitHub UI, download XLSX artifact

For Supabase in GitHub Actions, add `SUPABASE_URL` and `SUPABASE_KEY` as repository secrets.
