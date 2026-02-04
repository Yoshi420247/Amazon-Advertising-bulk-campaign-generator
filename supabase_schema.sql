-- ============================================================================
-- Amazon Ads Bulk Campaign Generator - Supabase Schema
-- ============================================================================
-- Run this in your Supabase SQL Editor:
--   https://supabase.com/dashboard → your project → SQL Editor → New Query
-- ============================================================================

-- 1) Listings table: persistent ASIN→SKU mapping from Active Listings Reports
create table if not exists listings (
    id bigint generated always as identity primary key,
    asin text not null,
    seller_sku text not null,
    item_name text,
    fulfillment_channel text,
    price numeric(10, 2),
    quantity integer,
    status text default 'active',
    first_seen_at timestamptz default now(),
    last_seen_at timestamptz default now(),
    unique (asin, seller_sku)
);

create index if not exists idx_listings_asin on listings (asin);
create index if not exists idx_listings_sku on listings (seller_sku);

-- 2) Generation runs: history of every bulk file generated
create table if not exists generation_runs (
    id bigint generated always as identity primary key,
    created_at timestamptz default now(),
    target_asins jsonb not null,         -- ["B08NV6CLGF", ...]
    matched_skus jsonb not null,         -- {"B08NV6CLGF": "sku-123", ...}
    missing_asins jsonb default '[]',    -- ASINs not found in listings
    tier_assignment jsonb,               -- {"B08NV6CLGF": "365+", ...}
    competitor_asins jsonb,              -- ["B07ABC1234", ...]
    config jsonb not null,               -- full CampaignConfig as JSON
    row_count integer not null,
    campaign_count integer not null,
    file_name text,
    file_storage_path text,              -- path in Supabase Storage bucket
    notes text
);

-- 3) Upload results: track what happened when the file was uploaded to Amazon
create table if not exists upload_results (
    id bigint generated always as identity primary key,
    generation_run_id bigint references generation_runs(id) on delete set null,
    uploaded_at timestamptz default now(),
    status text not null default 'pending',  -- pending, success, partial, failed
    total_rows integer,
    processed_rows integer,
    failed_rows integer,
    error_summary text,                  -- free text summary of issues
    error_details jsonb,                 -- structured per-row errors if available
    notes text
);

create index if not exists idx_upload_results_run
    on upload_results (generation_run_id);

-- 4) Storage bucket for generated XLSX files
insert into storage.buckets (id, name, public)
values ('campaign-files', 'campaign-files', false)
on conflict (id) do nothing;

-- Storage policy: allow authenticated users to read/write their own files
create policy "Allow authenticated uploads"
    on storage.objects for insert
    to authenticated
    with check (bucket_id = 'campaign-files');

create policy "Allow authenticated reads"
    on storage.objects for select
    to authenticated
    using (bucket_id = 'campaign-files');

-- 5) Helpful views

-- View: latest listing state per ASIN (most recently seen SKU wins)
create or replace view listings_latest as
select distinct on (asin)
    id, asin, seller_sku, item_name, fulfillment_channel,
    price, quantity, status, first_seen_at, last_seen_at
from listings
order by asin, last_seen_at desc;

-- View: generation history with upload status
create or replace view generation_history as
select
    g.id as run_id,
    g.created_at,
    g.campaign_count,
    g.row_count,
    jsonb_array_length(g.target_asins) as asin_count,
    jsonb_array_length(g.missing_asins) as missing_count,
    g.file_name,
    u.status as upload_status,
    u.uploaded_at,
    u.processed_rows,
    u.failed_rows
from generation_runs g
left join lateral (
    select * from upload_results ur
    where ur.generation_run_id = g.id
    order by ur.uploaded_at desc
    limit 1
) u on true
order by g.created_at desc;
