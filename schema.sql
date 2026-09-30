-- schema.sql (V2 Supabase Cloud Database)
-- Run this in your Supabase SQL Editor:
-- https://supabase.com/dashboard/project/eqdccbnlimuvchkuhjpe/sql

-- 1. Create the orders table
create table if not exists public.orders (
    id uuid default gen_random_uuid() primary key,
    order_id text not null,
    sku text not null,
    title text,
    item_price numeric,
    cost_price numeric,
    shipping_cost numeric,
    profit_loss numeric,
    indicator text,
    store_name text not null default 'WISEROUTLET',
    purchase_date timestamp with time zone,
    created_at timestamp with time zone default timezone('utc'::text, now()) not null,
    grand_total numeric,
    amazon_tax numeric default 0,
    quantity integer default 1,
    total_cost numeric,
    order_status text,
    unique(order_id, sku, store_name)
);

-- 2. Add columns if table already exists (safe non-destructive migration)
alter table public.orders add column if not exists grand_total numeric;
alter table public.orders add column if not exists amazon_tax numeric default 0;
alter table public.orders add column if not exists quantity integer default 1;
alter table public.orders add column if not exists total_cost numeric;
alter table public.orders add column if not exists order_status text;

-- 3. Enable Row Level Security (RLS)
alter table public.orders enable row level security;

-- 4. Policies: Allow read and write for Dashboard & anonymous sync
drop policy if exists "Enable anon read access" on public.orders;
create policy "Enable anon read access" 
    on public.orders for select 
    to anon 
    using (true);

drop policy if exists "Enable anon all access" on public.orders;
create policy "Enable anon all access" 
    on public.orders for all 
    to anon 
    using (true)
    with check (true);

drop policy if exists "Enable authenticated all access" on public.orders;
create policy "Enable authenticated all access" 
    on public.orders for all 
    to authenticated 
    using (true)
    with check (true);
