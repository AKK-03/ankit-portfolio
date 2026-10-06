-- Portfolio CMS setup for Supabase
-- Run once in Supabase -> SQL Editor -> New query -> Run

create table if not exists public.portfolio_content (
  id bigint primary key,
  html text not null default '',
  updated_at timestamptz not null default now()
);

alter table public.portfolio_content enable row level security;

drop policy if exists "Public can read portfolio" on public.portfolio_content;
create policy "Public can read portfolio"
on public.portfolio_content for select
to anon, authenticated
using (true);

drop policy if exists "Authenticated users can insert portfolio" on public.portfolio_content;
create policy "Authenticated users can insert portfolio"
on public.portfolio_content for insert
to authenticated
with check (true);

drop policy if exists "Authenticated users can update portfolio" on public.portfolio_content;
create policy "Authenticated users can update portfolio"
on public.portfolio_content for update
to authenticated
using (true)
with check (true);

insert into storage.buckets (id, name, public)
values ('portfolio', 'portfolio', true)
on conflict (id) do update set public = true;

drop policy if exists "Public can view portfolio images" on storage.objects;
create policy "Public can view portfolio images"
on storage.objects for select
to anon, authenticated
using (bucket_id = 'portfolio');

drop policy if exists "Authenticated users can upload portfolio images" on storage.objects;
create policy "Authenticated users can upload portfolio images"
on storage.objects for insert
to authenticated
with check (bucket_id = 'portfolio');

drop policy if exists "Authenticated users can update portfolio images" on storage.objects;
create policy "Authenticated users can update portfolio images"
on storage.objects for update
to authenticated
using (bucket_id = 'portfolio')
with check (bucket_id = 'portfolio');

drop policy if exists "Authenticated users can delete portfolio images" on storage.objects;
create policy "Authenticated users can delete portfolio images"
on storage.objects for delete
to authenticated
using (bucket_id = 'portfolio');
