-- Run this once in Supabase -> SQL Editor.
-- Then create your admin user in Authentication -> Users.

create table if not exists public.portfolio_content (
  id integer primary key,
  html text not null default '',
  updated_at timestamptz not null default now()
);

alter table public.portfolio_content enable row level security;

drop policy if exists "Public can read portfolio" on public.portfolio_content;
create policy "Public can read portfolio"
on public.portfolio_content for select
to anon, authenticated
using (true);

drop policy if exists "Authenticated admin can insert" on public.portfolio_content;
create policy "Authenticated admin can insert"
on public.portfolio_content for insert
to authenticated
with check (true);

drop policy if exists "Authenticated admin can update" on public.portfolio_content;
create policy "Authenticated admin can update"
on public.portfolio_content for update
to authenticated
using (true)
with check (true);

insert into public.portfolio_content(id, html)
values (1, '')
on conflict (id) do nothing;

-- Storage bucket for profile/project images.
insert into storage.buckets (id, name, public)
values ('portfolio', 'portfolio', true)
on conflict (id) do update set public = true;

drop policy if exists "Public portfolio images" on storage.objects;
create policy "Public portfolio images"
on storage.objects for select
to public
using (bucket_id = 'portfolio');

drop policy if exists "Admin uploads portfolio images" on storage.objects;
create policy "Admin uploads portfolio images"
on storage.objects for insert
to authenticated
with check (bucket_id = 'portfolio');

drop policy if exists "Admin updates portfolio images" on storage.objects;
create policy "Admin updates portfolio images"
on storage.objects for update
to authenticated
using (bucket_id = 'portfolio')
with check (bucket_id = 'portfolio');

drop policy if exists "Admin deletes portfolio images" on storage.objects;
create policy "Admin deletes portfolio images"
on storage.objects for delete
to authenticated
using (bucket_id = 'portfolio');
