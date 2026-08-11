create table if not exists public.portfolio_state (
  portfolio_id text not null default 'default',
  key text not null,
  value jsonb not null,
  updated_at timestamptz not null default now(),
  primary key (portfolio_id, key)
);

alter table public.portfolio_state enable row level security;

drop policy if exists "portfolio_state_public_read" on public.portfolio_state;
create policy "portfolio_state_public_read"
on public.portfolio_state
for select
to anon
using (true);

drop policy if exists "portfolio_state_public_write" on public.portfolio_state;
create policy "portfolio_state_public_write"
on public.portfolio_state
for insert
to anon
with check (true);

drop policy if exists "portfolio_state_public_update" on public.portfolio_state;
create policy "portfolio_state_public_update"
on public.portfolio_state
for update
to anon
using (true)
with check (true);

insert into storage.buckets (id, name, public)
values ('portfolio-files', 'portfolio-files', true)
on conflict (id) do update set public = true;

drop policy if exists "portfolio_files_public_read" on storage.objects;
create policy "portfolio_files_public_read"
on storage.objects
for select
to anon
using (bucket_id = 'portfolio-files');

drop policy if exists "portfolio_files_public_upload" on storage.objects;
create policy "portfolio_files_public_upload"
on storage.objects
for insert
to anon
with check (bucket_id = 'portfolio-files');

drop policy if exists "portfolio_files_public_update" on storage.objects;
create policy "portfolio_files_public_update"
on storage.objects
for update
to anon
using (bucket_id = 'portfolio-files')
with check (bucket_id = 'portfolio-files');
