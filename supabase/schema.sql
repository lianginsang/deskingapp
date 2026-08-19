-- Run this once in the Supabase dashboard: SQL Editor -> New query -> paste -> Run.
-- Creates a single-row table that holds the most recent uploaded spreadsheet,
-- so the "Toyota" button can replay it from any computer.

create table if not exists last_upload (
  id int primary key default 1,
  file_name text,
  raw_headers jsonb,
  raw_rows jsonb,
  field_map jsonb,
  updated_at timestamptz not null default now()
);

alter table last_upload enable row level security;

-- No login in front of this app, so reads/writes are open to anyone with the
-- publishable key (which itself ships in the app's public JS bundle).
create policy "Public read last_upload" on last_upload
  for select using (true);

create policy "Public write last_upload" on last_upload
  for insert with check (true);

create policy "Public update last_upload" on last_upload
  for update using (true) with check (true);
