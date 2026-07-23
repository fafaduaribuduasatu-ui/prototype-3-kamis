-- Jalankan ini di Supabase: buka project kamu > SQL Editor > New query > paste > Run

create table if not exists kv_store (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz default now()
);

-- Mengizinkan akses baca/tulis publik (dipakai bersama oleh 4 host lewat anon key).
-- Ini cukup untuk penggunaan internal tim kecil. Jangan taruh data sensitif selain
-- laporan GMV/gaji di tabel ini.
alter table kv_store enable row level security;

create policy "public read" on kv_store
  for select using (true);

create policy "public write" on kv_store
  for insert with check (true);

create policy "public update" on kv_store
  for update using (true) with check (true);
