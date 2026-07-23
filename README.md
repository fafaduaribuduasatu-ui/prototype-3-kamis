# Perolehan Live Streaming Faradisaaaa

Versi mandiri dari tracker GMV & gaji host live streaming, siap di-deploy ke domain sendiri.

## Langkah 1 — Buat database di Supabase (gratis)

1. Buka https://supabase.com → Sign up / Login → **New project**
2. Kasih nama project bebas, tunggu sampai selesai dibuat (±2 menit)
3. Di sidebar kiri, klik **SQL Editor** → **New query**
4. Copy seluruh isi file `supabase-schema.sql` dari folder ini → paste → klik **Run**
5. Di sidebar klik **Project Settings** (ikon gear) → **API**
6. Catat 2 hal ini, nanti dipakai di Langkah 3:
   - **Project URL**
   - **anon public key**

## Langkah 2 — Push kode ke GitHub

1. Buat akun GitHub kalau belum punya: https://github.com
2. Buat repository baru (bisa private), lalu upload semua isi folder ini ke repo tersebut
   (paling gampang: pakai GitHub Desktop, atau upload manual lewat web GitHub)

## Langkah 3 — Deploy ke Vercel (gratis)

1. Buka https://vercel.com → Sign up pakai akun GitHub kamu
2. Klik **Add New → Project** → pilih repo yang tadi kamu upload
3. Sebelum klik Deploy, buka bagian **Environment Variables**, tambahkan:
   - `VITE_SUPABASE_URL` → isi dengan Project URL dari Langkah 1
   - `VITE_SUPABASE_ANON_KEY` → isi dengan anon public key dari Langkah 1
4. Klik **Deploy**, tunggu 1-2 menit
5. Kamu akan dapat link sementara seperti `nama-project.vercel.app` — coba buka, pastikan aplikasi jalan normal

## Langkah 4 — Hubungkan domain kamu sendiri

1. Di dashboard Vercel, buka project kamu → tab **Settings → Domains**
2. Ketik domain kamu (misal `laporanlive.com` atau `live.tokoku.com`) → **Add**
3. Vercel akan kasih 1-2 baris DNS record (biasanya tipe `A` atau `CNAME`) yang harus ditambahkan
4. Login ke tempat kamu beli domain (Niagahoster, Domainesia, Rumahweb, GoDaddy, dll) → cari menu **DNS Management / Kelola DNS**
5. Tambahkan record persis seperti yang diminta Vercel (biasanya butuh 10 menit - 24 jam sampai aktif, biasanya lebih cepat)
6. Setelah aktif, domain kamu otomatis mengarah ke aplikasi ini, lengkap dengan HTTPS otomatis dari Vercel

## Catatan keamanan

- Data (laporan GMV, nama host, PIN admin) tersimpan di tabel `kv_store` yang bisa dibaca/ditulis siapa saja yang tahu link-nya (perlu, karena 4 host input dari HP masing-masing tanpa login).
- Kalau ke depan mau lebih aman (tiap host punya login sendiri), Supabase juga punya fitur **Authentication** yang bisa ditambahkan — kabari saja kalau mau dibantu ke arah situ.
- Jangan bagikan Supabase **service_role key** ke siapa pun (yang dipakai di sini hanya **anon key**, aman untuk dipakai di frontend).

## Menjalankan di komputer sendiri (opsional, untuk coba-coba dulu)

```bash
npm install
cp .env.example .env
# isi .env dengan URL & anon key dari Supabase
npm run dev
```
