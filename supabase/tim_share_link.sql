-- Link tim sales (hanya baca, tanpa login): dashboard.html?t=TOKEN
-- Sudah dijalankan di Supabase. Disimpan di sini sebagai catatan.
create table if not exists public.share_links (token text primary key, store_code text not null, label text, revoked boolean not null default false, created_at timestamptz not null default now());
alter table public.share_links enable row level security;  -- tanpa policy: tidak bisa dibaca lewat API
create or replace function public.tim_view_data(p_token text) returns jsonb language plpgsql stable security definer set search_path = public as $$
declare v_store text; r jsonb;
begin
  select store_code into v_store from public.share_links where token = p_token and not revoked;
  if v_store is null then return null; end if;
  select jsonb_build_object('days', coalesce(jsonb_agg(jsonb_build_object('sale_date', sale_date, 'doc', doc - 'hours' - 'mb' - 'prot') order by sale_date), '[]'::jsonb),
    'sched', (select doc from public.dashboard_schedule where store_code::text = v_store limit 1)) into r
  from public.dashboard_days where store_code::text = v_store;
  return r;
end $$;
revoke all on function public.tim_view_data(text) from public;
grant execute on function public.tim_view_data(text) to anon, authenticated;
-- Buat link baru:  insert into public.share_links(token,store_code,label) values (replace(gen_random_uuid()::text,'-',''),'10005','Tim sales') returning token;
-- Matikan link:    update public.share_links set revoked = true where token = '...';

-- ===== Link salesperson (kode salesperson) =====
-- Migrasi 'sales_self_view' sudah diterapkan: kolom share_links.scope ('tim'|'sales'),
-- tim_view_data hanya menerima scope='tim', dan fungsi sp_view_data(p_key, p_code)
-- (SECURITY DEFINER, hanya mengembalikan data 1 salesperson untuk key scope='sales').
-- Buat/cabut key:
--   insert into share_links(token, store_code, label, scope) values ('<key>', '10005', 'Link salesperson', 'sales');
--   update share_links set revoked = true where token = '<key>';
-- Halaman: public/salesperson.html?k=<key>
-- Migrasi 'sp_view_items': sp_view_data kini juga mengembalikan 'it' (rincian item per hari:
-- [category, master brand, no item, deskripsi, qty, value]) di setiap elemen 'sales'.
-- Migrasi 'sp_view_stamps': sp_view_data juga mengembalikan sched_at dan data_at (waktu unggah terakhir).
-- Migrasi 'sp_approvals': tabel sp_approvals (approval TL dan MOD per salesperson per tanggal).
-- RLS: baca untuk semua peran login, tulis hanya editor/admin; trigger mengisi tl_by/mod_by dan waktu dari JWT.
-- sp_view_data juga mengembalikan 'r' (jumlah invoice per hari) dan 'appr' (status approval).
-- Migrasi 'sp_appr_roles': tabel sp_appr_managers (email, jabatan) = Store Manager / Asst Manager.
-- Approval MOD hanya bisa diubah akun di tabel ini; Approval TL hanya bisa diubah editor/admin yang BUKAN di tabel ini.
-- Tambah/hapus manager: insert into sp_appr_managers(email, jabatan) values ('x@mitra10.com','Asst Manager');

-- migration sp_view_pb_cats: sp_view_data per hari juga mengembalikan pb (penjualan Master Type Private Brand) dan cats (penjualan per category).

-- migration sp_view_zone_pbc: sp_view_data mengembalikan zcats (category di zonasi salesperson) dan pbc per hari (penjualan Private Brand per category).

-- migration sp_pin_approval: tabel sp_pins (hash PIN bcrypt, tanpa akses langsung), fungsi sp_pin_status, sp_pin_admin_list, sp_pin_admin_set (admin), sp_pin_change (ganti PIN sendiri), sp_approve (approval TL/MOD dengan PIN 6 digit; salah 5x = kunci 5 menit; MOD hanya setelah TL; batal TL ditolak bila MOD sudah approve).
-- Policy insert/update langsung pada sp_approvals dihapus: approval hanya lewat sp_approve.
