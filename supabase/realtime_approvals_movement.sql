-- Realtime push untuk approval TL/MOD dan data Stock Movement / NAV (sudah diterapkan di proyek)
alter publication supabase_realtime add table public.sp_approvals;
alter publication supabase_realtime add table public.mv_moves;
alter publication supabase_realtime add table public.mv_nav;
