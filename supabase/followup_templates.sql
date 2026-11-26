-- Template pesan WhatsApp untuk blasting (halaman Follow Up Customer).
-- Jalankan di Supabase > SQL Editor. Aman dijalankan berulang kali.
create table if not exists public.fu_templates (
  id         bigint generated always as identity primary key,
  judul      text not null,
  isi        text not null,
  aktif      boolean not null default true,
  urut       int not null default 0,
  created_by uuid default auth.uid(),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.fu_templates enable row level security;
drop trigger if exists fu_templates_touch on public.fu_templates;
create trigger fu_templates_touch before update on public.fu_templates
  for each row execute function public.fu_touch();
drop policy if exists fu_t_select on public.fu_templates;
drop policy if exists fu_t_write  on public.fu_templates;
create policy fu_t_select on public.fu_templates for select to authenticated
  using ((select public.app_role()) is not null);
create policy fu_t_write on public.fu_templates for all to authenticated
  using ((select public.fu_is_manager())) with check ((select public.fu_is_manager()));

-- Gambar template (opsional): disimpan di Storage bucket publik "fu-templates"
alter table public.fu_templates add column if not exists gambar_url text;
alter table public.fu_templates add column if not exists gambar_path text;
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('fu-templates', 'fu-templates', true, 3145728, array['image/png','image/jpeg','image/webp'])
on conflict (id) do update set public = true, file_size_limit = excluded.file_size_limit, allowed_mime_types = excluded.allowed_mime_types;
drop policy if exists fu_tpl_img_insert on storage.objects;
drop policy if exists fu_tpl_img_update on storage.objects;
drop policy if exists fu_tpl_img_delete on storage.objects;
create policy fu_tpl_img_insert on storage.objects for insert to authenticated
  with check (bucket_id = 'fu-templates' and (select public.fu_is_manager()));
create policy fu_tpl_img_update on storage.objects for update to authenticated
  using (bucket_id = 'fu-templates' and (select public.fu_is_manager()))
  with check (bucket_id = 'fu-templates' and (select public.fu_is_manager()));
create policy fu_tpl_img_delete on storage.objects for delete to authenticated
  using (bucket_id = 'fu-templates' and (select public.fu_is_manager()));
