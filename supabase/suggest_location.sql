-- Suggest Location per item (dari file Nota Penjualan harian)
-- Aturan: W*/OV* = Gudang (otomatis hijau, tidak wajib movement), C*/D* + angka = Area Selling (wajib),
-- kosong = "Tidak input Lokasi", lainnya = teks apa adanya. Lokasi hanya informasi + gudang dianggap auto.

create table if not exists public.mv_loc (
  id bigserial primary key,
  store_code text not null,
  np_no text not null,
  item_no text not null,
  sales_id text,
  loc text not null default '',
  qty numeric not null default 0,
  uploaded_by uuid default auth.uid(),
  uploaded_at timestamptz not null default now()
);
create index if not exists mv_loc_np on public.mv_loc(store_code, np_no, item_no);
alter table public.mv_loc enable row level security;
drop policy if exists mv_loc_sel on public.mv_loc; drop policy if exists mv_loc_ins on public.mv_loc;
drop policy if exists mv_loc_upd on public.mv_loc; drop policy if exists mv_loc_del on public.mv_loc;
create policy mv_loc_sel on public.mv_loc for select using (app_role() is not null);
create policy mv_loc_ins on public.mv_loc for insert with check (app_role() = any (array['admin','editor']));
create policy mv_loc_upd on public.mv_loc for update using (app_role() = any (array['admin','editor'])) with check (app_role() = any (array['admin','editor']));
create policy mv_loc_del on public.mv_loc for delete using (app_role() = any (array['admin','editor']));
drop trigger if exists zz_store_license on public.mv_loc;
create trigger zz_store_license before insert or update on public.mv_loc for each row execute function trg_store_license();
revoke all on public.mv_loc from anon;

create or replace function public._loc_kind(t text) returns text language sql immutable as $$
  select case when btrim(coalesce(t,'')) = '' then 'nil'
              when btrim(t) ~* '^(w[0-9]|ov)' then 'wh'
              when btrim(t) ~* '^[cd][0-9]' then 'sel'
              else 'oth' end $$;

drop function if exists public._mv_lines(text, date);
create function public._mv_lines(p_store text, p_date date)
returns table(sales_id text, item_no text, item_name text, np text, tm text, qty numeric, typ text, need boolean, known boolean, alloc numeric, locs jsonb, oq numeric)
language sql stable security definer set search_path to 'public' as $$
with src as (
  select coalesce(p->>'id','') sid, e
    from public.dashboard_days d, jsonb_array_elements(d.doc->'people') p, jsonb_array_elements(coalesce(p->'mv','[]'::jsonb)) e
   where d.store_code = p_store and d.sale_date = p_date
  union all
  select '' sid, e
    from public.dashboard_days d, jsonb_array_elements(coalesce(d.doc->'mvx','[]'::jsonb)) e
   where d.store_code = p_store and d.sale_date = p_date
), l as (
  select sid sales_id, lpad(btrim(e->>0),10,'0') item_no, coalesce(e->>1,'') item_name,
         nullif(btrim(coalesce(e->>2,'')),'') np, coalesce(e->>4,'') tm, (e->>3)::numeric qty
    from src where (e->>3)::numeric > 0
), t as (
  select l.*, (n.np_no is not null) nav_found, lower(btrim(coalesce(n.type,''))) nav_type
    from l left join public.mv_nav n on n.store_code = p_store and n.np_no = l.np
), lc as (
  select t.*,
    (select coalesce(jsonb_agg(jsonb_build_array(x.loc, x.q) order by x.loc), '[]'::jsonb)
       from (select btrim(m.loc) loc, sum(m.qty) q from public.mv_loc m
              where m.store_code = p_store and m.np_no = t.np and lpad(btrim(m.item_no),10,'0') = t.item_no
              group by btrim(m.loc)) x) locs,
    (select coalesce(sum(m.qty),0) from public.mv_loc m
       where m.store_code = p_store and m.np_no = t.np and lpad(btrim(m.item_no),10,'0') = t.item_no
         and public._loc_kind(m.loc) = 'wh') whq
    from t
), c as (
  select lc.sales_id, lc.item_no, lc.item_name, lc.np, lc.tm, lc.qty oq, lc.locs, lc.nav_found, lc.nav_type,
    case when np is null or not nav_found then 'Bawa Langsung'
         when nav_type = '' then 'Tanpa Type (NAV kosong)'
         else (select n.type from public.mv_nav n where n.store_code = p_store and n.np_no = lc.np) end typ,
    true known,
    greatest(0, lc.qty - least(lc.qty, lc.whq)) qty,
    (np is null or not nav_found or nav_type in ('cash&carry via pickup & delivery','cash&carry via gudang store')) need0
    from lc
), c2 as (
  select c.*, (c.need0 and c.qty > 0) need from c
), a as (
  select c2.*, coalesce(m.qty,0) moved,
    coalesce(sum(case when c2.need then c2.qty end) over (partition by c2.item_no order by c2.tm, coalesce(c2.np,''), c2.sales_id, c2.oq rows between unbounded preceding and 1 preceding),0) before_q
    from c2 left join public.mv_moves m on m.store_code = p_store and m.move_date = p_date and m.item_no = c2.item_no
)
select a.sales_id, a.item_no, a.item_name, a.np, a.tm,
       case when a.need then a.qty else a.oq end,
       a.typ, a.need, a.known,
       case when a.need then greatest(0, least(a.qty, a.moved - a.before_q)) end,
       a.locs, a.oq
  from a order by a.sales_id, a.tm, a.item_no;
$$;

create or replace function public.mv_report(p_store text, p_date date)
returns jsonb language plpgsql stable security definer set search_path to 'public' as $$
declare r jsonb; hm boolean; mn int; ma timestamptz; nn int; ln int;
begin
  if public.app_role() is null then return jsonb_build_object('ok', false, 'err', 'auth'); end if;
  select exists(select 1 from public.dashboard_days d where d.store_code = p_store and d.sale_date = p_date
           and (d.doc ? 'mvx' or exists(select 1 from jsonb_array_elements(d.doc->'people') p where p ? 'mv'))) into hm;
  select count(*), max(uploaded_at) into mn, ma from public.mv_moves where store_code = p_store and move_date = p_date;
  select count(*) into nn from public.mv_nav where store_code = p_store;
  select count(*) into ln from public.mv_loc where store_code = p_store;
  select coalesce(jsonb_agg(jsonb_build_array(sales_id,item_no,item_name,np,tm,qty,typ,need,known,alloc,locs,oq)), '[]'::jsonb) into r from public._mv_lines(p_store, p_date);
  return jsonb_build_object('ok', true, 'has_mv', hm, 'moves_n', mn, 'moves_at', ma, 'nav_n', nn, 'loc_n', ln, 'lines', r);
end $$;

create or replace function public.sp_mv_view(p_key text, p_code text)
returns jsonb language plpgsql stable security definer set search_path to 'public' as $$
declare v_store text; v_code text; v_out jsonb := '[]'::jsonb; d record; ls jsonb;
begin
  select store_code::text into v_store from public.share_links where token = p_key and not revoked and scope = 'sales';
  if v_store is null then return jsonb_build_object('ok', false, 'err', 'key'); end if;
  v_code := upper(btrim(coalesce(p_code, '')));
  if v_code ~ '^[0-9]{1,10}$' then v_code := 'S' || v_store || '-' || lpad(v_code, 10, '0'); end if;
  for d in select x.sale_date from (select dd.sale_date from public.dashboard_days dd
             where dd.store_code = v_store
               and exists(select 1 from jsonb_array_elements(dd.doc->'people') p where p->>'id' = v_code and p ? 'mv')
             order by dd.sale_date desc limit 7) x order by x.sale_date loop
    select coalesce(jsonb_agg(jsonb_build_array(item_no,item_name,np,tm,qty,typ,need,known,alloc,locs,oq)), '[]'::jsonb) into ls
      from public._mv_lines(v_store, d.sale_date) where sales_id = v_code;
    v_out := v_out || jsonb_build_array(jsonb_build_object('d', d.sale_date,
        'moves_n', (select count(*) from public.mv_moves where store_code = v_store and move_date = d.sale_date), 'l', ls));
  end loop;
  return jsonb_build_object('ok', true, 'days', v_out);
end $$;

-- _mv_lines hanya internal (dipanggil dari fungsi SECURITY DEFINER lain)
revoke all on function public._mv_lines(text,date) from public, anon, authenticated;
grant execute on function public._mv_lines(text,date) to service_role;
revoke all on function public._loc_kind(text) from public;
grant execute on function public._loc_kind(text) to anon, authenticated, service_role;
