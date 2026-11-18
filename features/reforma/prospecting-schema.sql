create table public.bdflow_reforma_prospecting (
 id uuid primary key default gen_random_uuid(),
 user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
 report_date date not null,
 entries_text text not null check (length(entries_text) between 1 and 100000),
 source_label text not null default 'Registros de primera llamada' check (length(source_label) between 1 and 160),
 record_complete boolean not null default false,
 notes text not null default '' check (length(notes)<=2000),
 created_at timestamptz not null default now(),
 unique (user_id,report_date)
);
alter table public.bdflow_reforma_prospecting enable row level security;
revoke all on public.bdflow_reforma_prospecting from anon;
grant select,insert,update on public.bdflow_reforma_prospecting to authenticated;
create policy prospecting_select_own on public.bdflow_reforma_prospecting for select to authenticated using ((select auth.uid())=user_id);
create policy prospecting_insert_own on public.bdflow_reforma_prospecting for insert to authenticated with check ((select auth.uid())=user_id);
create policy prospecting_update_own on public.bdflow_reforma_prospecting for update to authenticated using ((select auth.uid())=user_id) with check ((select auth.uid())=user_id);
