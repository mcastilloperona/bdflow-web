create table public.bdflow_reforma_signatures (
 id uuid primary key default gen_random_uuid(),
 operation_id uuid not null unique references public.bdflow_reforma_operations(id),
 user_id uuid not null references auth.users(id),
 status text not null check(status in ('pending','signed')),
 signed_date date,
 next_followup date,
 notes text not null default '' check(length(notes)<=10000),
 constraint signature_date_consistent check ((status='signed' and signed_date is not null) or (status='pending' and signed_date is null))
);
alter table public.bdflow_reforma_signatures enable row level security;
revoke all on public.bdflow_reforma_signatures from public, anon, authenticated;
grant select,insert,update on public.bdflow_reforma_signatures to authenticated;
create policy signatures_select on public.bdflow_reforma_signatures for select to authenticated using (
 user_id=(select auth.uid()) and exists(select 1 from public.bdflow_access a where a.user_id=(select auth.uid())));
create policy signatures_insert on public.bdflow_reforma_signatures for insert to authenticated with check (
 user_id=(select auth.uid()) and exists(select 1 from public.bdflow_access a where a.user_id=(select auth.uid())) and exists(select 1 from public.bdflow_reforma_operations o where o.id=operation_id and o.user_id=(select auth.uid())));
create policy signatures_update on public.bdflow_reforma_signatures for update to authenticated using (
 user_id=(select auth.uid()) and exists(select 1 from public.bdflow_access a where a.user_id=(select auth.uid()))) with check (
 user_id=(select auth.uid()) and exists(select 1 from public.bdflow_access a where a.user_id=(select auth.uid())) and exists(select 1 from public.bdflow_reforma_operations o where o.id=operation_id and o.user_id=(select auth.uid())));
create index bdflow_reforma_signatures_user on public.bdflow_reforma_signatures(user_id);

