create table if not exists public.modelos_propostas (
  id text primary key,
  nome text not null,
  empresa text not null,
  estrutura jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.modelos_propostas enable row level security;

drop policy if exists "modelos_propostas_select" on public.modelos_propostas;
create policy "modelos_propostas_select"
on public.modelos_propostas for select
to anon, authenticated
using (true);

drop policy if exists "modelos_propostas_insert" on public.modelos_propostas;
create policy "modelos_propostas_insert"
on public.modelos_propostas for insert
to anon, authenticated
with check (true);

drop policy if exists "modelos_propostas_update" on public.modelos_propostas;
create policy "modelos_propostas_update"
on public.modelos_propostas for update
to anon, authenticated
using (true)
with check (true);

drop policy if exists "modelos_propostas_delete" on public.modelos_propostas;
create policy "modelos_propostas_delete"
on public.modelos_propostas for delete
to anon, authenticated
using (true);

grant select, insert, update, delete on public.modelos_propostas to anon, authenticated;
