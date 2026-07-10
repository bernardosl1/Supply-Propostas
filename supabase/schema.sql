-- Rode este arquivo no SQL Editor do Supabase.
-- Ele cria a tabela compartilhada de propostas e um bucket privado para DOCX/PDF.

create table if not exists public.propostas (
  id text primary key,
  numero_documento text not null default '',
  empresa_cliente text not null default '',
  data_documento text not null default '',
  preco_total_numero numeric not null default 0,
  data jsonb not null default '{}'::jsonb,
  docx_path text not null default '',
  pdf_path text not null default '',
  docx_storage_path text not null default '',
  pdf_storage_path text not null default '',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists propostas_updated_at_idx on public.propostas (updated_at desc);

alter table public.propostas enable row level security;

-- Para um primeiro uso interno sem tela de login, estas policies liberam acesso
-- para a anon key do app. Para produção com usuários, troque por policies
-- usando auth.role() = 'authenticated' ou regras por equipe.
drop policy if exists "Anon pode listar propostas" on public.propostas;
drop policy if exists "Anon pode inserir propostas" on public.propostas;
drop policy if exists "Anon pode atualizar propostas" on public.propostas;
drop policy if exists "Anon pode excluir propostas" on public.propostas;

create policy "Anon pode listar propostas"
  on public.propostas for select
  to anon
  using (true);

create policy "Anon pode inserir propostas"
  on public.propostas for insert
  to anon
  with check (true);

create policy "Anon pode atualizar propostas"
  on public.propostas for update
  to anon
  using (true)
  with check (true);

create policy "Anon pode excluir propostas"
  on public.propostas for delete
  to anon
  using (true);

insert into storage.buckets (id, name, public)
values ('propostas', 'propostas', false)
on conflict (id) do nothing;

drop policy if exists "Anon pode ler arquivos de propostas" on storage.objects;
drop policy if exists "Anon pode enviar arquivos de propostas" on storage.objects;
drop policy if exists "Anon pode atualizar arquivos de propostas" on storage.objects;
drop policy if exists "Anon pode excluir arquivos de propostas" on storage.objects;

create policy "Anon pode ler arquivos de propostas"
  on storage.objects for select
  to anon
  using (bucket_id = 'propostas');

create policy "Anon pode enviar arquivos de propostas"
  on storage.objects for insert
  to anon
  with check (bucket_id = 'propostas');

create policy "Anon pode atualizar arquivos de propostas"
  on storage.objects for update
  to anon
  using (bucket_id = 'propostas')
  with check (bucket_id = 'propostas');

create policy "Anon pode excluir arquivos de propostas"
  on storage.objects for delete
  to anon
  using (bucket_id = 'propostas');
