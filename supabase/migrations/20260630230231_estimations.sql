-- Estimador EBIM: tabla de cotizaciones guardadas (sustituye al storage local del artifact)
create table if not exists estimations (
  code text primary key,
  client text not null default '',
  project text not null default '',
  country text not null default '',
  data jsonb not null,
  created_by uuid references auth.users(id),
  saved_at timestamptz not null default now()
);

alter table estimations enable row level security;

-- Herramienta interna de EBIM: cualquier usuario autenticado (cuenta @ebim.pe / @grupoebim.com)
-- puede leer y escribir todas las cotizaciones (se comparten entre el equipo comercial).
create policy "authenticated users can read estimations"
  on estimations for select
  to authenticated
  using (true);

create policy "authenticated users can insert estimations"
  on estimations for insert
  to authenticated
  with check (true);

create policy "authenticated users can update estimations"
  on estimations for update
  to authenticated
  using (true);

create policy "authenticated users can delete estimations"
  on estimations for delete
  to authenticated
  using (true);

create index if not exists estimations_saved_at_idx on estimations (saved_at desc);
