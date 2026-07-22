create table if not exists public.solicitudes (
  id text primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.solicitudes enable row level security;
