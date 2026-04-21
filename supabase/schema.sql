-- Run this in Supabase SQL Editor.

create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  email text,
  display_name text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, email)
  values (new.id, new.email)
  on conflict (id) do update set email = excluded.email, updated_at = now();
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute procedure public.handle_new_user();

create table if not exists public.tender_comments (
  id uuid primary key default gen_random_uuid(),
  tender_key text not null,
  user_id uuid not null references auth.users(id) on delete cascade,
  user_email text,
  comment_text text not null check (length(trim(comment_text)) > 0),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists tender_comments_tender_key_idx on public.tender_comments (tender_key);
create index if not exists tender_comments_created_at_idx on public.tender_comments (created_at desc);

create or replace function public.set_comment_updated_at()
returns trigger
language plpgsql
as $$
begin
  new.updated_at = now();
  return new;
end;
$$;

drop trigger if exists set_comment_updated_at on public.tender_comments;
create trigger set_comment_updated_at
before update on public.tender_comments
for each row execute procedure public.set_comment_updated_at();

create table if not exists public.tender_score_overrides (
  id uuid primary key default gen_random_uuid(),
  tender_key text not null,
  user_id uuid not null references auth.users(id) on delete cascade,
  user_email text,
  score_value integer not null check (score_value between 1 and 10),
  reason_text text,
  created_at timestamptz not null default now()
);

create index if not exists tender_score_overrides_tender_key_idx on public.tender_score_overrides (tender_key);
create index if not exists tender_score_overrides_created_at_idx on public.tender_score_overrides (created_at desc);

alter table public.profiles enable row level security;
alter table public.tender_comments enable row level security;
alter table public.tender_score_overrides enable row level security;

drop policy if exists "profiles_select_authenticated" on public.profiles;
create policy "profiles_select_authenticated"
on public.profiles
for select
to authenticated
using (true);

drop policy if exists "profiles_update_own" on public.profiles;
create policy "profiles_update_own"
on public.profiles
for update
to authenticated
using (auth.uid() = id)
with check (auth.uid() = id);

drop policy if exists "comments_select_authenticated" on public.tender_comments;
create policy "comments_select_authenticated"
on public.tender_comments
for select
to authenticated
using (true);

drop policy if exists "comments_insert_authenticated" on public.tender_comments;
create policy "comments_insert_authenticated"
on public.tender_comments
for insert
to authenticated
with check (auth.uid() = user_id);

drop policy if exists "comments_update_own" on public.tender_comments;
create policy "comments_update_own"
on public.tender_comments
for update
to authenticated
using (auth.uid() = user_id)
with check (auth.uid() = user_id);

drop policy if exists "comments_delete_own" on public.tender_comments;
create policy "comments_delete_own"
on public.tender_comments
for delete
to authenticated
using (auth.uid() = user_id);

drop policy if exists "overrides_select_authenticated" on public.tender_score_overrides;
create policy "overrides_select_authenticated"
on public.tender_score_overrides
for select
to authenticated
using (true);

drop policy if exists "overrides_insert_authenticated" on public.tender_score_overrides;
create policy "overrides_insert_authenticated"
on public.tender_score_overrides
for insert
to authenticated
with check (auth.uid() = user_id);

-- Append-only history for overrides: no update/delete policy on purpose.
