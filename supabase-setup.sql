-- DocForge storage (run once in Supabase SQL Editor)
create table if not exists df_templates (id text primary key, doc jsonb not null);
create table if not exists df_history  (id text primary key, doc jsonb not null);
