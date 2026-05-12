-- FinFlow schema

create table if not exists debts (
  id          uuid primary key default gen_random_uuid(),
  name        text not null,
  type        text default 'Credit Card',
  balance     numeric(12,2) default 0,
  apr         numeric(6,3) default 0,
  min_payment numeric(12,2) default 0,
  due_day     integer check (due_day between 1 and 28),
  credit_limit numeric(12,2) default 0,
  notes       text default '',
  color       text default '#5c7cff',
  included    boolean default true,
  created_at  timestamptz default now()
);

create table if not exists income (
  id         uuid primary key default gen_random_uuid(),
  name       text not null,
  amount     numeric(12,2) default 0,
  frequency  text default 'monthly',
  created_at timestamptz default now()
);

create table if not exists expenses (
  id         uuid primary key default gen_random_uuid(),
  name       text not null,
  amount     numeric(12,2) default 0,
  category   text default 'other',
  created_at timestamptz default now()
);

create table if not exists goals (
  id                   uuid primary key default gen_random_uuid(),
  name                 text not null,
  target               numeric(12,2) default 0,
  current              numeric(12,2) default 0,
  monthly_contribution numeric(12,2) default 0,
  emoji                text default '🎯',
  notes                text default '',
  created_at           timestamptz default now()
);

create table if not exists settings (
  key   text primary key,
  value text
);

create table if not exists reminder_log (
  id         uuid primary key default gen_random_uuid(),
  created_at timestamptz default now(),
  type       text,
  recipient  text,
  subject    text,
  status     text,
  message    text
);

-- Default settings
insert into settings (key, value) values
  ('email',                    ''),
  ('phone_number',             ''),
  ('phone_carrier',            'att'),
  ('enable_email',             'true'),
  ('enable_sms',               'false'),
  ('reminder_days',            '7,3,1'),
  ('send_due_today',           'true'),
  ('overdue_alert_enabled',    'true'),
  ('overdue_grace_days',       '3'),
  ('weekly_digest_enabled',    'true'),
  ('monthly_report_enabled',   'true'),
  ('resend_api_key',           ''),
  ('payoff_strategy',          'avalanche'),
  ('extra_payment',            '0')
on conflict (key) do nothing;

-- Enable Row Level Security (open for now — lock down when you add auth)
alter table debts       enable row level security;
alter table income      enable row level security;
alter table expenses    enable row level security;
alter table goals       enable row level security;
alter table settings    enable row level security;
alter table reminder_log enable row level security;

-- Allow anon full access (replace with user-scoped policies when you add auth)
create policy "anon all" on debts        for all using (true) with check (true);
create policy "anon all" on income       for all using (true) with check (true);
create policy "anon all" on expenses     for all using (true) with check (true);
create policy "anon all" on goals        for all using (true) with check (true);
create policy "anon all" on settings     for all using (true) with check (true);
create policy "anon all" on reminder_log for all using (true) with check (true);
