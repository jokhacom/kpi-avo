-- Ставка налога с премии (в процентах), настраивается в «Настройках KPI».
-- Выполнить один раз: Supabase → SQL Editor → вставить → Run.
alter table kpi_settings add column if not exists tax_rate numeric default 7.5;
