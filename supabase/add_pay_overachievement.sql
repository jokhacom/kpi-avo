-- Режим премии «Учитывать перевыполнение выше 100%» (переключает только Super Admin).
-- По умолчанию выключен: каждый показатель для премии не выше 100%.
alter table kpi_settings add column if not exists pay_overachievement boolean default false;
