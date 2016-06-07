-- Запрос выгружает ко-во уникальных мастеров и салонов за период --
-- в запросе стоит исключение на год и месяцы --
select 
usr3.full_name, usr3.id, 
usr3.shtatnost, usr3.chief, 
(select distinct usr2.chief from users as USR2 where usr2.full_name=usr3.chief) as nPlus1,

count(usr.id) as count_reg_USR, 
count(distinct usr.id) as count_unq_USR,

count(sln.id) as count_reg_SLN, 
count(distinct sln.id) as count_unq__SLN,
count(distinct usr.email) as count_email,
sum(smu.paid),
count(distinct to_char(started_at, 'DDMMYYYY') )as wDay

from users as usr
left join salons as sln ON usr.salon_id = sln.id
left join seminar_users as SMU ON usr.id = smu.user_id
left join seminars as SMR ON smu.seminar_id = smr.id
left join users as usr3  ON smr.technolog_id = usr3.id or smr.partimer_id = usr3.id

where to_char(started_at, 'YYYY') in ('2014') and to_char(started_at, 'MM') in ('04', '05', '06') 
--and usr3.id = '1924'

GROUP BY --usr.id, sln.id, smu.id, smr.id,
usr3.full_name, usr3.id
