-- Запрос выгружает ко-во уникальных мастеров и салонов за период --
-- в запросе стоит исключение на год и месяцы --
select 
usr.full_name, usr.id, 
usr.shtatnost, usr.chief, 
(select distinct usr2.chief from users as usr2 where usr.chief = usr2.full_name) as nPlus1,

count(
select smu.user_id
from seminars as smr
left join seminar_users as smu ON smr.id = smu.seminar_id
where smr.technolog_id = usr.id 
and to_char(started_at, 'YYYY') in ('2015') 
and to_char(started_at, 'MM') in ('01', '02', '03') ) as count_reg_USR, 

from users as usr

where (select smr.technolog_id from seminars as smr) = usr.id
--and usr3.id = '1924'

GROUP BY --usr.id, sln.id, smu.id, smr.id,
usr3.full_name, usr3.id
