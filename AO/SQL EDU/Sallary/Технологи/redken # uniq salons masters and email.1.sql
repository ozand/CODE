-- Запрос выгружает ко-во уникальных мастеров и салонов за период --
-- в запросе стоит исключение на год и месяцы --
select 
usr_edu.full_name, 
usr_edu.id,
usr_edu.shtatnost, 
usr_edu.chief,
(select distinct usr_n1.chief from users as usr_n1 where usr_edu.chief = usr_n1.full_name limit 1) as nPlus1,

count(smu.user_id) as count_reg_USR ,
count(distinct smu.user_id) as count_unq_USR,
count(usr.salon_id) as count_reg_SLN,
count(distinct usr.salon_id) as count_unq__SLN,
count(distinct usr.email) as count_email,
sum(smu.paid),
count(distinct to_char(smr.started_at, 'DDMMYYYY')) as wDay

from seminars as SMR 

left join seminar_users as SMU ON smr.id = smu.seminar_id
left join users as usr ON smu.user_id = usr.id
left join users as usr_edu ON  smr.technolog_id = usr_edu.id or  smr.partimer_id = usr_edu.id

where to_char(smr.started_at, 'YYYY') in ('2015') and to_char(started_at, 'MM') in ('10', '11', '12') 

GROUP BY usr_edu.full_name, usr_edu.id, smr.partimer_id, smr.partimer_full_name, usr_edu.shtatnost, usr_edu.chief
