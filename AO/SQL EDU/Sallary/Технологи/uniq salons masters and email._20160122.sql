select 
usr.full_name, usr.id, 
usr.shtatnost, usr.chief, 
(select distinct usr2.chief from users as USR2 where usr2.full_name=usr.chief) as nPlus1,

count(smu.user_id) as count_reg_USR, 
count(distinct smu.user_id) as count_unq_USR,

count(select usr2.salon_id from users as usr2 where smu.user_id = usr2.id) as count_reg_SLN,
count(distinct(select usr2.salon_id from users as usr2 where smu.user_id = usr2.id)) as count_unq__SLN,
count(distinct usr.email) as count_email,
sum(smu.paid),
count(distinct to_char(started_at, 'DDMMYYYY') )as wDay

from seminars as smr
left join users as usr ON usr.id = smr.technolog_id or usr.id = smr.partimer_id
left join seminar_users as smu ON smr.id = smu.seminar_id

where to_char(started_at, 'YYYY') in ('2015') and to_char(started_at, 'MM') in ('10', '11', '12') 
--and usr3.id = '1924'

GROUP BY usr.id, usr.full_name, smr.started_at
