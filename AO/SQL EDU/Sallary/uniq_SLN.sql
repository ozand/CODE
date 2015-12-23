
select   DISTINCT sln.id, count(smu.user_id) , to_char(smr.started_at, 'MM') as nmMonth, to_char(started_at, 'DD') as nmDay,  to_char(started_at, 'DDMMYYYY') as nmData, smr.name, smr.studio_name

--(select count(user_id) from seminar_users where usr.id = smu.user_id and usr.salon_id = sln.id and smu.seminar_id = smr.id)



--count(distinct usr.id) as count_USR, 
--count(distinct sln.id) as count_SLN

from salons as sln
left join users as usr ON usr.salon_id = sln.id
left join seminar_users as SMU ON usr.id = smu.user_id
left join seminars as SMR ON smu.seminar_id = smr.id
--left join users as usr3  ON smr.technolog_id = usr3.id or smr.partimer_id = usr3.id

where to_char(started_at, 'YYYY') in ('2014') and smr.closed_at is not Null

GROUP BY sln.id, smr.started_at, smr.name, smr.studio_name


