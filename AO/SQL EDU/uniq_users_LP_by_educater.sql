select  usr.id , usr.full_name, count( distinct smu.user_id)
from seminars as SMR
left join seminar_users as SMU on smr.id = smu.seminar_id
left join users AS USR ON smr.technolog_id = usr.id or smr.partimer_id = usr.id

where to_char(started_at, 'YYYY') in ('2014') and usr.role not like 'master' and to_char(started_at, 'MM') in ('04', '05', '06') 

group by usr.id, usr.full_name