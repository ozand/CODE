select smt.name, 
case when (select count(usr.id) 
from seminars as smr 
left join seminar_users as smu ON smr.id = smu.seminar_id
left join users as usr on smu.user_id = usr.id

where smr.seminar_type_id = smt.id and usr.role = 'master' and smt.name not like '%я.%'
 ) >100 then 
(select count(usr.id) 
from seminars as smr 
left join seminar_users as smu ON smr.id = smu.seminar_id
left join users as usr on smu.user_id = usr.id

where smr.seminar_type_id = smt.id and usr.role = 'master' and smt.name not like '%я.%'
 )
 else '0'
 end  as "count"

from seminar_types as smt


order by  "count" desc