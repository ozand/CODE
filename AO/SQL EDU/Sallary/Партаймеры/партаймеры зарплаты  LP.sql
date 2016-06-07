select concat(sm.technolog_id, sm.partimer_id,beauty_consultant_id) as idUser, sm.seminar_type_id, to_char(started_at,'DD') as Day, 
to_char(started_at,'MM')as Month, to_char(started_at,'YYYY') as Year ,to_char(started_at,'dd.mm.YYYY') as FullDate , smt.duration, sm.Name,sm.city_name,
 
(case  when sm.business_trip = 't' then '1' else 0 end) as trip,
(case  when to_char (sm.closed_at, 'YYYY') in ( '2016')  then '1' else 0 end) as seminar_closed,
concat( sm.technolog_full_name, sm.partimer_full_name, beauty_consultant_full_name) as Name,
(case  when sm.users_count = '0' then 0 else 1 end) as Users_Count



 
from seminars as sm
left join seminar_types as smt ON sm.seminar_type_id = smt.id

where to_char(started_at,'YYYY')in ('2016')