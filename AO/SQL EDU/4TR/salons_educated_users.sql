

select   usr.id, usr.full_name, SLN.id,  concat(sln.name, '. ', SLN.address, '. ', sln.city_name),

(select Count( distinct seminar_id ) 
from  seminar_users as SMU
--left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id 
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where  sm.closed_at is not Null and usr.id = SMU.user_id) as "ALLTIME",
 
(select Count( distinct seminar_id ) 
from  seminar_users as SMU
--left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where  extract(year from sm.started_at) = '2014' and sm.closed_at is not Null and usr.id = SMU.user_id)  as "=2014",

(select Count( distinct seminar_id ) 
from  seminar_users as SMU
--left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where  extract(year from sm.started_at) = '2015' and sm.closed_at is not Null and usr.id = SMU.user_id)  as "=2015"



from users as usr

left join salons as sln on sln.id = usr.salon_id

GROUP BY usr.id, SLN.id 



