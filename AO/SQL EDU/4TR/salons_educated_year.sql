

select  SLN.id, concat(sln.name, '. ', SLN.address, '. ', sln.city_name), extract(year from sm.started_at),

(select Count( distinct usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id 
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and  sm.closed_at is not Null) as "ALLTIME",
 
(select Count( distinct usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and extract(year from sm.started_at) = '2014' and sm.closed_at is not Null) as "=2014",

(select Count( distinct usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and extract(year from sm.started_at) = '2015' and sm.closed_at is not Null) as "=2015"



from salons as SLN
left join users as usr on sln.id = usr.salon_id 
left join seminar_users as SMU on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id


GROUP BY sln.id , sm.started_at



