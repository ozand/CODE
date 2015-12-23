select   usr.id as "user_id", usr.full_name as "user_name", usr.email, usr.mobile_number, usr.salon_id as "ecad_salon_id", 
(select concat(sln.name, '. ', SLN.address, '. ', sln.city_name) from salons as SLN where usr.salon_id = sln.id) as "salon_name", 
(select smt.name from seminar_types as smt where smr.seminar_type_id = smt.id) as "Seminar_name",

--(select Count( sm.seminar_id  ) 
--from  seminar_users as SMU
--left join users as usr on usr.id = SMU.user_id 
--left join seminars as SM on SMU.seminar_id = SM.id 
--left join seminar_types as SMT on SM.seminar_type_id = SMT.id

case when  smr.closed_at is not Null and usr.id = SMU.user_id then '1'
else null
end as "ALLTIME",

case when extract(year from smr.started_at) = '2014' and smr.closed_at is not Null and usr.id = SMU.user_id then '1'
else null
end as "2014",

case when extract(year from smr.started_at) = '2015' and smr.closed_at is not Null and usr.id = SMU.user_id then '1'
else null
end as "2015"


from users as usr

left join seminar_users as SMU ON  usr.id = SMU.user_id
left join seminars as SMR on SMU.seminar_id = SMR.id  and SMR.closed_at is not Null

where smr.closed_at is not Null and smu.user_id = '2823'

--GROUP BY smu.seminar_id, usr.id,SMU.user_id, smr.started_at, smr.closed_at
order by usr.salon_id



