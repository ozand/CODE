

select  SLN.id, concat(sln.name, '. ', SLN.address, '. ', sln.city_name),  sm.name,  sm.id, sm.seminar_type_id,
 

case when extract(year from sm.started_at) is not Null then Count( distinct usr.id ) else null 
end as "ALLTIME",

case  extract(year from sm.started_at)
when  '2014' then Count( distinct usr.id ) else null
end as "2014",

case  extract(year from sm.started_at)
when  '2015' then Count( distinct usr.id ) else null
end as "2015"

from salons as SLN
left join users as usr on sln.id = usr.salon_id
left join seminar_users as SMU on usr.id = SMU.user_id
left join seminars as SM on SMU.seminar_id = SM.id


where sm.closed_at is not null --and sln.id = '8977'

GROUP BY sln.id, sm.id, sm.name,  sm.started_at


