

select distinct SLN.id, concat(sln.name, '. ', SLN.address, '. ', sln.city_name), 
sm.name,  sm.seminar_type_id, Count( distinct usr.id )


--case  extract(year from 
--(select sm.started_at 
--from seminars as sm 
--left join users as usr on sln.id = usr.salon_id
--left join seminar_users as SMU on usr.id = SMU.user_id
--where SMU.seminar_id = SM.id and usr.salon_id = sln.id ))
--when  '2014' then Count( distinct usr.id ) else null
--end as "2014",

--case  extract(year from 
--(select sm.started_at 
--from seminars as sm 
--left join users as usr on sln.id = usr.salon_id
--left join seminar_users as SMU on usr.id = SMU.user_id
--where SMU.seminar_id = SM.id and usr.salon_id = sln.id ))
--when  '2015' then Count( distinct usr.id ) else null
--end as "2015"



from  seminars as SM
left join seminar_users as smu ON sm.id = smu.seminar_id
left join users as usr on smu.user_id = usr.id
left join salons as sln ON usr.salon_id = sln.id  

where   
sm.closed_at is not null and seminar_type_id in ('2','3','10','21', '62', '6', '7', '9', '11', '19')

GROUP BY sm.name, sm.seminar_type_id, sln.id

