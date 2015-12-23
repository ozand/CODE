select  usr.full_name, usr.id, usr.role,
usr.position, usr.city_name, usr.region_name, usr.shtatnost, usr.chief, 
(select distinct usr2.chief from users as USR2 where usr2.full_name=usr.chief ) as nPlus1, 
count( distinct smu.user_id) as uniq_USR, count( distinct smu.user_id) as uniq_USR, 
count( distinct sln.salon_id) as uniq_SLN 



from seminars AS SMR
left join users AS USR ON smr.technolog_id = usr.id or smr.partimer_id = usr.id
left join seminar_users AS SMU ON smr.id = smu.seminar_id
left join salon AS SMU ON s.id = smu.seminar_id



where to_char(started_at, 'YYYY') in ('2014') and usr.role not like 'master' and to_char(started_at, 'MM') in ('04', '05', '06') 

group by usr.full_name, usr.id
limit 100

