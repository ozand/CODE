select sln.id, sln.name, sln.address, sln.law_name, sln.city_name, sln.city_name_geographic, 

(select count(DISTINCT usr.id) from users as usr 
left join seminar_users as smu ON usr.id = smu.user_id
where sln.id = usr.salon_id 
and extract(year from smu.created_at) in ('2015')
) as countUSR, 
sln.hide

from salons as sln
ORDER by sln.id


--where extract(year from sln.created_at) in ('2015')