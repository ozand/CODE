select  usr.full_name, sln.id, sln.salon_code, sln.name, sln.street, sln.house, sln.city_name, sln.com_mreg, sln.com_reg, sln.com_sect, sln.open_year, sln.address

,
(select to_char(MAX(smr.started_at), 'YYYY')
from seminars AS smr 
left join seminar_users AS usm ON usr.id = usm.user_id
where usm.seminar_id = smr.id) as MX


from users AS usr
left join salons AS sln ON usr.salon_id = sln.id


order by usr.id
