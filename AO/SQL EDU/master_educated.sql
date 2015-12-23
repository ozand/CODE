select  Distinct usr.id, usr.full_name, usr.role, SLN.id  as "ecad_salon_id",
concat(sln.name, '. ', SLN.address, '. ', sln.city_name)  as "salon_name", 
sln.is_closed, sln.hide, sln.com_mreg, sln.com_reg, sln.com_sect, sln.region_name as EDU_REG, sln.city_name_geographic, sln.city_name, sln.created_at, smr.city_name as Seminar_city, concat(smr.partimer_full_name, smr.technolog_full_name) as EDUCATER_NAME, to_char(smr.closed_at, 'DDMMYYYY') as SEMINAR_DATE,

(case when sum((case when extract(year from smr.closed_at) is not null then 1 else 0 end)) = 0 then Null else 1 end) as ALLTIME,
(case when sum((case when extract(year from smr.closed_at) in ('2014') then 1 else 0 end)) = 0 then Null else 1 end)  as EDU_PY,
(case when sum((case when extract(year from smr.closed_at) in ('2015') then 1 else 0 end)) = 0 then Null else 1 end)  as EDU_TY

from  users as usr
left join seminar_users as SMU ON usr.id = SMU.user_id
left join seminars as SMR ON SMU.seminar_id = SMR.id
left join salons as SLN ON usr.salon_id = sln.id 



--where 

GROUP BY usr.id, SLN.id, smr.city_name, smr.closed_at, EDUCATER_NAME
order by sln.id, usr.id

--limit 2000