
select  SLN.id as "ecad_salon_id",  trim (concat(sln.name, '. ', SLN.address, '. ', sln.city_name_geographic)) as "salon_name", sln.city_name, right(sln.com_mreg, char_length(sln.com_mreg)-3 ) as com_MREG,

(select Count( distinct usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id 
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and  sm.closed_at is not Null) as "U_ALLTIME",

 
(select Count( distinct usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and extract(year from sm.started_at) = '2015' and sm.closed_at is not Null) as "U_2015",

(select Count( distinct usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and extract(year from sm.started_at) = '2016' and sm.closed_at is not Null) as "U_2016",

(select Count(usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id 
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and  sm.closed_at is not Null) as "C_ALLTIME",

 
(select Count( usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and extract(year from sm.started_at) = '2015' and sm.closed_at is not Null) as "C_2015",

(select Count( usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id

where sln.id = usr.salon_id and extract(year from sm.started_at) = '2016' and sm.closed_at is not Null) as "C_2016",

(select Count( usr.id ) 
from users as usr

where sln.id = usr.salon_id ) as "Count_SLN_USRs",

(select Count( usr.id ) 
from users as usr
where sln.id = usr.salon_id and usr.last_request_at is not Null ) as "Count_actECAD_USRs",

(select Count(distinct usr.email ) 
from users as usr
where sln.id = usr.salon_id  ) as "Count_usr_email",

(select Count(distinct usr.mobile_number ) 
from users as usr
where sln.id = usr.salon_id  ) as "Count_usr_phone",


to_char(sln.created_at, 'DD.MM.YYYY')  as "add2ECAD",

(select usr.full_name from users as usr 
where sln.salon_manager_id = usr.id) as "Manager_SLN",

(select to_char(usr.last_request_at, 'DD.MM.YYYY')  from users as usr 
where sln.salon_manager_id = usr.id  and usr.last_request_at is not Null) as "Manager_SLN_last_accsess", 

(Select concat( smr.technolog_full_name, smr.partimer_full_name)
from seminar_users as smu
left join seminars as smr ON smr.id = smu.seminar_id
Left join users as usr ON smu.user_id = usr.id
left join salons as sln2 ON sln2.id = usr.salon_id
where sln.id = sln2.id
order by smr.started_at Desc
limit 1) as "last_educater_cont",

sln.show_on_locator, sln.com_sect, sln.com_reg, sln.city_name_geographic





from salons as SLN

GROUP BY sln.id 