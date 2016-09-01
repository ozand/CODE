Select
smr.id as smr_id,
concat(smr.id, '_', smt.name) as uniq_event,
smt.id as smr_type, 
smt.name, 
smt.kpis_type, 
smt.duration, 
(case when smt.is_free is true then 'free' else 'paid' end) as is_free,
smt.base_price,
extract(year from smr.started_at) as cd_year,
extract(month from smr.started_at)  as cd_month,
extract(day from smr.started_at)  as cd_day,
smr.started_at as cd_data_smr_start,
smr.closed_at as cd_data_smr_closed,
smr.created_at as cd_data_smr_creat,


(Case when smr.trip is true or false then
    (case when smr.trip = 't' then 1 else 0 end) 
else
    (case when smr.business_trip = 't' then '1' else 0 end)
end) as trip,

(case when smr.salon_id is not null or smr.salon_id not in ( '0' )then 'SALON' else
(case when smr.studio_id  is not null and (Select std.studio_type from studios as std where smr.studio_id = std.id) = 'classroom' then 'CLASSROOM' else 
(case when smr.studio_id  is not null and (Select std.studio_type from studios as std where smr.studio_id = std.id) = 'studio' then 'STUDIO' else  'NOT_TYPE' 
end) end) end) as type_place,

(case when smr.salon_id is not null or smr.salon_id not in ( '0' )then Concat('inSalon_', slnPlace.name,'. ',slnPlace.address) else
(case when smr.studio_id  is not null or smr.studio_id not in ( '0' ) then concat(std.name, '. ',std.address)  else '' end)
 end) as name_place,

(case when smr.closed_at is not null then 'CLOSED' else 'NOT_CLOSED' end) as seminar_closed,

(case when smr.partimer_id is not null then smr.partimer_id else
(case when smr.technolog_id  is not null then smr.technolog_id else
(case when smr.partner_id  is not null then smr.partner_id end) end) end) as educater_id,

(case when smr.partimer_id is not null then smr.partimer_full_name else
(case when smr.technolog_id  is not null then smr.technolog_full_name else 
(case when smr.partner_id  is not null then smr.partner_full_name else 'Not_found' end) end) end) as educater,

(case when smr.technolog_id is not null then (select usr_edu.role from users as usr_edu where smr.technolog_id = usr_edu.id) else
(case when smr.partimer_id is not null then (select usr_edu.role from users as usr_edu where smr.partimer_id = usr_edu.id) else
(case when smr.partner_id is not null then (select usr_edu.role from users as usr_edu where smr.partner_id = usr_edu.id) 
 end) end )end) as edu_role,

usr.id, usr.full_name, usr.role, 

(case when usr.email is not null then 1 else 0 end ) as status_email,
(case when usr.mobile_number is not null then 1 else 0 end ) as status_mobile,
(case when usr.last_request_at is not null then 1 else 0 end ) as status_ecad_active_user,

--usr.mobile_number, usr.email

(Case when usr.salon_id is not null then 'salon_master' else 
(Case when slnMNG.id is not null then 'salon_master' else 
(Case when usr.id is not null then 'hairdresser' else 'not_reg_user' end) end) end) as type_master,

(Case when usr.salon_id is not null then usr.salon_id else slnMNG.id end) as salon_id,

(Case when usr.salon_id is not null then concat(sln.id, '_', sln.name, '. ',sln.address) else 
(Case when slnMNG.id is not null then concat(slnMNG.id, '_', slnMNG.name, '. ',slnMNG.address) else '' end) end) as salon,

(Case when usr.salon_id is not null then sln.com_mreg else slnMNG.com_mreg end) as com_mreg,
(Case when usr.salon_id is not null then sln.com_mreg else slnMNG.com_reg end) as com_reg,
(Case when usr.salon_id is not null then sln.com_sect else slnMNG.com_sect end) as com_sect,
(Case when usr.salon_id is not null then sln.client_type else slnMNG.client_type end) as client_type,


(Case when sln.id is not null then
	(Case when sln.is_closed = 't' then 'in_ptnc_base' else 'in_act_base' end) else
	(Case when slnMNG.id is not null then
(Case when slnMNG.is_closed = 't' then 'in_ptnc_base' else 'in_act_base' end) end) end) as active_clnt

--(select distinct usr2.chief from users as USR2 where usr2.full_name=usr.chief limit 1 ) as nPlus1

from seminars as SMR
left join seminar_types as smt On smr.seminar_type_id = smt.id 
left join seminar_users as smu ON smu.seminar_id = smr.id
left join users as usr ON smu.user_id = usr.id
left join salons as sln ON usr.salon_id is not null and usr.salon_id = sln.id
left join salons as slnPlace ON smr.salon_id is not null and smr.salon_id = slnPlace.id
left join salons as slnMNG ON usr.salon_id is null and usr.id = slnMNG.salon_manager_id

--left join salons as slnManag ON usr.id = sln.salon_manager_id 

left join studios as std ON smr.studio_id is not null and smr.studio_id = std.id

Where   extract(year from smr.started_at) in ('2015', '2016') and extract(month from smr.started_at) <= 7
--group by smr.started_at, smu.user_Id