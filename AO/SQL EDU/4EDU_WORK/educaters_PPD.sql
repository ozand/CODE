select 
usr.id,
concat(usr.lname,' ' , usr.fname) as lname_fname ,
usr.role,
usr.position, 
usr.shtatnost, 
usr.email, 
(case  when usr.email  is Null then Null else count(usr.email) over (partition by usr.email) end ) as count_usr_email,
usr.mobile_number,  
(case  when usr.mobile_number  is Null then Null else count(usr.mobile_number) over (partition by usr.mobile_number) end ) as count_usr_phone,
usr.city_name, 
usr.commercial_megaregion, 
usr.region_name, 
usr.is_blocked, 
usr.chief, 
(count(smr_edu.id) ) as "smr_edu", 
SUM( smt.duration ) as w_day,
(select count(smu.user_id) from seminar_users as smu  where usr.id= smu.user_id) as "smr_mstr", 
(case  when  usr.login_count  is Null then 0 else 1 end) as "act_accnt",
(to_char(max(smr_edu.started_at),  'YYYY.MM.DD') ) as "last_smr_edu", 
to_char(usr.last_request_at, 'YYYY.MM.DD') as "last_access",
(Case when slnMng.id is not null then slnMng.id Else
(Case when sln.id is not null then sln.id else Null End ) End) as "salon_id",
(Case when slnMng.id is not null then concat(slnMng.name, '. ', slnMng.address, '. ', slnMng.city_name_geographic) Else
(Case when sln.id is not null then concat(sln.name, '. ', sln.address, '. ', sln.city_name_geographic) else '' End ) End) as "salon_name",
(Case when slnMng.id is not null then slnMng.com_mreg Else
(Case when sln.id is not null then sln.com_mreg else '' End ) End) as "com_mreg",
(Case when slnMng.id is not null then slnMng.com_reg Else
(Case when sln.id is not null then sln.com_reg else '' End ) End) as "com_reg",
(Case when slnMng.id is not null then slnMng.com_sect Else
(Case when sln.id is not null then sln.com_sect else '' End ) End) as "com_sec"


from users as usr
left join seminars as smr_edu  ON
	(case when smr_edu.partimer_id is not null then smr_edu.partimer_id else
	(case when smr_edu.technolog_id  is not null then smr_edu.technolog_id else
	(case when smr_edu.partner_id  is not null then smr_edu.partner_id end) end) end)  = usr.id
left join seminar_types as SMT ON smt.id = smr_edu.seminar_type_id
left join salons as sln ON usr.salon_id = sln.id
left join salons as slnMng ON usr.id = slnMng.salon_manager_id 

Where extract(year from smr_edu.started_at) in ('2016') and extract(month from smr_edu.started_at) = 8

GROUP BY  usr.id,  usr.full_name, slnmng.id, sln.id
Order by   usr.id 
