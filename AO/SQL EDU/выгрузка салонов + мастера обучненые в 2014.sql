select SLN.id, 
SLN.name, 
SMU.user_id, 
SMU.user_full_name, 
SM.seminar_type_id, 
to_char(sm.started_at, 'YYYY') as year,
SM.started_at, 
SM.name, 
CONCAT(sm.technolog_full_name, sm.partimer_full_name ) as eduPpl,
SLN.id,
SLN.city_id, 	
SLN.name, 	
SLN.address, 	
SLN.phone, 	
SLN.manager, 	
SLN.law_name, 	
SLN.inn, 	
SLN.salon_code, 	
SLN.salon_type, 	
SLN.description, 	
SLN.created_at, 	
SLN.updated_at, 	
SLN.city_name, 	
SLN.free_certificate_count, 	
SLN.active_certificate_count, 	
SLN.technolog_id, 	
SLN.technolog_full_name, 	
SLN.region_id, 	SLN.region_name, 	
SLN.representative_id, 	
SLN.representative_full_name, 	
SLN.partner_code, 	
SLN.is_closed, 	SLN.megaregion_id, 	SLN.megaregion_name, 	SLN.grade, 	SLN.street, 	SLN.house, 	
SLN.ext_address, 	SLN.client_type, 	SLN.chain_name, 	SLN.open_month, 	SLN.open_year, 	SLN.discount, 	SLN.which_club, 	
SLN.barber_chairs, 	SLN.color_cost, 	SLN.com_mreg, 	SLN.com_reg, 	SLN.com_sect, 	SLN.brand




from salons as SLN
left join users as usr ON sln.id = usr.salon_id
left join seminar_users as SMU on usr.id = SMU.user_id
left join seminars as SM on SMU.seminar_id = SM.id
left join seminar_types as SMT on SM.seminar_type_id = SMT.id




