
select usr.id, usr.full_name,  SLN.id as "ecad_salon_id",  concat(sln.name, '. ', SLN.address, '. ', sln.city_name) as "salon_name", sln.city_name, right(sln.com_mreg, char_length(sln.com_mreg)-3 ) as com_MREG,

(select Count( distinct usr.id ) 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id  and sm.closed_at is not Null
left join seminar_types as SMT on SM.seminar_type_id = SMT.id and  smt.id in ('74','78','77')

where sln.id = usr.salon_id and  smt.id in ('74','78','77') )as "profiber",

(select  to_char(Min( sm.closed_at ) , 'DD.MM.YYYY') 
from  seminar_users as SMU
left join users as usr on usr.id = SMU.user_id 
left join seminars as SM on SMU.seminar_id = SM.id and sm.closed_at is not Null
left join seminar_types as SMT on SM.seminar_type_id = SMT.id and  smt.id in ('74','78','77')

where sln.id = usr.salon_id and  smt.id in ('74','78','77')) as "data"

from salons as SLN

left join users as usr ON usr.salon_id = sln.id

GROUP BY sln.id, usr.id