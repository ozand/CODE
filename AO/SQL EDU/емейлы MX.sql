select usr.id, usr.full_name, usr.megaregion_name, usr.city_name, to_char(usr.created_at, 'YYYY') as add2Base , SM.studio_id, 
(select sm.technolog_id from seminar_users as SMU



sm.technolog_id, to_char(MAX(sm.started_at),'YYYY') as year_last_Seminar, sm.technolog_full_name, USR.email

from users as USR
left join seminar_users as SMU on USR.id = SMU.user_id
left join seminars as SM on SMU.seminar_id = SM.id


where usr.role like 'master'

GROUP BY usr.id,  SM.studio_id,sm.technolog_id, sm.technolog_full_name