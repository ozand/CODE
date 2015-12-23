select usr.id, usr.full_name, usr.role, usr.salon_id, usr.salon_name, 
-- (select to_char(smr.started_at,'YYYY') from seminars as smr where smr.id = smu.seminar_id) as SMR_Date, 
(select to_char(max(smr.started_at),'YYYY') from seminars as smr where smr.id = smu.seminar_id) as last_SMR_Date

from  users as usr
left join seminar_users as smu on usr.id = smu.user_id
 


--where user_salon_id = '4329'
--limit 100
