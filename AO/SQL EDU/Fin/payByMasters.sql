select smu.is_payed, smu.user_salon_name, smu.user_full_name, smu.user_salon_type, smu.paid, '0',--smu.coupon_code, 
smr.id, smr.seminar_type_id,
smr.name, smr.city_name, smr.studio_name, smr.technolog_full_name, '0' as nmData, extract( day from smr.started_at) as nmDay,
extract(month from smr.started_at) as nmMonth, smt.price_normal, smt.price_mbc

from seminar_users as smu
left join  seminars as smr ON smu.seminar_id = smr.id
left join seminar_types as smt ON smr.seminar_type_id = smt.id
where extract(year from smr.started_at) in ('2015')  and smr.closed_at is not Null

--GROUP BY  smr.started_at
--limit 100