select  SM.id , sum(smu.paid) as Paid_sem, sm.seminar_type_id, sm.studio_id, sm.technolog_id, 
extract(day from sm.started_at) as Day, 
extract(month from sm.started_at)as Month, extract(year from sm.started_at) as Year ,to_char(sm.started_at,'dd.mm.YYYY') as FullDate ,


(case  when  extract(year from sm.closed_at) in ('2016')  then '1' else 0 end) as seminar_closed, sm.city_id, sm.name,
(case when sm.name like '%CRAFT%' or  sm.name like '%твор%'  or sm.name like '%МП%' then '1' else 0 end) as id_craft, -- нужно добавить твр* и *МП*
(case when sm.seminar_type_id in ( '20', '88', '21') then '1' else 0 end) as id_Day_MX, sm.city_name, sm.studio_name, sm.technolog_full_name, 
sm.users_count, sm.salons_count, sm.category, sm.region_id, '0'
--,(select distinct ct.region_name from  cities as CT where sm.region_id = ct.region_id ) as region_name
,sm.trip, sm.megaregion_id
,(select distinct smt.duration from  seminar_types as SMT where smt.id = sm.seminar_type_id) as duration


from seminars as SM
left join seminar_users as SMU ON sm.id = smu.seminar_id

where extract(year from started_at) in ('2016') --and sm.id in ('116973')


GROUP BY sm.id