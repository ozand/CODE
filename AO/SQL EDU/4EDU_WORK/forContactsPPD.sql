select id, role ,lname, fname, email, mobile_number, city_name, commercial_megaregion, sector_name, 
position, shtatnost, is_blocked, created_at , chief, wstatus
from users
where role in ('admin', 'studio_administrator', 'technolog', 'regional_technolog', 'super_technolog', 'partner', 'partimer', 'parttimer')