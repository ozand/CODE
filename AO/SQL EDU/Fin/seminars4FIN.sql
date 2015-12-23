select * , extract(month from sm.created_at) as month, extract(day from sm.created_at) as day 
from seminars as sm
where extract(year from sm.created_at) = '2015' and extract(year from sm.closed_at) is not Null