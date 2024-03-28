

SELECT *
FROM emp_master
WHERE emp_name = '최상돈'
;

SELECT *
FROM emp_master_month
WHERE emp_month = '202012'
	AND emp_no = '102434'
;


SELECT *
FROM memb
WHERE user_id = '100084'
;


/*



한춘기 차장(100084)
-> [케이네트웍스 - SI2본부 - 수도권사업부 - 강원지사](한화생명 원주)

조용남 과장(102434)
-> [케이원 - OA수행본부 - 수도권사업부 - 강원지사]

안성민 대리(102440)
-> [케이원 - OA수행본부 - 수도권사업부 - 강원지사]

최상돈 차장(101717)
-> [케이원 - SI2본부 - 수도권사업부 - 강원지사](한화생명 강릉)

박선형 대리(101828)
-> [케이네트웍스 - SI2본부 - 수도권사업부 - 강원지사](한화생명 춘천)

황동훈 차장(100151)
-> [케이원 - SI1본부 - 수도권사업부 - 강원지사](한진 강릉)

장호열 대리(100858)
-> [케이네트웍스 - SI2본부 - 수도권사업부 - 강원지사](한화손보 원주)

박홍근 대리(101174)
-> [케이원 - SI1본부 -  - 고객지원 1팀](한국도로교통공단)

장준호 사원(102499)
-> [케이원 - SI1본부 -  - 고객지원 1팀](한국도로교통공단)

고지원 사원(102719)
-> [케이원 - SI1본부 -  - 고객지원 1팀](한국도로교통공단)

최재현 사원(102502)
-> [케이네트웍스 - NI본부 - 수도권사업부 - 강원지사](KT원주)

최종호 차장(101268)
-> [케이원 - SI2본부 - 수도권사업부 - 강원지사](한국가스공사 강원)

임대성 과장(101267)
-> [케이원 - SI2본부 - 수도권사업부 - 강원지사](한국가스공사삼척기지)

홍성민 대리(102147)
-> [케이원 - SI1본부 - 수도권사업부 - 강원지사](KB손보 강릉)

백종호 사원(102267)
-> [케이원 - SI1본부 - 수도권사업부 - 강원지사](KB손보 강릉)


*/

SELECT emp_org_code, emp_no, emp_name, emp_type, emp_grade, emp_first_date,
	emp_company, emp_bonbu, emp_saupbu, emp_team, emp_org_name,
	emp_reside_company, emp_reside_place, emp_stay_name, cost_center
FROM emp_master
WHERE emp_no = '102267'
;	


/*
select *
from transit_cost 
inner join (SELECT user_id, user_name FROM memb) memb
on transit_cost.mg_ce_id = memb.user_id 
inner join emp_masterON emp_master.emp_no = memb.user_id
where (run_date >= '2021-04-01' and run_date <= '2021-04-30')
and transit_cost.saupbu = emp_master.emp_saupbu
and (transit_cost.car_owner = '개인' or transit_cost.car_owner = '회사') 
ORDER BY memb.user_name asc, run_date desc, run_seq desc
limit 0,10 
;
*/

 
select *
from transit_cost 
inner join (SELECT user_id, user_name FROM memb) memb
	on transit_cost.mg_ce_id = memb.user_id 
inner join emp_master 
	ON emp_master.emp_no = memb.user_id
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	and transit_cost.mg_ce_id = '102267'
	and (transit_cost.car_owner = '개인' or transit_cost.car_owner = '회사') 
ORDER BY memb.user_name asc, run_date desc, run_seq desc
;


SELECT *
FROM transit_cost
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '102267'
	and (car_owner = '개인' OR car_owner = '회사') 
;
	
/*
한춘기 차장(100084)
-> [케이네트웍스 - SI2본부 - 수도권사업부 - 강원지사](한화생명 원주)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = '한화생명 원주'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '100084'
	and (car_owner = '개인' OR car_owner = '회사') 
;



조용남 과장(102434)
-> [케이원 - OA수행본부 - 수도권사업부 - 강원지사]

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = '강원지사'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '102434'
	and (car_owner = '개인' OR car_owner = '회사') 
;

안성민 대리(102440)
-> [케이원 - OA수행본부 - 수도권사업부 - 강원지사]

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = '강원지사'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '102440'
	and (car_owner = '개인' OR car_owner = '회사') 
;

최상돈 차장(101717)
-> [케이원 - SI2본부 - 수도권사업부 - 강원지사](한화생명 강릉)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = '한화생명 강릉'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '101717'
	and (car_owner = '개인' OR car_owner = '회사') 
;

박선형 대리(101828)
-> [케이네트웍스 - SI2본부 - 수도권사업부 - 강원지사](한화생명 춘천)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = '한화생명 춘천'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '101828'
	and (car_owner = '개인' OR car_owner = '회사') 
;

황동훈 차장(100151)
-> [케이원 - SI1본부 - 수도권사업부 - 강원지사](한진 강릉)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = '한진 강릉',
	reside_place = '한진 강릉'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '100151'
	and (car_owner = '개인' OR car_owner = '회사') 
;

장호열 대리(100858)
-> [케이네트웍스 - SI2본부 - 수도권사업부 - 강원지사](한화손보 원주)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = '한화손보 원주',
	reside_place = '한화손보 원주'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '100858'
	and (car_owner = '개인' OR car_owner = '회사') 
;

최재현 사원(102502)
-> [케이네트웍스 - NI본부 - 수도권사업부 - 강원지사](KT원주)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사',
	org_name = 'KT원주'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '102502'
	and (car_owner = '개인' OR car_owner = '회사') 
;


홍성민 대리(102147)
-> [케이원 - SI1본부 - 수도권사업부 - 강원지사](KB손보 강릉)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '102147'
	and (car_owner = '개인' OR car_owner = '회사') 
;

백종호 사원(102267)
-> [케이원 - SI1본부 - 수도권사업부 - 강원지사](KB손보 강릉)

UPDATE transit_cost SET
	saupbu = '수도권사업부',
	team = '강원지사'
where (run_date >= '2021-04-01' and run_date <= '2021-05-31')
	AND mg_ce_id = '102267'
	and (car_owner = '개인' OR car_owner = '회사') 
;

*/

