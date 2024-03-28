

SELECT *
FROM emp_org_mst
WHERE org_level = '회사'
;



SELECT org_owner_org
FROM emp_org_mst
WHERE org_code = '6514'
;

SELECT *
FROM emp_master
WHERE emp_org_code = '6717'
;

SELECT *
FROM pay_month_give
WHERE pmg_yymm = '202106'
AND pmg_emp_name = '김승일'
;

SELECT *
FROM emp_master_month
WHERE emp_month = '202106'
AND emp_name = '김승일'
;


SELECT IF(r.org_level = '회사', r.org_level, '') AS 'level1',
	IF(r.org_level = '본부', r.org_level, '') AS 'level2',
	IF(r.org_level = '사업부', r.org_level, '') AS 'level3',
	IF(r.org_level = '팀', r.org_level, '') AS 'level4',
	IF(r.org_level = '상주처', r.org_level, '') AS 'level5',
	IF(r.org_level = '지사', r.org_level, '') AS 'level6',
	IF(r.org_level = '파트', r.org_level, '') AS 'level7',
	r.org_code, r.org_name, 
	IF((SELECT count(*) FROM emp_org_mst WHERE org_owner_org = r.org_code) > 0, 'Y', 'N') AS 'ownerYn',
	r.org_owner_org,
	(SELECT org_name FROM emp_org_mst WHERE org_code = r.org_owner_org) AS 'ownerName',
	r.org_company, r.org_bonbu, r.org_saupbu, r.org_team,	
	r.org_reside_company, r.org_reside_place,
	r.org_cost_center, r.org_cost_group,		
	IF(r.org_end_date = '' OR r.org_end_date IS NULL OR r.org_end_date = '1900-01-01',
		'Y', 'N') AS 'orgYn',
	IF((SELECT COUNT(*) FROM emp_master WHERE emp_org_code = r.org_code AND emp_pay_id <> '2') > 0,
		'Y', 'N') 'empYn'
FROM (
	SELECT *
	FROM (
		SELECT * FROM emp_org_mst
		ORDER BY org_owner_org, org_code
		) AS org_sorted,
		(SELECT @pv := '1001') AS initialisation
	WHERE find_in_set(org_owner_org, @pv)
		AND LENGTH(@pv := concat(@pv, ',', org_code))
) r		
ORDER BY r.org_owner_org ASC,
	FIELD(r.org_level, '회사', '본부', '사업부', '팀', '상주처', '지사', '파트') ASC	
;


/*조직도 검색*/

SET @pid = '3001';

SELECT /*rr.p_id,*/
	if(FIND_IN_SET(rr.p_id, rr.tNode) = 1, rr.p_id, '') AS node1,
	if(FIND_IN_SET(p_id, tNode) = 2, rr.p_id, '') AS node2,
	if(FIND_IN_SET(p_id, tNode) = 3, rr.p_id, '') AS node3,
	if(FIND_IN_SET(p_id, tNode) = 4, rr.p_id, '') AS node4,
	if(FIND_IN_SET(p_id, tNode) = 5, rr.p_id, '') AS node5,
	if(FIND_IN_SET(p_id, tNode) = 6, rr.p_id, '') AS node6,
	if(FIND_IN_SET(p_id, tNode) = 7, rr.p_id, '') AS node7,
	
	eomt.org_owner_org, 	
	(SELECT org_name FROM emp_org_mst WHERE org_code = eomt.org_owner_org) AS 'parentOrg',
	
	eomt.org_level, eomt.org_name, 
	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team,
	eomt.org_reside_company, eomt.org_reside_place,
	eomt.org_cost_center, eomt.org_cost_group,	
	
	eomt.org_end_date,
	if((SELECT count(*) FROM emp_org_mst 
	WHERE (org_end_date = '' OR org_end_date IS NULL OR org_end_date = '1900-01-01')
		AND org_code = rr.p_id) > 0, 'Y', 'N') AS 'orgUseYn',
	if((SELECT count(*) FROM emp_master 
	WHERE emp_pay_id <> '2' AND emp_org_code = rr.p_id) > 0, 'Y', 'N') AS 'empUseYn'
	
FROM (
	SELECT r.p_id, 
		IF(r.p_id = @pid, r.p_id, 
			if(r.p1_id = @pid, concat(r.p1_id, ',', r.p_id), 
				if(r.p2_id = @pid, concat(r.p2_id, ',', r.p1_id, ',', r.p_id), 
					if(r.p3_id = @pid, concat(r.p3_id, ',', r.p2_id, ',', r.p1_id, ',', r.p_id), 
						if(r.p4_id = @pid, concat(r.p4_id, ',', r.p3_id, ',', r.p2_id, ',', r.p1_id, ',', r.p_id), 
							if(r.p5_id = @pid, concat(r.p5_id, ',', r.p4_id, ',', r.p3_id, ',', r.p2_id, ',', r.p1_id, ',', r.p_id), 
								if(r.p6_id = @pid, concat(r.p6_id, ',', r.p5_id, ',', r.p4_id, ',', r.p3_id, ',', r.p2_id, ',', r.p1_id, ',', r.p_id), '')
							)
						)
					)
				)
			)
		) AS tNode	
	FROM (
		SELECT 
			
			/*p8.org_owner_org AS p8_id,
			p7.org_owner_org AS p7_id,*/
			p6.org_owner_org AS p6_id,
			p5.org_owner_org AS p5_id,
			p4.org_owner_org AS p4_id,
			p3.org_owner_org AS p3_id,
			p2.org_owner_org AS p2_id,
			p1.org_owner_org AS p1_id,
			p1.org_code AS p_id
		FROM emp_org_mst AS p1
		LEFT JOIN emp_org_mst AS p2 ON p2.org_code = p1.org_owner_org
		LEFT JOIN emp_org_mst AS p3 ON p3.org_code = p2.org_owner_org
		LEFT JOIN emp_org_mst AS p4 ON p4.org_code = p3.org_owner_org
		LEFT JOIN emp_org_mst AS p5 ON p5.org_code = p4.org_owner_org
		LEFT JOIN emp_org_mst AS p6 ON p6.org_code = p5.org_owner_org
		/*LEFT JOIN emp_org_mst AS p7 ON p7.org_code = p6.org_owner_org
		LEFT JOIN emp_org_mst AS p8 ON p8.org_code = p7.org_owner_org*/
		WHERE @pid IN (
			p1.org_owner_org,
			p2.org_owner_org,
			p3.org_owner_org,
			p4.org_owner_org,
			p5.org_owner_org,
			p6.org_owner_org/*,
			p7.org_owner_org,
			p8.org_owner_org*/
		) 
	) r		
) rr
INNER JOIN emp_org_mst AS eomt ON rr.p_id = eomt.org_code
ORDER BY rr.tNode ASC 
;


/*인사_회사 */

SELECT eomt.org_company
FROM emp_master AS emtt
INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
WHERE emtt.emp_pay_id <> '2'
GROUP BY eomt.org_company
LIMIT 100
;

/*
에스유에이치
케이네트웍스
케이더봄
케이시스템
케이원
케이원정보통신
*/

/*인사_본부 */

SELECT eomt.org_bonbu
FROM emp_master AS emtt
INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
WHERE emtt.emp_pay_id <> '2'
AND eomt.org_company = '에스유에이치'
GROUP BY eomt.org_bonbu
;

/*인사 정보*/

SET @mm = '202106';
SET @company = '케이원';
SET @bonbu = '공공SI본부';

SELECT eomt.org_code, eomt.org_name, eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team,
	eomt.org_reside_company, eomt.org_reside_place, eomt.org_cost_center, eomt.org_cost_group,
	
	emtt.emp_no, emtt.emp_name,
	emtt.emp_org_name, emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team,
	emtt.emp_reside_company, emtt.emp_reside_place, emtt.emp_stay_name,
	emtt.cost_center, emtt.cost_group,
	
	emmt.emp_org_code, emmt.emp_org_name, emmt.emp_company, emmt.emp_bonbu, emmt.emp_saupbu, emmt.emp_team,
	emmt.emp_reside_company, emmt.emp_reside_place, emmt.emp_stay_name,
	emmt.cost_center, emmt.cost_group, emmt.mg_saupbu,
	
	IF(pmgt.pmg_emp_no IS NULL OR pmgt.pmg_emp_no = '', 'N', IF(pmgt.pmg_id = '1', 'Y', 'N')) AS '지급여부'
	
FROM emp_master AS emtt
INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code
LEFT OUTER JOIN emp_master_month AS emmt ON emtt.emp_no = emmt.emp_no
	AND emmt.emp_month = @mm
LEFT OUTER JOIN pay_month_give AS pmgt ON emmt.emp_no = pmgt.pmg_emp_no
	AND pmgt.pmg_yymm = @mm
WHERE emtt.emp_pay_id <> '2'
	AND eomt.org_company = @company	
	AND (eomt.org_bonbu = @bonbu
		/*OR eomt.org_bonbu IS NULL*/
	)
ORDER BY eomt.org_code ASC, emtt.emp_no ASC
;

