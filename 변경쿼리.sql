-- 2019.03.14 정원재 요구로 상주처 이름 변경 (33명)
-- 강명수,강영훈,김강열,김민석,김소라,김종엽,김휘민,박근재,박의종,박태원,서민혁,신정수,엄태수,오영민,오원강,오유미,우성주,유원상,윤우영,이경종,이금호,이도희,이진수,이현수,전영주,정상운,정우열,지미환,진윤재,최규성,최신형,한희정,황인석

SELECT user_name  , user_grade , USEr_id, org_name, bonbu, saupbu, team
 FROM memb a
 WHERE user_id in (
'101953','101318','101831','101057','101723','101536','101020','101839','101180','102029','100853','100742','101320','101321','101801','101794','100303','101319','101615','100874','101210','101952','101834','100563','101316','101074','101698','101685','101982','101907','101900','101928','101277' 
)
 ORDER by user_name;
 
SELECT emp_no, emp_name, emp_org_name
  FROM emp_master
 WHERE emp_no  in (
   '101953','101318','101831','101057','101723','101536','101020','101839','101180','102029','100853','100742','101320','101321','101801','101794','100303','101319','101615','100874','101210','101952','101834','100563','101316','101074','101698','101685','101982','101907','101900','101928','101277'
 )
 ORDER by emp_name;


-- 2019.02.22 박정신 요구 'N/W 1사업부','N/W 2사업부',"SI3사업부","솔루션사업부"	는 나오지않도록 조건으로 처리..
-- 비용관리 시스템 / 비용현황관리 / 비용마감관리 cost_end_mg.asp

 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='1088';   -- "N/W 1사업부"
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='4088';   -- "N/W 1사업부"
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='6088';   -- "N/W 1사업부"
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='6319';   -- "N/W 2사업부"
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='6332';   -- "N/W 2사업부"
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='6343';   -- "N/W 2사업부"
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='1052';   -- "SI3사업부"	
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='3052';   -- "SI3사업부"	
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='4052';   -- "SI3사업부"	
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='6052';   -- "SI3사업부"	 
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='1214';   -- "솔루션사업부"	
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='3214';   -- "솔루션사업부"	
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='4214';   -- "솔루션사업부"	
 UPDATE emp_org_mst SET org_end_date='2018-12-31' WHERE  org_code='6214';   -- "솔루션사업부"	


-- 2019.02.22 박정신 요청 '사업부별 손익총괄'에서 해당년도에 사업부를 셋팅하면됨


SELECT * FROM sales_org WHERE sales_year = 2019 -- 해당년도에 사업부를 셋팅하면됨 (saupbu_profit_loss_total.asp)


-- 2019.02.22 윤성희(박아영) 요청 자격증 등록

insa_individual_qual_add.asp

select * from emp_etc_code where emp_etc_type = '30' order by emp_etc_name asc

INSERT INTO `emp_etc_code` VALUES ('3124', '30', '자격증종류', '운전면허 1종 보통', '', '', '1', 'Y', NULL, '', NULL, 'N');
INSERT INTO `emp_etc_code` VALUES ('3125', '30', '자격증종류', '한식조리기능사', '', '', '1', 'Y', NULL, '', NULL, 'N');
INSERT INTO `emp_etc_code` VALUES ('3126', '30', '자격증종류', '조주기능사', '', '', '1', 'Y', NULL, '', NULL, 'N');
INSERT INTO `emp_etc_code` VALUES ('3127', '30', '자격증종류', '영양사', '', '', '1', 'Y', NULL, '', NULL, 'N');
 


-- 2019.02.20 지현주 요청  6452 조직변경..

SELECT * from emp_org_mst where org_code = 6452;

SET `org_saupbu`='KDC사업부'
, `org_team`='스마트워크'
, `org_name`='행안부'
, `org_reside_place`='행안부'
, `org_reside_company`='' ;

select emp_no /*,emp_name*/ from emp_master
where emp_org_code  = 6452;

'101986'	'주범수'
'101987'	'이명희'
'101988'	'오재훈'
'101989'	'신민호'
'101990'	'김은선'
'101991'	'김예슬'
'101992'	'문혜리'
'101993'	'박진아'
'101994'	'신동미'
'101995'	'이지연'
'101996'	'조혜림'
'101997'	'박소영'
'101998'	'임서영'
'101999'	'노승철'
'102000'	'조찬욱'
'102001'	'신두삼'
'102002'	'박아영'
'102003'	'박성은'
'102004'	'이은혜'
'102005'	'이연경'
'102006'	'황지나'
'102007'	'강연화'
'102008'	'이희진'
'102009'	'박나은'
'102010'	'홍은정'
'102011'	'김태민'
'102012'	'김소연'
'102013'	'김단비'
'102014'	'박성화'
'102015'	'강민정'
'102016'	'최미희'
'102017'	'박유진'
'102018'	'김나현'
'102019'	'김수현'
'102020'	'임선희'
'102021'	'박지숙'
'102022'	'손정애'
'102025'	'김병훈'


emp_master
  `emp_saupbu`='기술연구소'
, `emp_team`='플렉시블'
, `emp_org_name`='스마트워크'
, `emp_reside_place`='스마트워크'
, `emp_reside_company`='행안부' 
where emp_org_code  = 6452
;


UPDATE emp_master
SET `emp_saupbu`='KDC사업부'
, `emp_team`='스마트워크'
, `emp_org_name`='행안부'
, `emp_reside_place`='행안부'
, `emp_reside_company`='' 
where emp_org_code  = 6452
;

select * from memb
where user_id 
in(
'101986'
,'101987'
,'101988'
,'101989'
,'101990'
,'101991'
,'101992'
,'101993'
,'101994'
,'101995'
,'101996'
,'101997'
,'101998'
,'101999'
,'102000'
,'102001'
,'102002'
,'102003'
,'102004'
,'102005'
,'102006'
,'102007'
,'102008'
,'102009'
,'102010'
,'102011'
,'102012'
,'102013'
,'102014'
,'102015'
,'102016'
,'102017'
,'102018'
,'102019'
,'102020'
,'102021'
,'102022'
,'102025'

)
;

update memb
SET `saupbu`='KDC사업부'
, `team`='스마트워크'
, `org_name`='행안부'
, `reside_place`='행안부'
, `reside_company`='' 
where user_id 
in(
'101986'
,'101987'
,'101988'
,'101989'
,'101990'
,'101991'
,'101992'
,'101993'
,'101994'
,'101995'
,'101996'
,'101997'
,'101998'
,'101999'
,'102000'
,'102001'
,'102002'
,'102003'
,'102004'
,'102005'
,'102006'
,'102007'
,'102008'
,'102009'
,'102010'
,'102011'
,'102012'
,'102013'
,'102014'
,'102015'
,'102016'
,'102017'
,'102018'
,'102019'
,'102020'
,'102021'
,'102022'
,'102025'

)
;

