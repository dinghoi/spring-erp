
#신규 AS 데이터 Insert 쿼리(매월 진행 처리)

#해당 월 확인 쿼리

SET @s_date = '2021-08-01';
SET @e_date=  '2021-08-31';

SELECT count(*)
FROM as_acpt
WHERE CAST(acpt_date as date) >= @s_date AND CAST(acpt_date as date) <= @e_date
;


SET @s_date = '2021-08-01';
SET @e_date=  '2021-08-31';

SELECT count(*)
FROM as_acpt_end
WHERE CAST(acpt_date as date) >= @s_date AND CAST(acpt_date as date) <= @e_date
;


/*
#누락데이터 발생했을 경우 아래 쿼리 실행

INSERT INTO as_acpt_end
SELECT *
FROM as_acpt
WHERE CAST(acpt_date as date) >= '2021-05-01' AND CAST(acpt_date as date) <= '2021-05-31'
	AND acpt_no NOT IN (
		SELECT acpt_no FROM as_acpt_end
		WHERE CAST(acpt_date as date) >= '2021-05-01' AND CAST(acpt_date as date) <= '2021-05-31'
	)
;
*/

SET @s_date = '2021-08-01';
SET @e_date=  '2021-08-31';

INSERT INTO as_acpt_end
SELECT *
FROM as_acpt
WHERE CAST(acpt_date as date) >= @s_date AND CAST(acpt_date as date) <= @e_date
;

