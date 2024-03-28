
#비용 마감 취소
DROP PROCEDURE IF EXISTS USP_COST_CANCEL_UPDATE;
CREATE PROCEDURE USP_COST_CANCEL_UPDATE(		
	IN p_from_date varchar(7),
	IN p_to_date varchar(7),
	IN p_from_month varchar(6),
	IN p_to_month varchar(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : 허정호
DATE : 2021-10-06
DESC :
- 비용마감일괄취소 > 비용 마감 취소
'
proc_body :
BEGIN	
	#처리 상태 값
	DECLARE state INT DEFAULT 0;

	#에러 핸들러
	DECLARE EXIT HANDLER FOR SQLEXCEPTION
	BEGIN
		ROLLBACK;
		SET state = -1;
	END;	
	
	#트랜잭션 설정
	START TRANSACTION;	
		#야특근 마감 취소
		UPDATE overtime SET
			end_yn = 'N'
		WHERE SUBSTRING(work_date, 1, 7) >= p_from_date 
			AND SUBSTRING(work_date, 1, 7) <= p_to_date;		
		
		#일반 경비 마감 취소
		UPDATE general_cost SET
			end_yn = 'N'
		WHERE SUBSTRING(slip_date, 1, 7) >= p_from_date 
			AND SUBSTRING(slip_date, 1, 7) <= p_to_date;
		
		#교통비 마감 취소
		UPDATE transit_cost SET
			end_yn = 'N'
		WHERE SUBSTRING(run_date, 1, 7) >= p_from_date 
			AND SUBSTRING(run_date, 1, 7) <= p_to_date;
			
		#마감 데이터 삭제
		DELETE FROM cost_end 
		WHERE end_month >= p_from_month AND end_month <= p_to_month;
		
		#공통비 배분 취소
		DELETE FROM company_as
		WHERE as_month >= p_from_month AND as_month <= p_to_month;
		
		DELETE FROM company_asunit 
		WHERE as_month >= p_from_month AND as_month <= p_to_month;
		
		DELETE FROM management_cost
		WHERE cost_month >= p_from_month AND cost_month <= p_to_month;	
	COMMIT;	
	
	SELECT state;
END;


