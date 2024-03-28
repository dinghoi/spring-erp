
#��� ���� ���
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
AUTHOR : ����ȣ
DATE : 2021-10-06
DESC :
- ��븶���ϰ���� > ��� ���� ���
'
proc_body :
BEGIN	
	#ó�� ���� ��
	DECLARE state INT DEFAULT 0;

	#���� �ڵ鷯
	DECLARE EXIT HANDLER FOR SQLEXCEPTION
	BEGIN
		ROLLBACK;
		SET state = -1;
	END;	
	
	#Ʈ����� ����
	START TRANSACTION;	
		#��Ư�� ���� ���
		UPDATE overtime SET
			end_yn = 'N'
		WHERE SUBSTRING(work_date, 1, 7) >= p_from_date 
			AND SUBSTRING(work_date, 1, 7) <= p_to_date;		
		
		#�Ϲ� ��� ���� ���
		UPDATE general_cost SET
			end_yn = 'N'
		WHERE SUBSTRING(slip_date, 1, 7) >= p_from_date 
			AND SUBSTRING(slip_date, 1, 7) <= p_to_date;
		
		#����� ���� ���
		UPDATE transit_cost SET
			end_yn = 'N'
		WHERE SUBSTRING(run_date, 1, 7) >= p_from_date 
			AND SUBSTRING(run_date, 1, 7) <= p_to_date;
			
		#���� ������ ����
		DELETE FROM cost_end 
		WHERE end_month >= p_from_month AND end_month <= p_to_month;
		
		#����� ��� ���
		DELETE FROM company_as
		WHERE as_month >= p_from_month AND as_month <= p_to_month;
		
		DELETE FROM company_asunit 
		WHERE as_month >= p_from_month AND as_month <= p_to_month;
		
		DELETE FROM management_cost
		WHERE cost_month >= p_from_month AND cost_month <= p_to_month;	
	COMMIT;	
	
	SELECT state;
END;


