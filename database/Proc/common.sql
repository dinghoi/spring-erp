# SHOW PROCEDURE STATUS;

#================================================
# DB Error Log
DROP TABLE IF EXISTS nkp.error_log;
CREATE TABLE nkp.error_log(
	`error_log_id` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '���� �α� ID',
	`proc_name` VARCHAR(100) NOT NULL COMMENT '���ν��� �̸�',
	`proc_step` TINYINT(3) UNSIGNED NOT NULL COMMENT '���ν��� ������ ������ �߻��� ���� ��ȣ',
	`sql_state` VARCHAR(5) NOT NULL COMMENT 'SQLSTATE',
	`error_no` INT(11) NOT NULL COMMENT '���� ��ȣ',
	`error_msg` TEXT NOT NULL COMMENT '���� �޼���',
	#`call_stack` TEXT NULL COMMENT '���ν��� ȣ�� �Ķ����',
	`proc_call_date` DATETIME(0) NOT NULL COMMENT '���ν��� ȣ�� ����',
	`log_date` DATETIME(0) NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '�α� ���� ����',
PRIMARY KEY (`error_log_id`))
COMMENT = 'DB ��Ÿ�� ���� �α�';

#================================================
# DB Error Log ����
DROP PROCEDURE IF EXISTS nkp.USP_ERROR_LOG_INPUT;
CREATE PROCEDURE nkp.USP_ERROR_LOG_INPUT(
	IN p_proc_name VARCHAR(100),
	IN p_proc_step TINYINT(3),
	IN p_sql_state VARCHAR(5),
	IN p_error_no INT(11),
	IN p_error_msg TEXT,
	IN p_dt5_now DATETIME(0)
)
#NOT DETERMINISTIC
DETERMINISTIC
SQL SECURITY DEFINER
CONTAINS SQL
COMMENT '
AUTHOR : ����ȣ
DATE : 2021-08-25
DESC :
- DB ERROR LOG INSERT
RETURN VALUE :
'
proc_body :
BEGIN
	INSERT error_log(proc_name, proc_step, sql_state, error_no, error_msg, proc_call_date, log_date)
	VALUES(p_proc_name, p_proc_step, p_sql_state, p_error_no, p_error_msg, p_dt5_now, NOW(0));
END;

#================================================
# ���� ����(����)
DROP PROCEDURE USP_COMM_EMP_MASTER_NAME;
CREATE PROCEDURE USP_COMM_EMP_MASTER_NAME(
	IN p_emp_no VARCHAR(6)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : ����ȣ
DATE : 2021-08-20
DESC :
- ���� ����(����)'
proc_body :
BEGIN
	SELECT emp_name
	FROM emp_master
	WHERE emp_no = p_emp_no;
END;

#================================================
# �λ� �ڵ� ����
DROP PROCEDURE USP_COMM_ETC_CODE_INFO;
CREATE PROCEDURE USP_COMM_ETC_CODE_INFO(
	IN p_emp_etc_code VARCHAR(4)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : ����ȣ
DATE : 2021-08-20
DESC :
- �λ� �ڵ� ����'
proc_body :
BEGIN
	SELECT emp_etc_name
	FROM emp_etc_code
	WHERE emp_etc_type = p_emp_etc_code
	ORDER BY emp_etc_code ASC;
END;

#================================================
# ȸ�� �˻� ��ȸ
DROP PROCEDURE USP_COMM_ORG_LEVEL_INFO IF EXISTS USP_COMM_ORG_LEVEL_INFO;
CREATE PROCEDURE USP_COMM_ORG_LEVEL_INFO(
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : ����ȣ
DATE : 2021-08-05
DESC :
- ȸ�� �˻� Select'
proc_body :
BEGIN
	SELECT org_name
	FROM emp_org_mst
	WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00' OR org_end_date = '')
	AND org_level = 'ȸ��'
	ORDER BY org_company ASC;
END;

#================================================
# ���� SelectBox ��ȸ
DROP PROCEDURE IF EXISTS USP_COMM_ORG_SELECT_INFO;
CREATE PROCEDURE USP_COMM_ORG_SELECT_INFO(
	IN p_org_code VARCHAR(4)
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : ����ȣ
DATE : 2021-08-27
DESC :
- �λ� �ڵ� ����Ʈ �ڽ� ����
'
proc_body :
BEGIN
	SELECT org_company, org_bonbu, org_saupbu, org_team
	FROM emp_org_mst
	WHERE org_code = p_org_code;
END;

#================================================
# ������ �˻�
DROP PROCEDURE IF EXISTS USP_COMM_ORG_MST_INFO;
CREATE PROCEDURE USP_COMM_ORG_MST_INFO(
)
LANGUAGE SQL
#NOT DETERMINISTIC
DETERMINISTIC
CONTAINS SQL
SQL SECURITY DEFINER
COMMENT '
AUTHOR : ����ȣ
DATE : 2021-09-03
DESC :
- ������ �˻�'
proc_body :
BEGIN
	SELECT org_name
	FROM emp_org_mst
	WHERE (ISNULL(org_end_date)
			OR org_end_date = '0000-00-00'
			OR org_end_date = ''
			OR org_end_date = '1900-01-01')
		AND org_level = 'ȸ��'
	ORDER BY FIELD(org_company, '���̿�') DESC, org_code DESC;
END;
