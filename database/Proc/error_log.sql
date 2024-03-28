



/*
DROP TABLE IF EXISTS nkp.TranTest;
CREATE TABLE `TranTest` (
  `num` int(11) NOT NULL auto_increment,
  `col01` varchar(32) default NULL,
  PRIMARY KEY  (`num`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

DROP TABLE IF EXISTS nkp.TranTest2;
CREATE TABLE `TranTest2` (
  `num` int(11) NOT NULL auto_increment,
  `col01` varchar(32) default NULL,
  PRIMARY KEY  (`num`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

# DB Error Log
DROP TABLE IF EXISTS nkp.error_log;
CREATE TABLE nkp.error_log(
	`error_log_id` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '에러 로그 ID',
	`proc_name` VARCHAR(100) NOT NULL COMMENT '프로시저 이름',
	`proc_step` TINYINT(3) UNSIGNED NOT NULL COMMENT '프로시저 내에서 에러가 발생한 스텝 번호',
	`sql_state` VARCHAR(5) NOT NULL COMMENT 'SQLSTATE',
	`error_no` INT(11) NOT NULL COMMENT '에러 번호',
	`error_msg` TEXT NOT NULL COMMENT '에러 메세지',
	#`call_stack` TEXT NULL COMMENT '프로시저 호출 파라미터',
	`proc_call_date` DATETIME(0) NOT NULL COMMENT '프로시저 호출 일자',
	`log_date` DATETIME(0) NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '로그 적재 일자',
PRIMARY KEY (`error_log_id`))
COMMENT = 'DB 런타임 에러 로그';

DROP TABLE IF EXISTS nkp.web_error_log;
CREATE TABLE nkp.web_error_log(
	`error_log_id` SMALLINT(5) UNSIGNED NOT NULL AUTO_INCREMENT COMMENT '에러 로그 ID',
	`page_name` VARCHAR(100) NOT NULL COMMENT '페이지명',
	`error_no` INT(11) NOT NULL COMMENT '에러 번호',
	`error_msg` TEXT NOT NULL COMMENT '에러 메세지',
	`error_call_date` DATETIME(0) NOT NULL COMMENT '에러 발생 일자',
	`log_date` DATETIME(0) NOT NULL DEFAULT CURRENT_TIMESTAMP COMMENT '로그 적재 일자',
PRIMARY KEY (`error_log_id`))
ENGINE = INNODB DEFAULT CHARSET=utf8
COMMENT = 'WEB 런타임 에러 로그';

*/    

SELECT *
FROM error_log
;

SHOW FULL COLUMNS FROM error_log;

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
AUTHOR : 허정호
DATE : 2021-08-25
DESC : 
- DB ERROR LOG INSERT 
RETURN VALUE :
'
proc_body : BEGIN
	INSERT error_log(proc_name, proc_step, sql_state, error_no, error_msg, proc_call_date, log_date)
	VALUES(p_proc_name, p_proc_step, p_sql_state, p_error_no, p_error_msg, p_dt5_now, NOW(0));
END;

# Web Error Log
DROP PROCEDURE IF EXISTS nkp.USP_ERROR_WEB_LOG_INPUT;
CREATE PROCEDURE nkp.USP_ERROR_WEB_LOG_INPUT(
	IN p_page_name VARCHAR(100),
	IN p_error_no INT(11),
	IN p_error_msg TEXT CHARSET utf8,
	IN p_dt5_now DATETIME(0)
)
#NOT DETERMINISTIC
DETERMINISTIC
SQL SECURITY DEFINER
CONTAINS SQL
COMMENT ' 
AUTHOR : 허정호
DATE : 2021-08-25
DESC : 
- DB ERROR LOG INSERT 
RETURN VALUE :
'
proc_body : 
BEGIN	
	INSERT error_web_log(page_name, error_no, error_msg, error_call_date)
	VALUES(p_page_name, p_error_no, p_error_msg, p_dt5_now);
END;


CALL USP_ERROR_WEB_LOG_INPUT('/insa/insa_card_01.asp' , 
'3704' , 
'개체가 닫혀 있으면 작업이 허용되지 않습니다.' , 
'2021-08-25 17:38::12'); 

SELECT *
FROM error_web_log;


# test
DROP PROCEDURE IF EXISTS nkp.USP_PERSON_INDIVIDUAL_MOD;
CREATE PROCEDURE nkp.USP_PERSON_INDIVIDUAL_MOD(		
	IN p_dt5_now DATETIME(0)
	#OUT err_state INT
)
#NOT DETERMINISTIC
DETERMINISTIC
SQL SECURITY DEFINER
CONTAINS SQL
COMMENT ' 
AUTHOR : 허정호
DATE : 2021-08-24
DESC : 
- 개인 정보 관리 > 인사기본사항 변경 처리
RETURN VALUE :
	0 = 에러가 없습니다.
	1 = 예상하지 않은 런 타임 오류가 발생하였습니다.	
'
proc_body : BEGIN
	# ERROR LOG
	DECLARE v_vch_proc_name VARCHAR(100) DEFAULT 'USP_PERSON_INDIVIDUAL_MOD';
    DECLARE v_iny_proc_step TINYINT UNSIGNED DEFAULT 0;
    DECLARE v_vch_sql_state VARCHAR(5);
    DECLARE v_int_error_no INT;
    DECLARE v_txt_error_msg TEXT;
    
    
	DECLARE state INT DEFAULT 0;
	
	DECLARE EXIT HANDLER FOR SQLEXCEPTION
	BEGIN
		GET DIAGNOSTICS CONDITION 1 v_vch_sql_state = RETURNED_SQLSTATE
            , v_int_error_no = MYSQL_ERRNO
            , v_txt_error_msg = MESSAGE_TEXT;
            
		ROLLBACK;
		
		# error_log insert
		CALL USP_ERROR_LOG_INPUT(v_vch_proc_name
			, v_iny_proc_step
			, v_vch_sql_state
			, v_int_error_no
			, v_txt_error_msg
			, p_dt5_now);
		
		SET state = -1;
		
		SELECT state;
	END; 
	
	START TRANSACTION;
	
	INSERT INTO TranTest SET col01 = '1234';
    INSERT INTO TranTest2 SET col012 = '1234';
    COMMIT;   
     
    SET state = 1;
    
    SELECT state;
END;


CALL NKP.USP_PERSON_INDIVIDUAL_MOD(NOW(0));

SELECT *
FROM TranTest;

SELECT *
FROM TranTest2;

SELECT *
FROM error_log;

SELECT *
FROM web_error_log;



