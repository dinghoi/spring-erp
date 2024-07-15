### 조직 정보

|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK|org_code|조직 코드|varchar(4)|N||
||org_level|조직 구분|varchar(10)|Y||
||org_company|소속 회사|varchar(30)|Y||
||org_bonbu|소속 본부|varchar(30)|Y||
||org_saupbu|소속 사업부|varchar(30)|Y||
||org_team|소속 팀|varchar(30)|Y||
||org_name|조직명|varchar(30)|Y||
||org_reside_place|상주처|varchar(30)|Y||
||org_reside_company|상주 회사|varchar(50)|Y||
||org_cost_group|비용 그룹|varchar(50)|Y||
||org_cost_center|비용 구분|varchar(20)|Y||
||org_emp_no|조직장 사번|varchar(6)|N||
||org_owner_org|상위 조직 코드|varchar(4)|Y||
||org_table_org|T.O|int|Y||
||org_end_date|조직 폐쇄 일자|date|Y||
||trade_code|Field11|varchar(5)|Y||
||create_datetime|등록 일자|datetime|N|now()|
||update_datetime|수정 일자|datetime|N|now()|
||reg_user|등록/수정 사용자|varchar(6)|N||

### 직원 상세 정보
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|emp_seq|Key|int|N||
||emp_type|직원 구분|char(1)|N|'1'|계약직(0), 정직(1)|
||emp_extension_no|내선 번호|varchar(11)|Y||
||emp_in_date|입사일|date|Y||
||emp_gunsok_date|근속 일자|date|Y||
||emp_yuncha_date|연차가산일|date|Y||
||emp_end_gisan|퇴직기산일|date|Y||
||emp_end_date|퇴사일자|date|Y||
||emp_org_baldate|발령일|date|Y||
||emp_stay_code|실근무지코드|varchar(4)|Y||
||emp_stay_name|실근무지/주소|varchar(30)|Y||
||emp_reside_place|상주처|varchar(30)|Y||
||emp_reside_company|상주 회사|varchar(30)|Y||
||emp_grade|직급|varchar(10)|Y||
||emp_grade_date|승진 일자|date|Y||
||emp_job|직위|varchar(10)|Y||
||emp_position|직책|varchar(10)|Y||
||emp_jikgun|직군|varchar(20)|Y||
||emp_jikmu|직무|varchar(20)|Y||
||emp_sawoo_yn|사우회 가입 여부|char(1)|Y|'1'|미가입(0), 가입(1)|
||emp_pay_yn|급여 대상 여부|char(1)|N|'1'|급여 대상 아님(0), 급여 대상(1)|
||emp_pay_type|소득 구분|varchar(1)|Y||
||cost_group|비용센타그룹|varchar(50)|Y||
||cost_center|비용 배분 구분|varchar(20)|Y||
||mg_saupbu|관리 본부|varchar(30)|Y||
||cost_except|손익 대상|varchar(2)|Y||
||create_datetime|등록 일자|datetime|Y||
||update_datetime|수정일자|datetime|Y||
||reg_user|등록/수정자|varchar(6)|N||

### 사내/공지 게시판
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|board_seq|Key|int|N||
||board_type|게시 타입|char(1)|N|'1'|사내(0), 공지(1)|
||board_title|제목|varchar(100)|N||
||board_content|내용|text|N||
||att_file|첨부 파일|varchar(100)|Y||
||read_cnt|조회수|int|N|0||
||created_date|작성 일자|datetime|N|now()|
||updated_date|수정 일자|datetime|N|now()|
||reg_id|등록/수정 아이디|varchar(6)|N||

### 직원 마스터
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|emp_no|사원 번호|varchar(6)|N||
||emp_pwd|비밀 번호|varchar(20)|N||
||emp_grade|사용 권한|char(1)|N|'0'|권한 없음(0), 일반 권한(1), 관리자 권한(2)|
||insa_grade|인사 관리 권한|char(1)|N|'0'|권한 없음(0), 사용 가능(1)|
||pay_grade|급여 관리 권한|char(1)|N|'0'|권한 없음(0), 사용 가능(1)|
||cost_grade|비용 관리 권한|char(1)|N|'0'|권한 없음(0), 사용 가능(1)|
||account_grade|회계 관리 권한|char(1)|N|'0'|권한 없음(0), 사용 가능(1)|
||met_grade|자재 관리 권한|char(1)|N|'0'|권한 없음(0), 사용 가능(1)|
||sales_grade|영업 관리 권한|char(1)|N|'0'|권한 없음(0), 사용 가능(1)|
||created_datetime|등록 일자|datetime|N|now()||
||udpated_datetime|수정 일자|datetime|N|now()||
||reg_user|등록/수정 사용자|varchar(6)|N||

### 가족 사항
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|family_seq|Key|int|N||
||family_emp_no|사원 번호|varchar(3)|Y||
||family_rel|본인 관계|varchar(12)|Y||
||family_name|가족 성명|varchar(10)|Y||
||family_birthday|생년월일|date|Y||
||family_birthday_id|양력/음력 구분|varchar(2)|Y||
||family_job|직업|varchar(20)|Y||
||family_live|동거 여부|varchar(4)|Y||
||family_person1|주민번호 앞자리|varchar(6)|Y||
||family_person2|주민번호 뒷자리|varchar(7)|Y||
||family_tel_ddd|전화 지역번호|varchar(3)|Y||
||family_tel_no1|전화 번호 중간자리|varchar(4)|Y||
||family_tel_no2|전화번호 뒷자리|varchar(4)|Y||
||family_support_yn|부양여부|varchar(1)|Y||
||family_national|출생지|varchar(10)|Y||
||family_disab|장애인 구분|varchar(1)|Y||
||family_merit|국가유공자|varchar(1)|Y|||
||family_serius|중증환자|varchar(1)|Y|||
||family_pensioner|국민기초생활수급자|varchar(1)|Y||
||family_witak|위탁아동|varchar(1)|Y||
||family_holt|입양여부|varchar(1)|Y||
||family_holt_date|입양일자|date|Y||
||family_children|자녀 양육|varchar(1)|Y||
||family_reg_date|등록 일자|datetime|Y||
||family_reg_user|등록자 명|varchar(20)|Y||
||family_mod_date|수정일자|datetime|Y||
||family_mod_user|수정자 명|varchar(20)|Y||

### 학력 사항
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|sch_seq|Key|int|N||
||sch_empno|사원 번호|varchar(6)|Y||
||sch_start_date|입학 일자|date|Y||
||sch_end_date|졸업 일자|date|Y||
||sch_school_name|학교 명|varchar(30)|Y||
||sch_dept|학과|varchar(30)|Y||
||sch_major|전공|varchar(30)|Y||
||sch_sub_major|부전공|varchar(30)|Y||
||sch_degree|학위|varchar(10)|Y||
||sch_finish|졸업 여부|varchar(10)|Y||
||sch_comment|비고|varchar(30)|Y||
||sch_reg_date|등록 일자|datetime|Y||
||sch_reg_user|등록자 명|varchar(20)|Y||
||sch_mod_date|수정 일자|datetime|Y||
||sch_mod_user|수정자 명|varchar(20)|Y||

### 경력 사항
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|career_seq|Key|int|N||
||career_empno|사원 번호|varchar(6)|Y||
||career_join_date|재직시작기간|date|Y||
||career_end_date|재직 종료 기간|date|Y||
||career_office|회사명|varchar(30)|Y||
||career_dept|부서명|varchar(30)|Y||
||career_position|직위/직책|varchar(20)|Y||
||career_task|담당 업무|varchar(50)|Y||
||career_reg_date|등록 일자|datetime|Y||
||career_reg_user|등록자 명|varchar(20)|Y||
||career_mod_date|수정 일자|datetime|Y||
||career_mod_user|수정자 명|varchar(20)|Y||

### 자격 사항
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|qual_seq|Key|int|N||
||qual_empno|사원 번호|varchar(6)|Y||
||qual_type|순번|varchar(30)|Y||
||qual_grade|자격 종목|varchar(10)|Y||
||qual_pass_date|합격 일자|date|Y||
||qual_org|발급 기관|varchar(30)|Y||
||qual_no|자격등록번호|varchar(30)|Y||
||qual_passport|경력수첩번호|varchar(20)|Y||
||qual_pay_id|자격 수당 여부|varchar(1)|Y||
||qual_reg_date|등록 일자|dateetime|Y||
||qual_reg_user|등록자 명|varchar(20)|Y||
||qual_mod_date|수정 일자|datetime|Y||
||qual_mod_user|수정자 명|varchar(20)|Y||

### 교육 사항
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|edu_seq|Key||int|N||
||edu_empno|사원 번호|varchar(6)|Y||
||edu_name|교육과정명|varchar(30)|Y||
||edu_office|교육 기관|varchar(30)|Y||
||edu_finish_no|교육수료증번호|varchar(20)|Y||
||edu_start_date|교육 시작 일자|date|Y||
||edu_end_date|교육 종료 일자|date|Y||
||edu_pay|교육 비용|int|Y||
||edu_comment|교육 주요 내용|varchar(100)|Y||
||edu_reg_date|등록 일자|datetime|Y||
||edu_reg_user|등록자명|varchar(20)|Y||
||edu_mod_date|수정 일자|datetime|Y||
||edu_mod_user|수정자 명|varchar(20)|Y||

### 어학 사항
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|lang_seq|Key|int|N||
||lang_empno|사원 번호|varchar(6)|Y||
||lang_id|어학 구분|varchar(20)|Y||
||lang_id_type|어학 종류|varchar(20)|Y||
||lang_point|점수|varchar(3)|Y||
||lang_grade|급수|varchar(10)|Y||
||lang_get_date|취득일|date|Y||
||lang_reg_date|등록 일자|datetime|Y||
||lang_reg_user|등록자 명|varchar(20)|Y||
||lang_mod_date|수정 일자|datetime|Y||
||lang_mod_user|수정자 명|varchar(20)|Y||

### 관리 코드
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
PK|emp_etc_seq|Key|int|N||
||emp_etc_code|관리 코드|varchar(4)|Y||
||emp_etc_type|코드 구분|varchar(2)|Y||
||emp_type_name|Field3|varchar(20)|Y||
||emp_etc_name|Field4|varchar(30)|Y||
||emp_etc_group|Field5|varchar(2)|Y||
||emp_group_name|Field6|varchar(20)|Y||
||emp_mg_group|Field7|varchar(1)|Y||
||emp_used_sw|Field8|varchar(1)|Y||
||emp_comment|기타 내용|varchar(50)|Y||
||emp_tax_id|세금 구분|varchar(1)|Y||
||emp_payend_date|급여 지급월 마감월|varchar(6)|Y||
||emp_payend_yn|급여지급월 마감여부|varchar(1)|Y||

### 직원 월급여 지급
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK|pmg_seq|Key|int|N||
||pmg_yymm|지급 년월|varchar(6)|Y||
||pmg_id|급여 구분|varchar(1)|Y||
||pmg_emp_no|사원 번호|varchar(7)|Y||
||pmg_date|지급일|date|Y||
||pmg_in_date|입사일|date|Y||
||pmg_emp_name|사원명|varchar(20)|Y||
||pmg_emp_type|사원 구분|varchar(10)|Y||
||pmg_org_code|조직 코드|varchar(4)|Y||
||pmg_company|소속 회사명|varchar(20)|Y||
||pmg_bonbu|소속 본부|varchar(30)|Y||
||pmg_saupbu|소속 사업부|varchar(30)|Y||
||pmg_team|소속 팀|varchar(30)|Y||
||pmg_org_name|조직명|varchar(30)|Y||
||pmg_reside_place|상주처|varchar(30)|Y||
||pmg_reside_company|상주 회사|varchar(50)|Y||
||pmg_grade|직급|varchar(10)|Y||
||pmg_position|직책|varchar(10)|Y||
||pmg_base_pay|기본급|int|Y||
||pmg_meals_pay|식대|int|Y||
||pmg_postage_pay|통신비|int|Y||
||pmg_re_pay|소급 급여|int|Y||
||pmg_overtime_pay|연장근로수당|int\|Y||
||pmg_car_pay|주차지원금|int|Y||
||pmg_position_pay|직책수당|int|Y||
||pmg_custom_pay|고객관리 수당|int|Y||
||pmg_job_pay|직무보조비|ing|Y||
||pmg_job_support|업무장려비|int|Y||
||pmg_jisa_pay|복지사근무비|int|Y||
||pmg_long_pay|근속수당|ing|Y||
||pmg_disabled_pay|장애인 수당|int|Y||
||pmg_family_pay|가족 수당|int|Y||
||pmg_school_pay|학자금|int|Y||
||pmg_qual_pay|자격증수당|int|Y||
||pmg_bonus_pay|상여금|int|Y||
||pmg_perf_pay|성과급|int|Y||
||pmg_other_pay1|기타지급1|int|Y||
||pmg_other_pay2|기타지급2||Y||
||pmg_other_pay3|기타지급3|int|Y||
||pmg_tax_yes|과세|int|Y||
||pmt_tax_no|비과세|int|Y||
||pmg_tax_reduced|감면세액|int|Y||
||pmg_give_total|지급총액|int|Y||
||pmg_bank_name|입금 은행|varchar(20)|Y||
||pmg_account_no|계좌 번호|varchar(30)|Y||
||pmg_account_holder|예금주|varchar(20)|Y||
||cost_group|비용센타그룹|varchar(50)|Y||
||cost_center|비용배분구분|varchar(20)|Y||
||pmg_comment|기타사항|varchar(100)|Y||
||pmg_reg_date|등록 일자|datetime|Y||
||pmg_reg_user|등록자 명|varchar(20)|Y||
||pmg_mod_date|수정일자|datetime|Y||
||pmg_mod_user|수정자명|varchar(20)|Y||
||mg_saupbu|관리 사업부|varchar(30)|Y||
||pmg_research_pay|연구비|int|Y||

### 직원 월급여 공제
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK|de_seq|Key|int|N||
||de_yymm|Field|varchar(6)|Y||
||de_id|급여  구분|varchar(1)|Y||
||de_emp_no|Field3|varchar(7)|Y||
||de_date|지급일|date|Y||
||de_emp_name|Field5|varchar(20)|Y||
||de_emp_type|Field6|varchar(10)|Y||
||de_org_code|Field7|varchar(4)|Y||
||de_company|Field8|varchar(20)|Y||
||de_bonbu|Field9|varchar(30)|Y||
||de_saupbu|Field10|varchar(30)|Y||
||de_team|Field11|varchar(30)|Y||
||de_org_name|Field12|varchar(30)|Y||
||de_reside_place|Field13|varchar(30)|Y||
||de_reside_company|Field14|varchar(30)|Y||
||de_grade|Field15|varchar(10)|Y||
||de_position|Field16|varchar(10)|Y||
||de_nps_amt|국민연급|int|Y||
||de_nhis_amt|건강보험|int|Y||
||de_epi_amt|고용보험|int|Y||
||de_longcare_amt|장기요양보험|int|Y||
||de_income_tax|소득세|int|Y||
||de_wetax|지방소득세|int|Y||
||de_special_tax|농특세|int|Y||
||de_year_incom_tax|연말정산소득세|int|Y||
||de_year_wetax|연말정산지방소득세|int|Y||
||de_year_special_tax|연말정산농특세|int|Y||
||de_year_incom_tax2|연말정산재정산소득세|int|Y||
||de_year_wetax2|연말정산재정산지방세|int|Y||
||de_saving_amt|재형저축|int|Y||
||de_sawo_amt|경조회비|int|Y||
||de_johab_amt|노동조합비|int|Y||
||de_hyubjo_amt|협조비|int|Y||
||de_school_amt|학자금대출|int|Y||
||de_other_amt1|기타공제1|int|Y||
||de_other_amt2|기타공제2|int|Y||
||de_other_amt3|기타공제3|int|Y||
||de_nhis_bla_amt|건강보험료정산|int|Y||
||de_long_bla_amt|장기요양보험정산|int|Y||
||de_deduct_total|공제총액|int|Y||
||cost_group|Field23|varchar(30)|Y||
||cost_center|Field24|varchar(20)|Y||
||de_reg_date|등록일자|datetime|Y||
||de_reg_user|등록자명|varchar(20)|Y||
||de_mod_date|수정일자|datetime|Y||
||de_mod_user|수정자명|varchar(20)|Y||

### 발령 사항
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|app_seq|Key|int|N||
||app_empno|사원 번호|varchar(6)|Y||
||app_id|Field|varchar(10)|Y||
||app_date|Field2|date|Y||
||app_emp_name|Field3|varchar(20)|Y||
||app_id_type|발령 유형|varchar(20)|Y||
||app_to_compamy|Field5|varchar(30)|Y||
||app_to_orgcode|Field6|varchar(4)|Y||
||app_to_org|Field7|varchar(30)|Y||
||app_to_grade|Field8|varchar(10)|Y||
||app_to_job|Field9|varchar(10)|Y||
||app_to_position|Field10|varchar(10)|Y||
||app_to_enddate|Field11|date|Y||
||app_be_company|Field12|varchar(30)|Y||
||app_be_orgcode|Field13|varchar(4)|Y||
||app_be_org|Field14|varchar(30)|Y||
||app_be_grade|Field15|varchar(10)|Y||
||app_be_job|Field16|varchar(10)|Y||
||app_be_position|Field17|varchar(10)|Y||
||app_be_enddate|Field18|date|Y||
||app_start_date|Field19|date|Y||
||app_finish_date|Field20|date|Y||
||app_reward|Field21|varchar(50)|Y||
||app_commit|Field22|varchar(50)|Y||
||app_bokjik_id|Field4|varchar(1)|Y||
||app_reg_date|Field23|datetime|Y||
||app_reg_user|Field24|varchar(20)|Y||

### 경조회 정보
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
|PK, FK|emp_no|사원 번호2|varchar(6)|N||
|PK, FK|org_code|조직 코드|varchar(4)|N||
|PK|sawo_seq|Key|int|N||
||sawo_empno|사원 번호|varchar(7)|Y||
||sawo_date|Field|date|Y||
||sawo_id|가입 구분|varchar(10)|Y||
||sawo_emp_name|가입자 명|varchar(20)|Y||
||sawo_out|탈퇴 구분|varchar(10)|Y||
||sawo_out_date|탈퇴 일자|date|Y||
||sawo_company|Field6|varchar(30)|Y||
||sawo_org_code|Field7|varchar(4)|Y||
||sawo_org_name|Field8|varchar(30)|Y||
||sawo_target|Field9|varchar(1)|Y||
||sawo_in_count|Field10|int|Y||
||sawo_in_pay|Field11|int|Y||
||sawo_give_count|Field2|int|Y||
||sawo_give_pay|Field3|int|Y||
||sawo_reg_date|등록 일자|dastetime|Y||
||sawo_reg_user|등록자 명|varchar(20)|Y||
||sawo_mod_date|수정 일자|datetime|Y||
||sawo_mod_user|수정자 명|varchar(20)|Y||

### 4대보험 기준등급설정
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
PK|insu_seq|Key|int|N||
||insu_yyyy|Field|varchar(4)|Y||
||insu_id|4대보험 구분|varchar(4)|Y||
||insu_class|Field3|varchar(2)|Y||
||insu_id_name|Field4|varchar(20)|Y||
||from_amt|Field5|int|Y||
||to_atm|Field6|int|Y||
||st_amt|Field7|int|Y||
||hap_rate|Field8|float|Y||
||emp_rate|Field9|float|Y||
||com_rate|Field10|float|Y||
||insu_comment|Field11|varchar(50)|Y||
||reg_date|등록일자|datetime|Y||
||reg_user|등록자 명|varchar(20)|Y||
||mod_date|수정 일자|datetime|Y||
||mod_user|수정자 명|varchar(20)|Y||

### 직원 연봉
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
PK|incom_seq|Key|int|N||
||incom_emp_no|사원 번호|varchar(6)|Y||
||incom_year|Field2|varchar(4)|Y||
||incom_emp_name|사원명|varchar(20)|Y||
||incom_in_date|입사일자|date|Y||
||incom_grade|Field5|varchar(20)|Y||
||incom_emp_type|직원 구분|varchar(10)|Y||
||imcom_pay_type|소득 구분|varchar(10)|Y||
||incom_company|소속 회사명|varchar(20)|Y||
||incom_org_code|조직 코드|varchar(4)|Y||
||incom_org_name|조직명|varchar(30)|Y||
||incom_base_pay|기본급|int|Y||
||incom_overtime_pay|연장근로수당|int|Y||
||incom_meals_pay|식대|int|Y||
||incom_serverance_pay|퇴직금|int|Y||
||incom_total_pay|총수령액|int|Y||
||incom_first3_percent|수습 급여|int|Y||
||incom_month_amount|과세평균소득월액|int|Y||
||incom_nps_amount|국민연금표준월액|int|Y||
||incom_nps|국민연금월납부금액|int|Y||
||incom_nhis_amount|건강보험표준월액|int|Y||
||incom_nhis|건강보험월납부금액|int|Y||
||incom_go_yn|고용보험적용여부|varchar(2)|Y||
||incom_san_yn|산재보험적용여부|varchar(2)|Y||
||incom_long_yn|장기요양보험가입여부|varchar(2)|Y||
||incom_incom_yn|청년소득세감면여부|varchar(2)|Y||
||incom_family_cnt|부양가족수|int|Y||
||incom_wife_yn|배우자 유무|varchar(1)|Y||
||incom_age20|20세이하|int|Y||
||incom_age60|60세 이상|int|Y||
||income_old|경로우대|int|Y||
||incom_disab|장애인|int|Y||
||incom_woman|부녀자|varchar(1)|Y||
||incom_retirement_bank|퇴직연금가입 금융기관|varchar(20)|Y||
||incom_reg_date|등록일자|datetime|Y||
||incom_reg_user|등록자명|varchar(20)|Y||
||incom_mod_date|수정일자|datetime|Y||
||incom_mod_user|수정자 명|varchar(20)|Y||

### 근로소득 간이세액
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
PK|inc_seq|Key|int|N||
||inc_yyyy|Field|varchar(4)|Y||
||inc_from_amt|Field2|int|Y||
||inc_to_amt|Field3|int|Y||
||inc_st_amt|Field4|int|Y||
||inc_incom1|Field5|int|Y||
||inc_incom2|Field6|int|Y||
||inc_incom3|Field7|int|Y||
||inc_incom4|Field8|int|Y||
||inc_incom5|Field9|int|Y||
||inc_incom6|Field10|int|Y||
||inc_incom7|Field11|int|Y||
||inc_incom8|Field12|int|Y||
||inc_incom9|Field13|int|Y||
||inc_incom10|Field14|int|Y||
||inc_incom11|Field15|int|Y||
||inc_incom12|Field16|int|Y||
||inc_comment|Field17|varchar(200)|Y||
||inc_reg_date|Field18|date|Y||
||inc_reg_user|varchar(20)||Y||

### 직원급여 은행계좌
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
PK|bank_seq|Key|int|N||
||emp_no|사원 번호|varchar(7)|Y||
||emp_name|사원명|varchar(20)|Y||
||person_no1|주민번호 앞자리|varchar(6)|Y||
||person_no2|주빈번호 뒷자리|varchar(7)|Y||
||emp_type|직원 구분|varchar(10)|Y||
||emp_pay_type|소득 구분|varchar(1)|Y||
||bank_code|은행 구분 코드|varchar(4)|Y||
||bank_name|은행명|varchar(20)|Y||
||account_no|계좌 번호|varchar(30)|Y||
||account_holder|예금주|varchar(20)|Y||
||reg_date|등록일자|datetime|Y||
||reg_user|등록자 명|varchar(20)|Y||
||mod_date|수정일자|datetime|Y||
||mod_user|수정자 명|varchar(20)|Y||

### 직원 개인 정보
|키|논리|물리|타입|Null|기본값|코멘트|
|----|----|----|----|----|----|----|
PK, FK|emp_no|사원 번호|varchar(6)|N||
PK, FK|||org_code|조직 코드|varchar(4)|N||
PK|mem_seq|Key|int|N||
||mem_name|직원명|varchar(20)|N||
||eng_name|영문 이름|varchar(30)|Y||
||jumin_no|주민 번호|varchar(13)|N||
||mem_email|회사 이메일|varchar(30)|Y||
||sex|성별|char(1)|Y|'1'|남(1), 여(0)
||birthday|생년월일|date|Y||
||birthday_id|양력/음력|char(1)|Y|'1'|양력(1), 음력(0)
||zip_code|우편 번호|varchar(12)|Y||
||sido|시/도|varchar(10)|Y||
||gugun|구/군|varchar(20)|Y||
||dong|동/읍|varchar(50)|Y||
||addr_detail|상세 주소|varchar(50)|Y||
||tel_no|전화 번호|varchar(11)|Y||
||phone_no|휴대폰 번호|varchar(11)|Y||
||military_id|병역 사항|varchar(10)|Y||
||military_start_date|입대 일자|date|Y||
||military_end_date|제대 일자|date|Y||
||military_grade|병역 유형/계급|varchar(10)|Y||
||military_except|면제 사유|varchar(20)|Y||
||hobby|취미|varchar(20)|Y||
||faith|종교|varchar(20)|Y||
||last_edu|최종 학력|varchar(30)|Y||
||marry_yn|결혼/미혼|char(1)|Y|'0'|미혼(0), 결혼(1)
||marry_date|결혼 기념일|date|Y||
||disable_yn|장애 여부|char(1)|Y|'0'|장애 없음(0), 장애 있음(1)
||disable_type|장애 유형|varchar(20)|Y||
||disable_grade|장애 등급|varchar(4)|Y||
||emergency_no|비상연락망|varchar(11)|Y||
||nation_code|내국인 구분|varchar(3)|Y||
||car_yn|차량 보유 여부|char(1)|N|'0'|미보유(0), 보유(1)
||id_photo|증명 사진|varchar(100)|Y||
||create_datetime|등록 일자|datetime|N|now()|
||update_datetime|수정 일자|datetime|N|now()|
||reg_user|등록/수정자|varchar(6)|N||

