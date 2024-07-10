
## 인사 관리

1. 조직 정보 - emp_org_mst

조직 인덱스 - org_seq [int] : pk
조직 코드 - org_code [varchar(4)]
조직 등급 - org_level [varchar(2)] : 회사(00), 본부(01), 사업부(02), 상주처(03), 지사(04), 팀(05), 파트(06)
조직명 - org_name [varchar(30)] 

2. 사원 기본 정보 - emp_master

사원 인덱스 - emp_seq [int] : pk
사원 번호 - emp_no [varchar(6)]
비밀 번호 - emp_pwd [varchar(20)]
직원 구분 - emp_type [char(1)] : 인턴(0), 정직원(1), 임원(2), 계약직(3)
직원 권한 - emp_grade [char(1)] : 퇴사자(0), 사용자(1), 관리자(2)

성별
주민번호 앞자리
주민번호 뒷자리
증명사진 이미지
최초 입사일
입사일
근속 일자
퇴직가산일
퇴사 일자

소속 발령일


조직 인덱스 - org_seq[int] : fk

3. 개인 정보 - emp_member

개인 정보 인덱스 - mem_seq [int] : pk
아이디 - mem_id [varchar(20)] : unique
비밀번호 - mem_pwd [varchar(20)]
사원명 - mem_name [varchar(20)]
영문이름 - mem_ename [varchar(30)]

사원 인덱스 - emp_seq [int] : fk

## 게시판

1. emp_board - 사내 게시판

게시판 인덱스 - board_seq [int]
게시 구분 - board_type [char(1)] : 사내(0), 공지(1)
제목 - board_title [varchar(100)]
내용 - board_content [text]
첨부파일 - att_file [varchar(100)]
조회수 - read_cnt [int]
수정 아이디 - modify_id [varchar(6)]
작성 일자 - create_date [datetime]
수정 일자 - update_date [datetime]

작성자 - emp_seq [int] : fk


## 급여 관리

## 비용 관리

## 회계 관리

## 영업 관리

