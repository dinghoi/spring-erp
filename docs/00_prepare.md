## 프로젝트 도구 관리

1. Slack (Communication) - 슬랙 책갈피를 통해 연결

2. Notion (프로젝트 관리 - 일정, 기록)

3. github 레포 (ERP) (소스코드 관리)

4. DB - MariaDB 10.6.x / MySql 8.0.x

5. 데이터베이스 관리 도구 - DBeaver, ERDCloud, HeidiSQL

6. 개발 툴 - Visual Studio Code

7. UI/UX : Figma

8. 기획 : PPT, PDF


## Ground Rule

1. Daily Scrum (당일 일정 및 진행 내역은 Notion으로 정리)

2. Git branch convention (develop을 기본 브랜치로 사용)

   - master (최종 브랜치)
     
   - develop (개발용 브랜치)
     
   - feature (기능 추가 브랜치)
   
       feat/crawler/naver_news
       feat/model/nbc

3. Git commit 메시지 컨벤션

   - feat : 새로운 기능 추가
  
       FEAT: 네이버 뉴스기사 크롤러 추가
       feat: add crawler for naver news
       feat: update crawler for naver news

   - Fix : 버그 수정
     
       fix: date formatting error

   - Rename : 파일 혹은 폴더명 수정

   - Remove : 파일 삭제

4. 폴더/파일명 컨벤션
   
   - 폴더명 
     - 영문 소문자 구성
     - 축약어 사용
     - 특수문자 및 공백 x
     - 단어 사이 구분은 '-'(Hyphen)으로 구성
   
   - 파일명 
     - 단어 사이 구분은 '_'(Underbar)로 구성
     - 단어 사이 구분 외 나머지는 위와 동일
     
5. 변수명 컨벤션

   - snake case(스테이크 표기법 : DB에서 주로 사용, 전부 소문자, 단어 사이에 '_' 표시)
  
       nake_case_naming_convention
  
   - camel case (카멜 표기법 : JAVA 권장 표기법, 첫 단어 소문자, 첫 단어 제외하고 첫 글자 대문자 표시)
    
       camelCaseNamingConvention
       
=====================

# Stacks

1. Programming tool
- Code : Visual Studio Code
- DB : HeidiSQL, DBeaver, ERDCloud
- UI/UX : Figma
- Proposal : PPT, PDF

2. Development
- Font-End : 
- Back-End :
- Infra :


# Proposal

- 업무 계획서
- 기대 결과물 : Web Server / API Server / DB Server