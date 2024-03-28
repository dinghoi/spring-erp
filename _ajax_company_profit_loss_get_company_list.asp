<!--#include virtual="/include/JSON_2.0.4.asp"-->
<!--#include virtual="/include/JSON_UTIL_0.1.1.asp"-->

<%
    On Error Resume Next

    Dim DbConnect
    DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=kwon_admin(*)14;"

    ' https://code.google.com/archive/p/aspjson/ 
    
    Set Dbconn = Server.CreateObject("ADODB.Connection")
    Set Rs     = Server.CreateObject("ADODB.Recordset")

    Dbconn.open DbConnect

    date1 = request("date1") ' 부터
    date2 = request("date2") ' 까지
    saupbu = unescape(request("saupbu")) ' 사업부

    Sql="SHOW TABLES LIKE 'temp_person_cost'"
    rs.Open Sql, Dbconn, 1

    if not rs.eof then

        sql = "drop table  temp_person_cost"
        'Response.write sql       
        dbconn.execute(sql)

    end if 

    rs.close()
    Set rs = Nothing


    sql = "  CREATE TABLE temp_person_cost                                                                 " & chr(13) &_
          "  AS                                                                                            " & chr(13) &_
          "       SELECT '매입세금계산서' as gubun                                                         " & chr(13) &_
          "            , a.emp_no as emp_no                                                                " & chr(13) &_
          "            , a.company as org_name                                                             " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.cost,concat(a.account,' '                                                       " & chr(13) &_
          "            , a.account_item) as cost_detail                                                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM general_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE a.tax_bill_yn = 'Y'                                                               " & chr(13) &_
          "          AND (a.slip_date >= '"& date1 &"' AND a.slip_date <= '"& date2 &"')                   " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '야특근' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.work_date as slip_date                                                          " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , a.work_item as slip_memo                                                          " & chr(13) &_
          "            , a.overtime_amt as cost                                                            " & chr(13) &_
          "            , a.work_gubun as cost_detail                                                       " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM overtime  a                                                                       " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.work_date,1,4), SUBSTRING(a.work_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.work_date >= '"& date1 &"' AND a.work_date <= '"& date2 &"')                   " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '일반경비' as gubun                                                               " & chr(13) &_
          "            , a.emp_no as emp_no                                                                " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.cost                                                                            " & chr(13) &_
          "            , concat(a.account,' ',a.account_item) as cost_detail                               " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM general_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.slip_date >= '"& date1 &"' AND a.slip_date <= '"& date2 &"')                   " & chr(13) &_
          "          AND a.tax_bill_yn <> 'Y'                                                              " & chr(13) &_
          "          AND a.slip_gubun = '비용'                                                             " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '대중교통' as gubun                                                               " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.company,' ',a.run_memo) as slip_memo                                     " & chr(13) &_
          "            , a.fare as cost                                                                    " & chr(13) &_
          "            , a.transit as cost_detail                                                          " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '"& date1 &"' AND a.run_date <= '"& date2 &"')                     " & chr(13) &_
          "          AND a.car_owner = '대중교통'                                                          " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '주행거리' as gubun                                                               " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.far as cost                                                                     " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '"& date1 &"' AND a.run_date <= '"& date2 &"')                     " & chr(13) &_
          "          AND a.car_owner = '개인'                                                              " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '주유비' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.oil_price as cost                                                               " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '"& date1 &"' AND a.run_date <= '"& date2 &"')                     " & chr(13) &_
          "          AND a.car_owner = '회사'                                                              " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '주차료' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.parking as cost                                                                 " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '"& date1 &"' AND a.run_date <= '"& date2 &"')                     " & chr(13) &_
          "          AND a.parking > 0                                                                     " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '통행료' as gubun                                                                 " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.toll as cost                                                                    " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '"& date1 &"' AND a.run_date <= '"& date2 &"')                     " & chr(13) &_
          "          AND a.toll > 0                                                                        " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '차량수리비' as gubun                                                             " & chr(13) &_
          "            , a.mg_ce_id as emp_no                                                              " & chr(13) &_
          "            , a.org_name as org_name                                                            " & chr(13) &_
          "            , a.run_date as slip_date                                                           " & chr(13) &_
          "            , a.user_name                                                                       " & chr(13) &_
          "            , a.user_grade                                                                      " & chr(13) &_
          "            , concat(a.start_company,' -> ',a.end_company) as slip_memo                         " & chr(13) &_
          "            , a.repair_cost as cost                                                             " & chr(13) &_
          "            , concat(a.car_owner,' ',a.car_no,' ',a.oil_kind) as cost_detail                    " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM transit_cost a                                                                    " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.run_date,1,4), SUBSTRING(a.run_date,6,2))       " & chr(13) &_
          "          AND em.emp_no = a.mg_ce_id                                                            " & chr(13) &_
          "        WHERE (a.cancel_yn = 'N')                                                               " & chr(13) &_
          "          AND (a.run_date >= '"& date1 &"' AND a.run_date <= '"& date2 &"')                     " & chr(13) &_
          "          AND a.car_owner = '회사'                                                              " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT a.emp_no as emp_no                                                                " & chr(13) &_
          "            , '주유카드' as gubun                                                               " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.price as cost                                                                   " & chr(13) &_
          "            , concat(a.account,' ',a.account_item) as cost_detail                               " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM card_slip  a                                                                      " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE a.card_type like '%주유%'                                                         " & chr(13) &_
          "          AND (a.slip_date >= '"& date1 &"' AND a.slip_date <= '"& date2 &"')                   " & chr(13) &_
          " UNION ALL                                                                                      " & chr(13) &_
          "       SELECT '법인카드' as gubun                                                               " & chr(13) &_
          "            , a.emp_no as emp_no                                                                " & chr(13) &_
          "            , a.org_name                                                                        " & chr(13) &_
          "            , a.slip_date                                                                       " & chr(13) &_
          "            , a.emp_name as user_name                                                           " & chr(13) &_
          "            , a.emp_grade as user_grade                                                         " & chr(13) &_
          "            , a.customer as slip_memo                                                           " & chr(13) &_
          "            , a.cost as cost                                                                    " & chr(13) &_
          "            , concat(a.account,' ',a.account_item) as cost_detail                               " & chr(13) &_
          "            , em.emp_saupbu                                                                     " & chr(13) &_
          "            , em.cost_center                                                                    " & chr(13) &_
          "         FROM card_slip a                                                                       " & chr(13) &_
          "   INNER JOIN emp_master_month em                                                               " & chr(13) &_
          "           ON em.emp_month = concat(SUBSTRING(a.slip_date,1,4), SUBSTRING(a.slip_date,6,2))     " & chr(13) &_
          "          AND em.emp_no = a.emp_no                                                              " & chr(13) &_
          "        WHERE a.card_type not like '%주유%'                                                     " & chr(13) &_
          "          AND (a.slip_date >= '"& date1 &"' AND a.slip_date <= '"& date2 &"')                   "
    'Response.write sql       
    dbconn.execute(sql)
      

    if saupbu <> "" then
        sql_cond = "   and saupbu = '"& saupbu &"' " 
    Else
        sql_cond = " "
    end if

    sql = "  SELECT saupbu, company                                                   " & chr(13) &_ 
          "       , CONCAT( (  SELECT b.sales_memo                                    " & chr(13) &_ 
          "                      FROM saupbu_sales  b                                 " & chr(13) &_ 
          "                     WHERE b.saupbu = a.saupbu                             " & chr(13) &_ 
          "                       AND b.company = a.company                           " & chr(13) &_ 
          "                  ORDER BY sales_date                                      " & chr(13) &_ 
          "                     LIMIT 1                                               " & chr(13) &_ 
          "                 )                                                         " & chr(13) &_ 
          "               , ' 외 '                                                    " & chr(13) &_            
          "               , count(*)-1                                                " & chr(13) &_            
          "               , '건'                                                      " & chr(13) &_            
          "               ) sales_memo                                                " & chr(13) &_ 
          "       , (  SELECT sum(b.sales_amt) sales_amt                              " & chr(13) &_                                                                                                          
          "              FROM saupbu_sales  b                                         " & chr(13) &_                                                                                                           
          "             WHERE b.saupbu  = a.saupbu                                    " & chr(13) &_                                                                                                           
          "               AND b.company = a.company                                   " & chr(13) &_                                                                                                                     
          "         ) sales_amt                                                       " & chr(13) &_ 
          "       , count(*) cnt                                                      " & chr(13) &_ 
          "     FROM saupbu_sales a                                                   " & chr(13) &_ 
          "    WHERE ( sales_date >= '"& date1 &"' AND sales_date <= '"& date2 &"' )  " & chr(13) &_ 
          "      AND saupbu <> '' AND company <> ''                                   " & chr(13) &_ 
          sql_cond                                                                      & chr(13) &_ 
          " GROUP BY saupbu, company                                                  "

    'Response.write sql 
    QueryToJSON(dbconn, sql).Flush

    If Err.number <> 0 Then     '오류 발생 시 이 부분 실행
	    Response.Write "" & Err.Source & "<br>"
	    Response.Write "오류 번호 : " & Err.number & "<br>"
	    Response.Write "내용 : " & Err.Description & "<br>"
	Else
	    ' Response.Write "오류가 없습니다."
	End If

%>

