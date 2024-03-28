<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'   on Error resume next

Server.ScriptTimeOut = 1200

dim saupbu_tab(10,3) ' 1.사업부명, 2.사업부의 인원수, 3.사업부의 매출

end_month=Request("end_month")
end_yn=Request("end_yn")

from_date = mid(end_month,1,4) + "-" + mid(end_month,5,2) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

mm = mid(end_month,5,2)
cost_year = mid(end_month,1,4)
cost_month = mid(end_month,5)

' 원격 5%, 방문 95%
won_per = 5
bang_per = 95

for i = 1 to 10
    saupbu_tab(i,1) = "" ' 1.사업부명
    saupbu_tab(i,2) = 0  ' 2.사업부의 인원수
    saupbu_tab(i,3) = 0  ' 3.사업부의 매출
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

sql = "select * from cost_end where end_month = '"&end_month&"' and (end_yn = 'Y') and (saupbu = '상주비용')"
Set rs_check = Dbconn.Execute (sql)	
if rs_check.eof or rs_check.bof then
	check_sw = "N"
  else
  	check_sw = "Y"
end if

if check_sw = "N" then
	response.write"<script language=javascript>"
	response.write"alert('상주비용 마감을 먼저하셔야 합니다 !!');"
	response.write"location.replace('cost_end_mg.asp');"
	response.write"</script>"
	Response.End
  else		
	response.write"<script language=javascript>"
	response.write"alert('마감처리중!!!');"
	response.write"</script>"

dbconn.BeginTrans

sql = "SELECT sum(cost_amt_"&mm&") as tot_cost  " & chr(13) &_    
      "  FROM company_cost                      " & chr(13) &_    
      " WHERE cost_year   = '"&cost_year&"'     " & chr(13) &_    
      "   AND cost_center = '부문공통비'        "
'Response.write "<pre>"& sql &"</pre><br>"          
Set rs=DbConn.Execute(SQL)
if not (rs.eof or rs.bof) then
    tot_cost = clng(rs("tot_cost")) ' (12) 총 부문공통비
end if
rs.close()

' 4개 사업부 총매출액
sales_date = left(end_month,4) & "-" &right(end_month,2)
sql = "  SELECT sum(cost_amt) as sum_cost_amt                " & chr(13) &_
      "    FROM saupbu_sales                                 " & chr(13) &_
      "   WHERE (    sales_date >= '" + from_date + "'       " & chr(13) &_    
      "          AND sales_date <= '" + to_date + "'  )      " & chr(13) &_    
      "     AND saupbu IN ( 'SI1사업부'                      " & chr(13) &_
      "                   , 'SI2사업부'                      " & chr(13) &_
      "                   , 'N/W 사업부'                     " & chr(13) &_
      "                   , '공공사업부' )                   "
'Response.write "<pre>"& sql &"</pre><br>"          
Rs.Open Sql, Dbconn, 1
if not (rs.eof or rs.bof) then
    tot_sum_cost_amt = CCur(rs("sum_cost_amt")) ' (10)전체 매출액 (4개사 합계)
end if  
rs.close()

' 전체 AS건수 (To Be)
sql = "    SELECT count(*) as tot_cnt                               " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_process <> '취소'                         " & chr(13) &_    
      "            and as_type    <> '야특근' )                     " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='부문공통비'                        "
'Response.write "<pre>"& sql &"</pre><br>"          
Set rs=DbConn.Execute(SQL)
tot_as_cnt = clng(rs("tot_cnt")) ' (11)전체 AS건수 (To Be)
if tot_as_cnt = "" or isnull(tot_as_cnt) then
    tot_as_cnt = 0
end if
rs.close()

over_less =  tot_cost - (tot_as_cnt * 40000) ' (2)과부족분 ((12)총 부문공통비 - ((11)총 AS 건수 * 40000) )


' 고객사별 AS 건수를 구한다
sql = "    SELECT company, count(*) as as_cnt                       " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_process <> '취소'                         " & chr(13) &_    
      "            AND as_type    <> '야특근'  )                    " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='부문공통비'                        " & chr(13) &_    
      "  GROUP BY company                                           " & chr(13) &_    
      "  ORDER BY company Asc                                       "
'Response.write "<pre>"& sql &"</pre><br>"          
Rs.Open Sql, Dbconn, 1
do until rs.eof 

    sql = "select saupbu from trade where trade_name = '"&rs("company")&"'" 
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        saupbu = "Error"
    else
        saupbu = rs_trade("saupbu")
    end if
    rs_trade.close()

    sql = "  SELECT company, saupbu                              " & chr(13) &_
          "       , ifnull(sum(cost_amt),0) as sum_cost_amt      " & chr(13) &_
          "    FROM saupbu_sales                                 " & chr(13) &_
          "   WHERE (    sales_date >= '" + from_date + "'       " & chr(13) &_    
          "          AND sales_date <= '" + to_date + "'  )      " & chr(13) &_    
          "     AND saupbu  = '"&saupbu&"'                       " & chr(13) &_
          "     AND company = '"&rs("company")&"'                " 
'Response.write "<pre>"& sql &"</pre><br>"
    rs_etc.Open Sql, Dbconn, 1
    if not (rs_etc.eof or rs_etc.bof) then
        sum_cost_amt = Clng(rs_etc("sum_cost_amt"))  ' 회사별,사업부별 매출액
        if (sum_cost_amt=0) then
            sum_cost_per = 0
        else
            sum_cost_per = sum_cost_amt / tot_sum_cost_amt * 100.0  ' 회사별,사업부별 매출액 / 4개 사업부 총매출액 (%)
        end if
    end if
    rs_etc.close

    divide_amt_1 = CInt(rs("as_cnt")) * 40000                    ' (3)1차 배부금액 = (회사별사업부별)AS건수*표준단가
    divide_amt_2 = (sum_cost_amt / tot_sum_cost_amt) * over_less ' (4)2차 배부금액 = (회사별사업부별)매출액/4개사업부총매출액)*과부족분
'Response.write "<pre>(sum_cost_amt / tot_sum_cost_amt) * over_less "& "("&sum_cost_amt&" / "&tot_sum_cost_amt&") *"& over_less &"</pre><br>" 
    cost_amt = divide_amt_1 + divide_amt_2 ' (5)부문공통비 = 1차 배부금액 + 2차 배부금액

    sql = "SELECT ifnull(sum(cost_amt_07),0) as saupbu_cost " & chr(13) &_    
          "  FROM company_cost                              " & chr(13) &_    
          " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_    
          "   AND cost_center = '부문공통비'                " & chr(13) &_    
          "   AND saupbu      = '"&saupbu&"'                " & chr(13) &_    
          "   AND company     = '"&rs("company")&"'         " 
'Response.write "<pre>"& sql &"</pre><br>"          
    set rs_etc=dbconn.execute(sql)
    if not (rs_etc.eof or rs_etc.bof) then
        saupbu_cost = clng(rs_etc("saupbu_cost")) ' 회사별사업부별 부문공통비
        if (saupbu_cost=0) then
            charge_per = 0
        else
            charge_per = saupbu_cost / tot_cost * 100.0  ' (6)차지율 = 회사별사업부별 부문공통비 / 총 부문공통비 (%)
        end if
    end if
    rs_etc.close()

  
  charge_per = cost_amt / tot_cost * 100.0  ' (6)차지율 = 회사별사업부별 부문공통비 / 총 부문공통비 (%)
'Response.write "<pre>(6)차지율 = cost_amt / tot_cost * 100.0 = "& cost_amt &"/"& tot_cost &"*"& 100.0  &"</pre><br>"          

    
    sql = "SELECT * FROM company_asunit WHERE as_month = '"&end_month&"' AND as_company = '"&rs("company")&"'"
'Response.write "<pre>"& sql &"</pre><br>"              
    set rs_etc=dbconn.execute(sql)
    if rs_etc.eof or rs_etc.bof then
        sql="INSERT INTO company_asunit (as_month, as_company, saupbu, as_cnt, divide_amt_1, divide_amt_2, charge_per, cost_amt, reg_id, reg_name, reg_date) "&_
            "     VALUES ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&divide_amt_1&"',"&divide_amt_2&","&charge_per&","&cost_amt&",'"&user_id&"','"&user_name&"',now())"
        dbconn.execute(sql)
    else
        sql = "UPDATE company_asunit                     " & chr(13) &_    
              "   SET saupbu       = '"&saupbu&"'        " & chr(13) &_    
              "     , as_cnt       = '"&rs("as_cnt")&"'  " & chr(13) &_    
              "     , divide_amt_1 = '"&divide_amt_1&"'  " & chr(13) &_    
              "     , divide_amt_2 = '"&divide_amt_2&"'  " & chr(13) &_    
              "     , charge_per   = '"&charge_per&"'    " & chr(13) &_    
              "     , cost_amt     = "&cost_amt&"        " & chr(13) &_    
              " WHERE as_company = '" &rs("company")& "' " & chr(13) &_    
              "   AND as_month   = '" &end_month& "'     "
        dbconn.execute(sql)   
    end if
'Response.write "<pre>"& sql &"</pre><br>"            
    rs_etc.close()
    
    rs.movenext()
loop    
rs.close()

' 고객사별 손익 자료 생성
' 부문공통비 배분
' 처리전 zero
sql = "UPDATE company_profit_loss SET cost_amt_"&cost_month&"= '0' WHERE cost_year ='"&cost_year&"' AND (cost_center ='부문공통비') "
dbconn.execute(sql)

sql = "  SELECT as_company as company          " & chr(13) &_    
      "       , sum(charge_per) as charge_per  " & chr(13) &_    
      "    FROM company_asunit                 " & chr(13) &_    
      "   WHERE (as_month = '"&end_month&"')   " & chr(13) &_    
      "GROUP BY as_company                     "
Rs.Open Sql, Dbconn, 1
do until rs.eof
    charge_per = rs("charge_per")

    sql = "SELECT * FROM trade WHERE trade_name = '"&rs("company")&"'"
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        group_name = "Error"
    else
        group_name = rs_trade("group_name")
    end if        

    sql = "  SELECT cost_id, cost_detail, sum(cost_amt_"&cost_month&") as cost " & chr(13) &_    
          "    FROM company_cost                                               " & chr(13) &_    
          "   WHERE (cost_center = '부문공통비' )                              " & chr(13) &_    
          "     AND cost_year ='"&cost_year&"'                                 " & chr(13) &_    
          "GROUP BY cost_id, cost_detail                                       "
    rs_etc.Open sql, Dbconn, 1
    do until rs_etc.eof
        
        cost = int(charge_per * clng(rs_etc("cost")))

        sql = "SELECT * FROM company_profit_loss                " & chr(13) &_    
              " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_    
              "   AND company     = '"&rs("company")&"'         " & chr(13) &_    
              "   AND group_name  = '"&group_name&"'            " & chr(13) &_    
              "   AND cost_center = '부문공통비'                " & chr(13) &_    
              "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_    
              "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
        set rs_cost=dbconn.execute(sql)
        
        if rs_cost.eof or rs_cost.bof then
            sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&")                       " & chr(13) &_    
                  "     values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','부문공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&") "
            dbconn.execute(sql)
          else
            sql = "UPDATE company_profit_loss                       " & chr(13) &_    
                  "   SET cost_amt_"&cost_month&"="&cost&"          " & chr(13) &_    
                  " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_    
                  "   AND company     = '"&rs("company")&"'         " & chr(13) &_    
                  "   AND group_name  = '"&group_name&"'            " & chr(13) &_    
                  "   AND cost_center = '부문공통비'                " & chr(13) &_    
                  "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_    
                  "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
            dbconn.execute(sql)
        end if      
        
        rs_etc.movenext()
    loop
    rs_etc.close()
    rs.movenext()
loop
rs.close()
' 부문공통비 배분 끝
























' 원격
sql = " SELECT count(*) as tot_cnt FROM as_acpt                             " & chr(13) &_    
      "  WHERE acpt_man in ('조민순','주영미','한수정','안태환')            " & chr(13) &_    
      "    AND (Cast(visit_date as date) >= '" + from_date + "'             " & chr(13) &_    
      "    AND  Cast(visit_date as date) <= '" + to_date + "'  )            " & chr(13) &_    
      "    AND company not in ('코웨이','웅진씽크빅','웅진식품','롯데렌탈') "  
Set rs=DbConn.Execute(SQL)
won_cnt = clng(rs("tot_cnt")) ' 원격지원 총 건수
if won_cnt = "" or isnull(tot_cnt) then
    won_cnt = 0
end if
rs.close()

sql = "   SELECT company, count(*) as as_cnt FROM as_acpt                    " & chr(13) &_    
      "    WHERE acpt_man in ('조민순','주영미','한수정','안태환')           " & chr(13) &_    
      "      AND (Cast(visit_date as date) >= '" + from_date + "'            " & chr(13) &_    
      "      AND  Cast(visit_date as date) <= '" + to_date + "'  )           " & chr(13) &_    
      "      AND company not in('코웨이','웅진씽크빅','웅진식품','롯데렌탈') " & chr(13) &_    
      " GROUP BY company                                                     " & chr(13) &_    
      " ORDER BY company Asc                                                 "
Rs.Open Sql, Dbconn, 1

do until rs.eof 

    sql = "SELECT saupbu FROM trade WHERE trade_name = '"&rs("company")&"'" 
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        saupbu = "Error"
      else
        saupbu = rs_trade("saupbu")
    end if
    rs_trade.close()
    
    charge_per = clng(rs("as_cnt")) / won_cnt * won_per / 100 ' (사업부별 원격AS건수 / 총 원격AS건수) * 5%
    cost_amt = int(charge_per * tot_cost)     ' 차지율 * 총 부문공통비

    sql = "select * from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
    set rs_etc=dbconn.execute(sql)
    if not (rs_etc.eof or rs_etc.bof) then
        sql = "delete from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
        dbconn.execute(sql)   
    end if

    sql="INSERT INTO company_as (as_month,as_company,saupbu,remote_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date)  " & chr(13) &_    
        "     VALUES ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now()) "
    dbconn.execute(sql)
    
    rs.movenext()
loop    
rs.close()

' 원격외 (방문)

'  총 방문AS 건수를 구한다
sql = "    SELECT count(*) as tot_cnt                               " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_type    <> '원격처리'                     " & chr(13) &_    
      "            and as_process <> '취소'                         " & chr(13) &_    
      "            and as_type    <> '야특근' )                     " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='부문공통비'                        "
'sql = sql & " and ( team not  like '%지원%') and (company <> '기타' and company <> '케이원정보통신') "
'sql = sql & " and  mg_ce not in ('파라','박양준','김시욱','도현석','백지운','이종욱') "

Set rs=DbConn.Execute(SQL)
bang_cnt = clng(rs("tot_cnt")) ' 총 방문AS 건수
if bang_cnt = "" or isnull(tot_cnt) then
	bang_cnt = 0
end if
rs.close()

' 고객사별 방문AS 건수를 구한다
sql = "    SELECT company, count(*) as as_cnt                       " & chr(13) &_    
      "      FROM as_acpt a                                         " & chr(13) &_    
      "INNER JOIN emp_master_month b                                " & chr(13) &_    
      "        ON a.mg_ce_id  = b.emp_no                            " & chr(13) &_    
      "       AND b.emp_month = '" & end_month & "'                 " & chr(13) &_    
      "     WHERE (    as_type    <> '원격처리'                     " & chr(13) &_    
      "            AND as_process <> '취소'                         " & chr(13) &_    
      "            AND as_type    <> '야특근'  )                    " & chr(13) &_    
      "       AND reside       = '0'                                " & chr(13) &_    
      "       AND reside_place = ' '                                " & chr(13) &_    
      "       AND (Cast(visit_date as date) >= '" + from_date + "'  " & chr(13) &_    
      "       AND  Cast(visit_date as date) <= '" + to_date + "'  ) " & chr(13) &_    
      "       AND b.cost_center='부문공통비'                        " & chr(13) &_    
      "  GROUP BY company                                           " & chr(13) &_    
      "  ORDER BY company Asc                                        "
'sql = sql & " and ( team not  like '%지원%') and (company <> '기타' and company <> '케이원정보통신') "
'sql = sql & " and  mg_ce not in ('파라','박양준','김시욱','도현석','백지운','이종욱') 
Rs.Open Sql, Dbconn, 1

do until rs.eof 

    sql = "select saupbu from trade where trade_name = '"&rs("company")&"'" 
    set rs_trade=dbconn.execute(sql)
    if rs_trade.eof or rs_trade.bof then
        saupbu = "Error"
      else
        saupbu = rs_trade("saupbu")
    end if
    rs_trade.close()
    
    sql = "select * from company_as where as_month = '"&end_month&"' and as_company = '"&rs("company")&"'"
    set rs_etc=dbconn.execute(sql)
    if rs_etc.eof or rs_etc.bof then
        charge_per = clng(rs("as_cnt")) / bang_cnt * bang_per / 100  ' 고객사별 방문AS 건수 /  총 방문AS 건수 * 95%
        cost_amt = int(charge_per * tot_cost) ' 차지율 * 총 부문공통비
        sql="INSERT INTO company_as (as_month,as_company,saupbu,visit_cnt,charge_per,cost_amt,reg_id,reg_name,reg_date) "&_
            "     VALUES ('"&end_month&"','"&rs("company")&"','"&saupbu&"','"&rs("as_cnt")&"','"&charge_per&"',"&cost_amt&",'"&user_id&"','"&user_name&"',now())"
        dbconn.execute(sql)
    else
        charge_per = clng(rs("as_cnt")) / bang_cnt * bang_per / 100 + rs_etc("charge_per") ' 고객사별 방문AS 건수 /  총 방문AS 건수 * 95% + 차지율
        cost_amt = int(charge_per * tot_cost) ' 차지율 * 총 부문공통비
        sql = "UPDATE company_as SET visit_cnt='"&rs("as_cnt")&"', charge_per='"&charge_per&"', cost_amt="&cost_amt&_
              " WHERE as_company='" &rs("company")& "' and as_month = '" &end_month& "'"
        dbconn.execute(sql)   
    end if
    
    rs.movenext()
loop    
rs.close()
rs_etc.close()

' 사업부별 손익 자료 생성
' 부문공통비 배분
' 처리전 zero
sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='부문공통비') "
dbconn.execute(sql)

sql = "   SELECT saupbu, sum(charge_per) as charge_per " & chr(13) &_
      "     FROM company_as                            " & chr(13) &_
      "    WHERE (as_month = '"&end_month&"')          " & chr(13) &_
      " GROUP BY saupbu                                "
Rs.Open Sql, Dbconn, 1
do until rs.eof
    charge_per = rs("charge_per")

    sql = "  SELECT cost_id, cost_detail, sum(cost_amt_"&cost_month&") as cost    " & chr(13) &_
          "    FROM company_cost                                                  " & chr(13) &_
          "   WHERE (cost_center = '부문공통비' ) AND cost_year ='"&cost_year&"'  " & chr(13) &_
          "GROUP BY cost_id, cost_detail                                          "
    rs_etc.Open sql, Dbconn, 1
    do until rs_etc.eof
        
        cost = int(charge_per * clng(rs_etc("cost")))

        sql = "SELECT * FROM saupbu_profit_loss                 " & chr(13) &_
              " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_
              "   AND saupbu      = '"&rs("saupbu")&"'          " & chr(13) &_
              "   AND cost_center = '부문공통비'                " & chr(13) &_
              "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_
              "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
        set rs_cost=dbconn.execute(sql)
        
        if rs_cost.eof or rs_cost.bof then
            sql = "INSERT INTO saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&")                  " & chr(13) &_
                  "     VALUES ('"&cost_year&"','"&rs("saupbu")&"','부문공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&") "
            dbconn.execute(sql)
        else
            sql = "UPDATE saupbu_profit_loss                        " & chr(13) &_
                  "   SET cost_amt_"&cost_month&" = "&cost&"        " & chr(13) &_
                  " WHERE cost_year   = '"&cost_year&"'             " & chr(13) &_
                  "   AND saupbu      = '"&rs("saupbu")&"'          " & chr(13) &_
                  "   AND cost_center = '부문공통비'                " & chr(13) &_
                  "   AND cost_id     = '"&rs_etc("cost_id")&"'     " & chr(13) &_
                  "   AND cost_detail = '"&rs_etc("cost_detail")&"' "
            dbconn.execute(sql)
        end if      
        
        rs_etc.movenext()
    loop
    rs_etc.close()
    rs.movenext()
loop
rs.close()
' 부문공통비 배분 끝

' 고객사별 손익 자료 생성
' 부문공통비 배분
' 처리전 zero
'sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='부문공통비') "
'dbconn.execute(sql)
'
'sql = " select as_company as company, sum(charge_per) as charge_per from company_as Where (as_month = '"&end_month&"') GROUP BY as_company"
'Rs.Open Sql, Dbconn, 1
'do until rs.eof
'   charge_per = rs("charge_per")
'
'   sql = "select * from trade where trade_name = '"&rs("company")&"'"
'   set rs_trade=dbconn.execute(sql)
'   if rs_trade.eof or rs_trade.bof then
'       group_name = "Error"
'     else
'       group_name = rs_trade("group_name")
'   end if        
'
'   sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '부문공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
'   rs_etc.Open sql, Dbconn, 1
'   do until rs_etc.eof
'       
'       cost = int(charge_per * clng(rs_etc("cost")))
'
'       sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='부문공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
'       set rs_cost=dbconn.execute(sql)
'       
'       if rs_cost.eof or rs_cost.bof then
'           sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','부문공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
'           dbconn.execute(sql)
'         else
'           sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='부문공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
'           dbconn.execute(sql)
'       end if      
'       
'       rs_etc.movenext()
'   loop
'   rs_etc.close()
'   rs.movenext()
'loop
'rs.close()
' 부문공통비 배분 끝


' 추가 로직
' 사업부별 인원수,매출액 집계
sql = " select saupbu from sales_org where sales_year='" & cost_year & "' order by saupbu asc"
Rs.Open Sql, Dbconn, 1
i = 0
tot_person       = 0 ' 사업부 전체 인원
tot_saupbu_sales = 0 ' 사업부 전체 매출액

do until rs.eof 

	'sql = "select count(*) from pay_month_give where pmg_id = '1' and pmg_yymm = '"&end_month&"' and mg_saupbu ='"&rs("saupbu")&"'"
	
	''
	'' KDC사업부 사원 예외처리
	''
	
	' KDC사업부에 이름이 같은 건은 케이원정보통신에 소속된 사원을 cost_except = '2' 로 변경한다.
	sql = "SELECT emp_name, count(*)               "&chr(13)&_
        "  FROM                                  "&chr(13)&_
        "(                                       "&chr(13)&_
        "  SELECT B.*                            "&chr(13)&_
        "    FROM pay_month_give A               "&chr(13)&_
        "        ,emp_master_month B             "&chr(13)&_
        "   WHERE A.pmg_id = '1'                 "&chr(13)&_
        "     AND A.pmg_emp_no = B.emp_no        "&chr(13)&_
        "     AND B.cost_except in ('0','1','2') "&chr(13)&_
        "     AND A.pmg_yymm  = '"&end_month&"'  "&chr(13)&_
        "     AND B.emp_month = '"&end_month&"'  "&chr(13)&_
        "     AND A.mg_saupbu = 'KDC사업부'      "&chr(13)&_
        ") A                                     "&chr(13)&_
        "GROUP BY emp_name                       "&chr(13)&_
        "  HAVING count(*) = 2                   "
  'Response.write "<pre>"& sql &"</pre><br>"
  set rs_emp = dbconn.execute(sql)
  do until rs_emp.eof 
    emp_name = rs_emp("emp_name")
    
    sql = "UPDATE emp_master_month              "&chr(13)&_    
          "   SET cost_except = '2'             "&chr(13)&_
          " WHERE emp_name    = '"&emp_name&"'  "&chr(13)&_
          "   AND emp_month   = '"&end_month&"' "&chr(13)&_
          "   AND emp_company = '케이원정보통신'"
    'Response.write "<pre>"& sql &"</pre><br>"
    dbconn.execute(sql)
     
    rs_emp.movenext()
  loop  
  rs_emp.close()
  
  ' KDC사업부에 상주직접비에 해당하는 사원을 cost_except = '2' 로 변경한다.
	sql = "UPDATE emp_master_month                               "&chr(13)&_
        "   SET cost_except = '2'                              "&chr(13)&_
        " WHERE emp_month   = '"&end_month&"'                  "&chr(13)&_
        "   AND cost_center = '상주직접비'                     "&chr(13)&_
        "   AND emp_saupbu  = 'KDC사업부'                      "&chr(13)&_
        "   AND emp_no IN ( SELECT pmg_emp_no                  "&chr(13)&_
        "                     FROM pay_month_give              "&chr(13)&_
        "                    WHERE pmg_id    = 1               "&chr(13)&_
        "                      AND pmg_yymm  = '"&end_month&"' "&chr(13)&_
        "                      AND mg_saupbu ='KDC사업부'      "&chr(13)&_
        "                 )                                    "
  'Response.write "<pre>"& sql &"</pre><br>"        
  dbconn.execute(sql)

	''
	'' KDC사업부 사원 예외처리 _ 끝
	''

	
	'공통비 배부기준 변경 처리(2016-01-15)
    ' 사업부별 인원수 (영업관리 / 손익현황 / '사업부별 인원현황' 과 일치, B.cost_except=2가 손익제외)
    sql = "SELECT count(*)                          " & chr(13) &_
          "  FROM pay_month_give  A                 " & chr(13) &_
          "     , emp_master_month B                " & chr(13) &_
          " WHERE A.pmg_id     = '1'                " & chr(13) &_
          "   AND A.pmg_yymm   = '"&end_month&"'    " & chr(13) &_
          "   AND A.mg_saupbu  = '"&rs("saupbu")&"' " & chr(13) &_
          "   AND A.pmg_emp_no =  B.emp_no          " & chr(13) &_
          "   /* AND B.cost_except in ('0','1') */  " & chr(13) &_
          "   AND B.cost_center <> '손익제외'       " & chr(13) &_
          "   AND B.emp_month  ='"&end_month&"'     "   
    'Response.write "<pre>"&sql&"</pre><br>"
    set rs_emp=dbconn.execute(sql)
    if rs_emp(0) = "" or isnull(rs_emp(0)) then
        saupbu_person = 0
    else
        saupbu_person = clng(rs_emp(0))
    end if
    rs_emp.close()

    ' 사업부별 매출액
    sql = "SELECT ifnull(sum(cost_amt),0) AS cost_amt          " & chr(13) &_
          "  FROM saupbu_sales                                 " & chr(13) &_
          "   WHERE (    sales_date >= '" + from_date + "'     " & chr(13) &_    
          "          AND sales_date <= '" + to_date + "'  )     " & chr(13) &_    
          "   AND saupbu ='"&rs("saupbu")&"'                   "
'Response.write "<pre>"&sql&"</pre><br>"
    set rs_cost=dbconn.execute(sql)
    if rs_cost(0) = "" or isnull(rs_cost(0)) then
        saupbu_sales = 0
    else
        saupbu_sales = CCur(rs_cost(0))
    end if
    rs_cost.close()


    i = i + 1
    saupbu_tab(i,1) = rs("saupbu")  ' 사업부 명
    saupbu_tab(i,2) = saupbu_person ' 해당 사업부의 인원수
    saupbu_tab(i,3) = saupbu_sales  ' 해당 사업부의 매출액

    tot_person       = tot_person       + saupbu_person ' 사업부 전체 인원
    tot_saupbu_sales = tot_saupbu_sales + saupbu_sales  ' 사업부 전체 매출액
    
    rs.movenext()
loop    
rs.close()

'전사공통비 총액
sql = "select sum(cost_amt_"&mm&") as tot_cost from company_cost where cost_year ='"&cost_year&"' and cost_center = '전사공통비'"
'Response.write sql&"<br>"
Set rs=DbConn.Execute(SQL)
tot_cost_amt = clng(rs("tot_cost")) ' (1)전사공통비 = (2) + (3) 
'Response.write "<pre>(1)전사공통비 = "&tot_cost_amt&"</pre><br>"       

tot_cost_amt_h1 = round(tot_cost_amt/2,0)        ' (2)전사공통비 총액의 50%
tot_cost_amt_h2 = tot_cost_amt - tot_cost_amt_h1 ' (3)전사공통비 총액의 50%
'Response.write "<pre>(2)전사공통비 = "&tot_cost_amt_h1&"</pre><br>"       
'Response.write "<pre>(3)전사공통비 = "&tot_cost_amt_h2&"</pre><br>"       

rs.close()

' 사업부별 손익 자료 생성
' 처리전 zero
sql = "update saupbu_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='전사공통비') "
dbconn.execute(sql)
sql = "delete from management_cost where cost_month ='"&end_month&"'"
'Response.write sql&"<br>"
dbconn.execute(sql)


' 전사공통비 배분

for i = 1 to 10 ' 각 사업부로 순환 
    if saupbu_tab(i,1) = "" or isnull(saupbu_tab(i,1)) then
        exit for
    end if

    'saupbu_cost_amt = int(tot_cost_amt * saupbu_per) ' 전체금액의 전사공통비 (As Is)

    if saupbu_tab(i,2) = 0 then
        saupbu_per       = 0 ' (8)사업부 인원 / (4)총 사업부 인원
        saupbu_cost_amt  = 0 ' (6)전사공통비(인력) = ((2)전체공통비1 50%) * ((8)사업부 인원 / (4)총 사업부 인원) (To Be)  =E47*C49/C45
        chage_rate_inwon = 0 ' (10) 인력 차지율 = (6)전사공통비(인력) / (1)전사공통비 
    else
        saupbu_per       = saupbu_tab(i,2) / tot_person        ' (8)사업부 인원 / (4)총 사업부 인원
        saupbu_cost_amt  = round(tot_cost_amt_h1 * saupbu_per) ' (6)전사공통비(인력) = ((2)전체공통비1 50%) * ((8)사업부 인원 / (4)총 사업부 인원) (To Be)  =E47*C49/C45
        chage_rate_inwon = saupbu_cost_amt / tot_cost_amt_h1   ' (10) 인력 차지율 = (6)전사공통비(인력) / (1)전사공통비 
    end if 

    if saupbu_tab(i,3) = 0 then
        saupbu_sale_per = 0 ' (9)사업부 매출 / (5)총 사업부매출 (To Be)
        saupbu_sale_amt = 0 ' (7)전사공통비(매출) = ((3)전체공통비2 50%) * ((9)사업부 매출 / (5)총 사업부매출) (To Be) =G47*C50/C46
        chage_rate_sale = 0 ' (11) 매출 차지율 = (7)전사공통비(매출) / (1)전사공통비 
    else
        saupbu_sale_per = saupbu_tab(i,3)  / tot_saupbu_sales      ' (9)사업부 매출 / (5)총 사업부매출 (To Be)
        saupbu_sale_amt = round(tot_cost_amt_h2 * saupbu_sale_per) ' (7)전사공통비(매출) = ((3)전체공통비2 50%) * ((9)사업부 매출 / (5)총 사업부매출) (To Be) =G47*C50/C46
        chage_rate_sale = saupbu_sale_amt / tot_cost_amt_h2        ' (11) 매출 차지율 = (7)전사공통비(매출) / (1)전사공통비 
    end if 

    ' 사업부 별 매출 합계를 구한다.

    ' 각 사업부의 고객사별 매출합계
    sql = "  SELECT saupbu, company, sum(cost_amt) as cost  " & chr(13) &_
          "    FROM saupbu_sales                            " & chr(13) &_
          "   WHERE (    sales_date >= '" + from_date + "'  " & chr(13) &_    
          "          AND sales_date <= '" + to_date + "'  ) " & chr(13) &_    
          "     AND saupbu ='"&saupbu_tab(i,1)&"'           " & chr(13) &_
          "GROUP BY saupbu, company                         "
'Response.write "<pre>"&sql&"</pre><br>"
    rs_etc.Open sql, Dbconn, 1
    k = 0
    do until rs_etc.eof
        k = k + 1
        if saupbu_sales = 0 then
            charge_per = 0 ' ex, 기타사업부
        else
            charge_per = rs_etc("cost") / saupbu_sales ' 고객사매출 합계 / 사업부매출 합계
        end if
        cost_amt = int(charge_per * saupbu_cost_amt) ' 배분된 전사공통비 (고객사별)
        
        sql = "INSERT INTO management_cost ( cost_month              " & chr(13) &_
              "                            , saupbu                  " & chr(13) &_
              "                            , company                 " & chr(13) &_
              "                            , tot_person              " & chr(13) &_
              "                            , saupbu_person           " & chr(13) &_
              "                            , saupbu_per              " & chr(13) &_
              "                            , tot_cost_amt            " & chr(13) &_
              "                            , saupbu_cost_amt         " & chr(13) &_
              "                            , saupbu_sale             " & chr(13) &_
              "                            , tot_sale                " & chr(13) &_
              "                            , sale_per                " & chr(13) &_
              "                            , saupbu_sale_amt         " & chr(13) &_
              "                            , charge_per              " & chr(13) &_
              "                            , cost_amt                " & chr(13) &_
              "                            , reg_id                  " & chr(13) &_
              "                            , reg_name                " & chr(13) &_
              "                            , reg_date                " & chr(13) &_
              "                            )                         " & chr(13) &_
              "                     VALUES ( '"&end_month&"'         " & chr(13) &_
              "                            , '"&saupbu_tab(i,1)&"'   " & chr(13) &_
              "                            , '"&rs_etc("company")&"' " & chr(13) &_
              "                            , "&tot_person&"          " & chr(13) &_
              "                            , "&saupbu_tab(i,2)&"     " & chr(13) &_
              "                            , "&chage_rate_inwon&"    " & chr(13) &_
              "                            , "&tot_cost_amt&"        " & chr(13) &_
              "                            , "&saupbu_cost_amt&"     " & chr(13) &_
              "                            , "&saupbu_tab(i,3)&"     " & chr(13) &_
              "                            , "&tot_saupbu_sales&"    " & chr(13) &_
              "                            , "&chage_rate_sale&"      " & chr(13) &_
              "                            , "&saupbu_sale_amt&"     " & chr(13) &_
              "                            , "&charge_per&"          " & chr(13) &_
              "                            , "&cost_amt&"            " & chr(13) &_
              "                            , '"&user_Id&"'           " & chr(13) &_
              "                            , '"&user_name&"'         " & chr(13) &_
              "                            , now()                   " & chr(13) &_
              "                            )                         "
'Response.write "<pre>"&sql&"</pre><br>"
        dbconn.execute(sql)

        rs_etc.movenext()
    loop
' 매출이 제로인 경우
    if k = 0 then
        sql = "INSERT INTO management_cost ( cost_month              " & chr(13) &_
              "                            , saupbu                  " & chr(13) &_
              "                            , company                 " & chr(13) &_
              "                            , tot_person              " & chr(13) &_
              "                            , saupbu_person           " & chr(13) &_
              "                            , saupbu_per              " & chr(13) &_
              "                            , tot_cost_amt            " & chr(13) &_
              "                            , saupbu_cost_amt         " & chr(13) &_
              "                            , saupbu_sale             " & chr(13) &_
              "                            , tot_sale                " & chr(13) &_
              "                            , sale_per                " & chr(13) &_
              "                            , saupbu_sale_amt         " & chr(13) &_
              "                            , charge_per              " & chr(13) &_
              "                            , cost_amt                " & chr(13) &_
              "                            , reg_id                  " & chr(13) &_
              "                            , reg_name                " & chr(13) &_
              "                            , reg_date                " & chr(13) &_
              "                            )                         " & chr(13) &_
              "                     VALUES ( '"&end_month&"'         " & chr(13) &_
              "                            , '"&saupbu_tab(i,1)&"'   " & chr(13) &_
              "                            , ''                      " & chr(13) &_
              "                            , "&tot_person&"          " & chr(13) &_
              "                            , "&saupbu_tab(i,2)&"     " & chr(13) &_
              "                            , "&chage_rate_inwon&"    " & chr(13) &_
              "                            , "&tot_cost_amt&"        " & chr(13) &_
              "                            , "&saupbu_cost_amt&"     " & chr(13) &_
              "                            , "&saupbu_tab(i,3)&"     " & chr(13) &_
              "                            , "&tot_saupbu_sales&"    " & chr(13) &_
              "                            , "&chage_rate_sale&"     " & chr(13) &_
              "                            , "&saupbu_sale_amt&"     " & chr(13) &_
              "                            , "&charge_per&"          " & chr(13) &_
              "                            , "&cost_amt&"            " & chr(13) &_
              "                            , '"&user_Id&"'           " & chr(13) &_
              "                            , '"&user_name&"'         " & chr(13) &_
              "                            , now()                   " & chr(13) &_
              "                            )                         " 
'Response.write "<pre>"&sql&"</pre><br>"
        dbconn.execute(sql)
    end if
    rs_etc.close()

    sql = "   SELECT cost_id, cost_detail                  " & chr(13) &_
          "        , sum(cost_amt_"&cost_month&") as cost  " & chr(13) &_
          "     FROM company_cost                          " & chr(13) &_
          "    WHERE (cost_center = '전사공통비' )         " & chr(13) &_
          "      AND cost_year ='"&cost_year&"'            " & chr(13) &_
          " GROUP BY cost_id, cost_detail                  "
'Response.write "<pre>"&sql&"</pre><br>"
    rs_etc.Open sql, Dbconn, 1
    do until rs_etc.eof

        'cost = int(saupbu_per * clng(rs_etc("cost"))) ' (As is)
        if clng(rs_etc("cost")) = 0 then
            cost_h1 = 0 ' 인원배분 50%
            cost_h2 = 0 ' 매출액배분 50%
        else
            cost_h1 = round(clng(rs_etc("cost")) / 2) ' 인원배분 50%
            cost_h2 = clng(rs_etc("cost")) - cost_h1  ' 매출액배분 50%
        end if

        cost1 = int(saupbu_per      * cost_h1)
        cost2 = int(saupbu_sale_per * cost_h2)
        cost  = cost1 + cost2 ' (To Be) 인원배분 + 매출액배분

        sql = "SELECT * FROM saupbu_profit_loss                " & chr(13) &_
              " WHERE cost_year   ='"&cost_year&"'             " & chr(13) &_
              "   AND saupbu      ='"&saupbu_tab(i,1)&"'       " & chr(13) &_
              "   AND cost_center ='전사공통비'                " & chr(13) &_
              "   AND cost_id     ='"&rs_etc("cost_id")&"'     " & chr(13) &_
              "   AND cost_detail ='"&rs_etc("cost_detail")&"' "
        set rs_cost=dbconn.execute(sql)
            
        if rs_cost.eof or rs_cost.bof then
            sql = "INSERT INTO saupbu_profit_loss (cost_year,saupbu,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&")                     " & chr(13) &_
                  "     VALUES ('"&cost_year&"','"&saupbu_tab(i,1)&"','전사공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&") "
            dbconn.execute(sql)
        else
            sql = "UPDATE saupbu_profit_loss                       " & chr(13) &_
                  "   SET cost_amt_"&cost_month&"="&cost&"         " & chr(13) &_
                  " WHERE cost_year   ='"&cost_year&"'             " & chr(13) &_
                  "   AND saupbu      ='"&saupbu_tab(i,1)&"'       " & chr(13) &_
                  "   AND cost_center ='전사공통비'                " & chr(13) &_
                  "   AND cost_id     ='"&rs_etc("cost_id")&"'     " & chr(13) &_
                  "   AND cost_detail ='"&rs_etc("cost_detail")&"' "
            dbconn.execute(sql)
        end if      

        rs_etc.movenext()
    loop
    rs_etc.close()
next
' 전사공통비 배부 끝

' 고객사별 손익 자료 생성
' 전사공통비 배부
' 처리전 zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='전사공통비') "
dbconn.execute(sql)

sql = " select company,saupbu_per, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY company"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	charge_per = rs("charge_per")

	sql = "select * from trade where trade_name = '"&rs("company")&"'"
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		group_name = "Error"
	  else
		group_name = rs_trade("group_name")
	end if		  

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '전사공통비' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * clng(rs_etc("cost")) * rs("saupbu_per"))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','전사공통비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='전사공통비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()

' 고객사별 직접비 배부
' 처리전 zero
sql = "update company_profit_loss set cost_amt_"&cost_month&"= '0' where cost_year ='"&cost_year&"' and (cost_center ='직접비') "
dbconn.execute(sql)

sql = " select saupbu,company, sum(charge_per) as charge_per from management_cost Where (cost_month = '"&end_month&"') GROUP BY saupbu,company"
Rs.Open Sql, Dbconn, 1
do until rs.eof
	charge_per = rs("charge_per")

	sql = "select * from trade where trade_name = '"&rs("company")&"'"
	set rs_trade=dbconn.execute(sql)
	if rs_trade.eof or rs_trade.bof then
		group_name = "Error"
	  else
		group_name = rs_trade("group_name")
	end if		  

	sql = "select cost_id,cost_detail,sum(cost_amt_"&cost_month&") as cost from company_cost where (cost_center = '직접비' ) and (saupbu = '"&rs("saupbu")&"' ) and cost_year ='"&cost_year&"' group by cost_id,cost_detail"
	rs_etc.Open sql, Dbconn, 1
	do until rs_etc.eof
		
		cost = int(charge_per * Cdbl(rs_etc("cost")))

		sql = "select * from company_profit_loss where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='직접비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
		set rs_cost=dbconn.execute(sql)
		
		if rs_cost.eof or rs_cost.bof then
			sql = "insert into company_profit_loss (cost_year,company,group_name,cost_center,cost_id,cost_detail,cost_amt_"&cost_month&") values ('"&cost_year&"','"&rs("company")&"','"&group_name&"','직접비','"&rs_etc("cost_id")&"','"&rs_etc("cost_detail")&"',"&cost&")"
			dbconn.execute(sql)
		  else
			sql = "update company_profit_loss set cost_amt_"&cost_month&"="&cost&" where cost_year ='"&cost_year&"' and company ='"&rs("company")&"' and group_name ='"&group_name&"' and cost_center ='직접비' and cost_id ='"&rs_etc("cost_id")&"' and cost_detail ='"&rs_etc("cost_detail")&"'"
			dbconn.execute(sql)
		end if		
		
		rs_etc.movenext()
	loop
	rs_etc.close()
	rs.movenext()
loop
rs.close()
' 고객사별 직접비 배부 끝

if end_yn = "C" then
	sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
	"' and saupbu = '공통비/직접비배분'"
  else
	sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
	"','공통비/직접비배분','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
end if
dbconn.execute(sql)

if Err.number <> 0 then
	dbconn.RollbackTrans 
	end_msg = emp_msg + "처리중 Error가 발생하였습니다...."
else    
	dbconn.CommitTrans
	end_msg = emp_msg + "마감처리 되었습니다...."
end if

Response.write "<script language=javascript>"
Response.write "alert('"&end_msg&"');"
Response.write "location.replace('cost_end_mg.asp');"
Response.write "</script>"
Response.End

dbconn.Close()
Set dbconn = Nothing

end if
%>