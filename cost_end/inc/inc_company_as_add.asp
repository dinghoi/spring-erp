<%
' 사업부별 인원수 집계
'sql = " select saupbu from sales_org where sales_year='" & cost_year & "' order by saupbu asc"
objBuilder.Append "SELECT saupbu FROM sales_org WHERE sales_year='" & cost_year & "' ORDER BY saupbu ASC "

'rs
Set rsSalesExcept = Server.CreateObject("ADODB.RecordSet")
rsSalesExcept.Open objBuilder.ToString(), DBConn, 1

i = 0
tot_person = 0

Do Until rsSalesExcept.EOF

	' KDC사업부 사원 예외처리

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
		Response.write sql
		dbconn.rollbacktrans
		Response.end
  'Response.write "<pre>"& sql &"</pre><br>"
'  set rs_emp = dbconn.execute(sql)
'  do until rs_emp.eof
 '   emp_name = rs_emp("emp_name")
 '
 '   sql = "UPDATE emp_master_month              "&chr(13)&_
 '         "   SET cost_except = '2'             "&chr(13)&_
 '         " WHERE emp_name    = '"&emp_name&"'  "&chr(13)&_
 '         "   AND emp_month   = '"&end_month&"' "&chr(13)&_
 '         "   AND emp_company = '케이원정보통신'"
 '   'Response.write "<pre>"& sql &"</pre><br>"
 '   dbconn.execute(sql)
 '
 '   rs_emp.movenext()
 ' loop
 ' rs_emp.close()

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
  'dbconn.execute(sql)

	''
	'' KDC사업부 사원 예외처리 _ 끝
	''


	'공통비 배부기준 변경 처리(2016-01-15)
	sql = "select count(*) from pay_month_give  A ,emp_master_month B "
	sql = sql & "where A.pmg_id = '1'  "
	sql = sql & "and A.pmg_yymm = '"&end_month&"' "
	sql = sql & "and A.mg_saupbu ='"&rs("saupbu")&"' "
	sql = sql & "and A.pmg_emp_no=  B.emp_no "
	sql = sql & "and B.cost_except in('0','1') "
	sql = sql & "and B.emp_month ='"&end_month&"' "

	'Response.write sql&"<br>"

	set rs_emp=dbconn.execute(sql)
	if rs_emp(0) = "" or isnull(rs_emp(0)) then
		saupbu_person = 0
	  else
		saupbu_person = clng(rs_emp(0))
	end if
	rs_emp.close()
	i = i + 1
	saupbu_tab(i,1) = rs("saupbu")
	saupbu_tab(i,2) = saupbu_person
	tot_person = tot_person + saupbu_person

	rs.movenext()
loop
rs.close()
%>