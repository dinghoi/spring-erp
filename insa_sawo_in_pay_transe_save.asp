<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

pmg_yymm=Request.form("pmg_yymm1")
view_condi=Request.form("view_condi1")

'response.write(pmg_yymm)
'response.write(view_condi)
'response.End
w_cnt = 0
m_cnt = 0

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_this = Server.CreateObject("ADODB.Recordset")
Set Rs_mem = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


if view_condi = "전체" then
         Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0) ORDER BY de_company,de_emp_no ASC"
   else
         Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_sawo_amt > 0) and (de_company = '"+view_condi+"') ORDER BY de_company,de_emp_no ASC"
end if
Rs.Open Sql, Dbconn, 1

in_empno = rs("de_emp_no")
in_date = rs("de_date")
in_seq = "001"

'in_date = mid(pmg_yymm,1,4) + "-" + mid(pmg_yymm,5,2) + "-" + "30" 

Sql = "select * from emp_sawo_in where (in_date = '"&in_date&"') and (in_company = '"&view_condi&"')"
'Sql = "SELECT * FROM emp_sawo_in WHERE (in_date = '"+in_date+"' ) and (in_seq = '"+in_seq+"') and (in_empno = '"+in_empno+"') and (in_company = '"+view_condi+"')"
Set Rs_this=Dbconn.Execute(sql)
if Rs_this.eof then
   do until rs.eof

	    in_empno = rs("de_emp_no")
	    emp_no = rs("de_emp_no")
        in_company = rs("de_company")
	    in_emp_name = rs("de_emp_name")
	    in_org = rs("de_org_code")
	    in_org_name = rs("de_org_name")
	    in_pay = int(rs("de_sawo_amt"))
	    in_comment = ""
   
	    sql="insert into emp_sawo_in (in_empno,in_seq,in_date,in_emp_name,in_company,in_org,in_org_name,in_pay,in_comment,in_reg_date,in_reg_user) values ('"&in_empno&"','"&in_seq&"','"&in_date&"','"&in_emp_name&"','"&in_company&"','"&in_org&"','"&in_org_name&"','"&in_pay&"','"&in_comment&"',now(),'"&emp_user&"')"
		
		dbconn.execute(sql)
		
		w_cnt = w_cnt + 1
		
        sql="select * from emp_sawo_mem where sawo_empno='"&in_empno&"'"
        set rs_mem=dbconn.execute(sql)

        if rs_mem.eof then
               m_cnt = m_cnt + 1

			   sql = "insert into emp_sawo_mem(sawo_empno,sawo_date,sawo_id,sawo_emp_name,sawo_company,sawo_orgcode,sawo_org_name,sawo_target,sawo_in_count,sawo_in_pay,sawo_give_count,sawo_give_pay) values "
		       sql = sql +	" ('"&in_empno&"','"&in_date&"','입사','"&in_emp_name&"','"&in_company&"','"&in_org&"','"&in_org_name&"','Y',1,'"&in_pay&"',0,0)"
			   dbconn.execute(sql)	  
	       else
		       sawo_in_count = rs_mem("sawo_in_count") + 1
			   sawo_in_pay = rs_mem("sawo_in_pay") + in_pay
			   rs_mem.close()
			   
	           sql = "update emp_sawo_mem set sawo_in_count='"&sawo_in_count&"',sawo_in_pay='"&sawo_in_pay&"',sawo_mod_user='"&emp_user&"',sawo_mod_date=now() where sawo_empno='"&in_empno&"'"

		       dbconn.execute(sql)	  
        end if
		
		Rs.MoveNext()
    loop
	    end_msg = cstr(w_cnt) + "/" + cstr(m_cnt) + " 건 급여에서 경조금 공제금을 경조회 회비 데이터가 만들어 졌습니다..."		
		response.write"<script language=javascript>"
		response.write"alert('"&end_msg&"');"		
		response.write"location.replace('insa_sawo_in_pay_trans.asp');"
		response.write"</script>"
		Response.End
   else
		response.write"<script language=javascript>"
		response.write"alert('이미 처리된 내역이 있습니다...');"		
		response.write"location.replace('insa_sawo_in_pay_trans.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
