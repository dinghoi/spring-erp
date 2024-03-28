<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

	u_type = request.form("u_type")
	cost_month = request.form("cost_month")
	saupbu = request.form("saupbu")
	saupbu_memo = request.form("saupbu_memo")
	saupbu_memo = Replace(saupbu_memo,"'","&quot;")
	bonbu_memo = request.form("bonbu_memo")
	bonbu_memo = Replace(bonbu_memo,"'","&quot;")
	memo_id = request.form("memo_id")
	
	set dbconn = server.CreateObject("adodb.connection")
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dbconn.open dbconnect

	dbconn.BeginTrans

	sql="select * from emp_org_mst where org_company = '케이원정보통신' and org_level='사업부' and org_saupbu ='"&saupbu&"'"
	set rs_org=dbconn.execute(sql)
	if rs_org.eof or rs_org.bof then
		org_bonbu = "error"
	  else
		org_bonbu = rs_org("org_bonbu")
	end if

	sql="select * from saupbu_memo where cost_month='"&cost_month&"' and saupbu='"&saupbu&"'"
	set rs=dbconn.execute(sql)
		
	if rs.eof or rs.bof then
		sql = "insert into saupbu_memo (cost_month,saupbu,saupbu_memo,end_yn,saupbu_reg_name,saupbu_reg_date) values "& _
		"('"&cost_month&"','"&saupbu&"','"&saupbu_memo&"','N','"&user_name&"',now())"
		dbconn.execute(sql)
	  else
		if memo_id = "1" then
			sql = "Update saupbu_memo set saupbu_memo ='"&saupbu_memo&"',saupbu_reg_name='"&user_name&"',saupbu_reg_date=now()"& _
			" where cost_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
			dbconn.execute(sql)
		  else
			sql = "Update saupbu_memo set bonbu_memo ='"&bonbu_memo&"',bonbu_reg_name='"&user_name&"',bonbu_reg_date=now()"& _
			" where cost_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
			dbconn.execute(sql)
		end if
	end if

'	if org_bonbu = "직할사업부" then
'		if saupbu = "KAL지원사업부" or saupbu = "공항지원사업부" then
'			sql = "Update cost_end set batch_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
'		  else
'			sql = "Update cost_end set batch_yn ='Y',bonbu_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
'	  	end if
'	  else
'		sql = "Update cost_end set batch_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
'	end if
'	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "등록중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "등록되었습니다...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"parent.opener.location.reload();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
