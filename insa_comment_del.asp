<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
position = request.cookies("nkpmg_user")("coo_position")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

cmt_empno=Request.form("cmt_empno")
cmt_date=Request.form("cmt_date")
cmt_empname=Request.form("cmt_empname")
owner_view=Request.form("owner_view")

'response.write(cmt_empno)
'response.write(cmt_date)
'response.write(cmt_empname)
'response.End

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_fam = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect
   
    sql = " delete from emp_comment " & _
	            "  where cmt_empno ='"&cmt_empno&"' and cmt_date = '"&cmt_date&"'"    
  
	dbconn.execute(sql)
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		' dbconn.CommitTrans
		end_msg = "삭제되었습니다...."
	end if
    'url = "insa_family_mg.asp?ck_sw="y"&view_condi=" + family_empno + "&ck_sw= y&view_condi="+view_condi+"&condi="+ condi
	if owner_view = "C" then 
	           url = "insa_comment_list.asp?ck_sw=y&view_condi=" + cmt_empname + "&owner_view="+ owner_view
		else 
		       url = "insa_comment_list.asp?ck_sw=y&view_condi=" + cmt_empno + "&owner_view="+ owner_view
	end if
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
