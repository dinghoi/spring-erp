<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	company = request("company")
	seq = request("seq")

	response.write("����ó����!!!!")

	set dbconn = server.CreateObject("adodb.connection")
	dbconn.open dbconnect

	dbconn.BeginTrans
	
	sql = "select * from company_form where company = '"&company&"'"
	set rs = dbconn.execute(sql)
	if rs.eof or rs.bof then
		response.write("���� �����Ͱ� �����ϴ�")
	  else
		if seq = 1 then
			forms = rs("form1")
			sql = "Update company_form set form1 ='', up_date1 = now(), up_id1= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 2 then
			forms = rs("form2")
			sql = "Update company_form set form2 ='', up_date2 = now(), up_id2= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 3 then
			forms = rs("form3")
			sql = "Update company_form set form3 ='', up_date3 = now(), up_id3= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 4 then
			forms = rs("form4")
			sql = "Update company_form set form4 ='', up_date4 = now(), up_id4= '"&user_id&"' where company = '"&company&"'"
		end if
		if seq = 5 then
			forms = rs("form5")
			sql = "Update company_form set form5 ='', up_date5 = now(), up_id5= '"&user_id&"' where company = '"&company&"'"
		end if
		dbconn.execute(sql)
	end if
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "������ Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = "�����Ǿ����ϴ�...."
	end if
	
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"self.opener.location.reload();"		
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing

%>
