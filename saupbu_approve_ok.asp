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

	sql="select * from emp_org_mst where org_company = '���̿��������' and org_level='�����' and org_saupbu ='"&saupbu&"'"
	set rs_org=dbconn.execute(sql)
	if rs_org.eof or rs_org.bof then
		org_bonbu = "error"
	  else
		org_bonbu = rs_org("org_bonbu")
	end if

	if position = "�������" then
		if org_bonbu = "���һ����" then
			if saupbu = "KAL���������" or saupbu = "�������������" then
				sql = "Update cost_end set batch_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
			  else
				sql = "Update cost_end set batch_yn ='Y',bonbu_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
			end if
		  else
			sql = "Update cost_end set batch_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
		end if
	  else
		if saupbu = "�������������" then			
			sql = "Update cost_end set batch_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
		  else
			sql = "Update cost_end set batch_yn ='Y',bonbu_yn ='Y',mod_date=now() where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
		end if
	end if
	
	dbconn.execute(sql)

	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "����� Error�� �߻��Ͽ����ϴ�...."
	else    
		dbconn.CommitTrans
		end_msg = "��ϵǾ����ϴ�...."
	end if

	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
'	response.write"parent.opener.location.reload();"
	response.write"opener.document.frm.submit();"
	response.write"window.close();"		
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
	
%>
