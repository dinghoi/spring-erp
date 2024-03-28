<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	bill_id=Request.form("bill_id")
	bill_month=Request.form("bill_month")

	from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans

	sql = "Update tax_bill set end_yn='N' where bill_date >= '"&from_date&"' and bill_date <= '"&to_date&"' and bill_id = '"&bill_id&"'"
	dbconn.execute(sql)
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "처리중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "마감처리 취소 되었습니다...."
	end if

	url = "tax_esero_mg.asp?bill_month="&bill_month&"&bill_id="&bill_id&"&cost_reg_yn="&"N"&"&end_yn="&"Y"&"&ck_sw="&"y"
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


