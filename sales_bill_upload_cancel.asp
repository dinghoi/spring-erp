<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'	on Error resume next

	sales_month=Request("sales_month")
	from_date = mid(sales_month,1,4) + "-" + mid(sales_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect
	
	dbconn.BeginTrans

	sql = "delete from saupbu_sales where sales_date >= '"&from_date&"' and sales_date <= '"&to_date&"'"
	dbconn.execute(sql)
	
	if Err.number <> 0 then
		dbconn.RollbackTrans 
		end_msg = "삭제중 Error가 발생하였습니다...."
	else    
		dbconn.CommitTrans
		end_msg = "삭제 처리 되었습니다...."
	end if

	sales_saupbu = "전체"
	field_check = "total"
	field_view = ""
	ck_sw = "y"
	url = "sales_bill_mg.asp?sales_month="&sales_month&"&sales_saupbu="&sales_saupbu&"&field_check="&field_check&"&field_view="&field_view&"&ck_sw="&ck_sw
	response.write"<script language=javascript>"
	response.write"alert('"&end_msg&"');"
	response.write"location.replace('"&url&"');"
	response.write"</script>"
	Response.End

	dbconn.Close()
	Set dbconn = Nothing
%>


