<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
On Error Resume Next

'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim bill_id, bill_month
Dim from_date, end_date, to_date
Dim sql, end_msg, url
Dim cost_reg_yn, end_yn

bill_id = Request.Form("bill_id")
bill_month = Request.Form("bill_month")
cost_reg_yn = Request.Form("cost_reg_yn")
end_yn = Request.Form("end_yn")

from_date = Mid(bill_month, 1, 4) & "-" & Mid(bill_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

DBConn.BeginTrans

sql = "DELETE FROM tax_bill WHERE (bill_date >= '"&from_date&"' AND bill_date <= '"&to_date&"') AND bill_id ='"&bill_id&"' "
DBConn.Execute(sql)

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "업로드 삭제 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "업로드 삭제 처리 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

url = "/finance/tax_esero_mg.asp?bill_month="&bill_month&"&bill_id="&bill_id&"&cost_reg_yn="&cost_reg_yn&"&end_yn="&end_yn

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('"&url&"');"
Response.Write "</script>"
Response.End
%>


