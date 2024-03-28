<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
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
Dim slip_month, card_type, owner_company
Dim field_check, field_view
Dim from_date, end_date, to_date
Dim end_msg, url

'on Error resume next

slip_month = Request.Form("slip_month")	'검색 년월
card_type = Request.Form("card_type")	'카드 유형
owner_company = Request.Form("owner_company")	'소유 회사

field_check = Request.Form("field_check")
field_view = Request.Form("field_view")

from_date = Mid(slip_month, 1, 4) + "-" + Mid(slip_month, 5, 2) + "-01"
end_date = DateValue (from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

DBConn.BeginTrans

objBuilder.Append "DELETE FROM card_slip "
objBuilder.Append "WHERE (slip_date >= '"&from_date&"' AND slip_date <= '"&to_date&"') "

'소유 회사 검색 조건
If Trim(owner_company) <> "전체" Then
	objBuilder.Append "AND owner_company = '"&owner_company&"' "
End If

'카드 유형 검색 조건
If Trim(card_type) <> "전체" Then
	objBuilder.Append "AND card_type ='"&card_type&"' "
End If

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "삭제중 Error가 발생하였습니다...."
Else
	DBConn.CommitTrans
	end_msg = "삭제 처리 되었습니다...."
End If

url = "card_slip_mg.asp?slip_month="&slip_month&"&card_type="&card_type&"&field_check="&field_check&"&field_view="&field_view&"&ck_sw="&"y"

Response.Write"<script language=javascript>"
Response.Write"alert('"&end_msg&"');"
Response.Write"location.replace('"&url&"');"
Response.Write"</script>"
Response.End

DBConn.Close()
Set DBConn = Nothing
%>


