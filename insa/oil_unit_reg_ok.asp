<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### 작업 내역
'===================================================
' 허정호_20210722 :
'	- 신규 페이지 작성 및 코드 정리

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
Dim u_type, oil_unit_month
Dim oil_unit_middle11, oil_unit_last11, oil_unit_average11
Dim oil_unit_middle12, oil_unit_last12, oil_unit_average12
Dim oil_unit_middle13, oil_unit_last13, oil_unit_average13
Dim oil_unit_middle21, oil_unit_last21, oil_unit_average21
Dim oil_unit_middle22, oil_unit_last22, oil_unit_average22
Dim oil_unit_middle23, oil_unit_last23, oil_unit_average23
Dim end_msg, strSql


u_type = Request.Form("u_type")
oil_unit_month = Request.Form("oil_unit_month")

oil_unit_middle11 = Int(Request.Form("oil_unit_middle11"))
oil_unit_last11 = Int(Request.Form("oil_unit_last11"))

If oil_unit_last11 = 0 Then
	oil_unit_average11 = oil_unit_middle11
Else
	oil_unit_average11 = (oil_unit_middle11 + oil_unit_last11) / 2
End If

oil_unit_middle12 = Int(Request.Form("oil_unit_middle12"))
oil_unit_last12 = Int(Request.Form("oil_unit_last12"))

If oil_unit_last12 = 0 Then
	oil_unit_average12 = oil_unit_middle12
Else
	oil_unit_average12 = (oil_unit_middle12 + oil_unit_last12) / 2
End If

oil_unit_middle13 = Int(Request.Form("oil_unit_middle13"))
oil_unit_last13 = Int(Request.Form("oil_unit_last13"))

If oil_unit_last13 = 0 Then
	oil_unit_average13 = oil_unit_middle13
Else
	oil_unit_average13 = (oil_unit_middle13 + oil_unit_last13) / 2
End If

oil_unit_middle21 = Int(Request.Form("oil_unit_middle21"))
oil_unit_last21 = Int(Request.Form("oil_unit_last21"))

If oil_unit_last21 = 0 Then
	oil_unit_average21 = oil_unit_middle21
Else
	oil_unit_average21 = (oil_unit_middle21 + oil_unit_last21) / 2
End If

oil_unit_middle22 = Int(Request.Form("oil_unit_middle22"))
oil_unit_last22 = Int(Request.Form("oil_unit_last22"))

If oil_unit_last22 = 0 Then
	oil_unit_average22 = oil_unit_middle22
Else
	oil_unit_average22 = (oil_unit_middle22 + oil_unit_last22) / 2
End If

oil_unit_middle23 = Int(Request.Form("oil_unit_middle23"))
oil_unit_last23 = Int(Request.Form("oil_unit_last23"))

If oil_unit_last23 = 0 Then
	oil_unit_average23 = oil_unit_middle23
Else
	oil_unit_average23 = (oil_unit_middle23 + oil_unit_last23) / 2
End If

DBConn.BeginTrans

If u_type = "U" Then
	'sql = "delete from oil_unit where oil_unit_month ='"&oil_unit_month&"'"
	objBuilder.Append "DELETE FROM oil_unit WHERE oil_unit_month ='"&oil_unit_month&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If

objBuilder.Append "INSERT INTO oil_unit("
objBuilder.Append "oil_unit_month, oil_unit_id, oil_kind, oil_unit_middle, oil_unit_last,"
objBuilder.Append "oil_unit_average, reg_id, reg_user, reg_date"
objBuilder.Append ")VALUES("

strSql = objBuilder.ToString()
objBuilder.Clear()

'수도권 휘발유
objBuilder.Append strSql
objBuilder.Append "'"&oil_unit_month&"','1','휘발유',"&oil_unit_middle11&","&oil_unit_last11&","
objBuilder.Append ""&oil_unit_average11&",'"&user_id&"','"&user_name&"', NOW());"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'수도권 디젤
objBuilder.Append strSql
objBuilder.Append "'"&oil_unit_month&"','1','디젤',"&oil_unit_middle12&","&oil_unit_last12&","
objBuilder.Append ""&oil_unit_average12&",'"&user_id&"','"&user_name&"',NOW());"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'수도권 가스
objBuilder.Append strSql
objBuilder.Append "'"&oil_unit_month&"','1','가스',"&oil_unit_middle13&","&oil_unit_last13&","
objBuilder.Append ""&oil_unit_average13&",'"&user_id&"','"&user_name&"',NOW());"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'수도권 외 휘발유
objBuilder.Append strSql
objBuilder.Append "'"&oil_unit_month&"','2','휘발유',"&oil_unit_middle21&","&oil_unit_last21&","
objBuilder.Append ""&oil_unit_average21&",'"&user_id&"','"&user_name&"', NOW());"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'수도권 외 디젤
objBuilder.Append strSql
objBuilder.Append "'"&oil_unit_month&"','2','디젤',"&oil_unit_middle22&","&oil_unit_last22&","
objBuilder.Append ""&oil_unit_average22&",'"&user_id&"','"&user_name&"',NOW());"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'수도권 외 가스
objBuilder.Append strSql
objBuilder.Append "'"&oil_unit_month&"','2','가스',"&oil_unit_middle23&","&oil_unit_last23&","
objBuilder.Append ""&oil_unit_average23&",'"&user_id&"','"&user_name&"',NOW());"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "처리 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "처리 되었습니다."
End If

Response.Write "<script language=javascript>"
Response.Write "	alert('처리 완료 되었습니다.');"
'response.Redirect "/insa/oil_unit_mg.asp"
Response.Write "	location.href='/insa/oil_unit_mg.asp';"
response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>