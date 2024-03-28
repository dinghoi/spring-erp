<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'On Error Resume Next
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
Dim result : result = "fail"
Dim grade

user_id = Request("user_id")
grade = Request("grade")

DBConn.BeginTrans

'sql = " UPDATE memb                         "&chr(13)&_
'	  "    SET grade = '" & grade & "'      "&chr(13)&_
'	  "  WHERE user_id = '" & user_id & "'  "
objBuilder.Append "UPDATE memb SET "
objBuilder.Append "	grade = '"&grade&"' "
objBuilder.Append "WHERE user_id = '"&user_id&"';"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	result = "error"
Else
	DBConn.CommitTrans
	result = "succ"
End If

Dbconn.close : Set Dbconn = Nothing

If Trim(result) <> "" Then
	result = "{""result"" : """ & result & """}"
End If

Response.write result
%>