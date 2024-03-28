<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
Response.expires=-1
Response.ContentType = "application/json"
Response.Charset = "euc-kr"

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
Dim result : result = "fail"
Dim m_seq

m_seq = replaceXSS(Unescape(toString(request("m_seq"),"")))

DBConn.BeginTrans

If f_toString(m_seq, "") = "" Then
	result = "invalid"
Else
	objBuilder.Append "UPDATE member_info Set "
	objBuilder.Append "	m_del_yn = 'Y' "
	objBuilder.Append "WHERE m_seq = '"&m_seq&"'"

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	result = "succ"
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	result = "error"
End IF

'//return json data
If f_toString(result, "") <> "" Then
	DBConn.CommitTrans
	result = "{""result"":""" & result & """}"
End If
DBConn.Close() : Set DBConn = Nothing

Response.Write result
%>