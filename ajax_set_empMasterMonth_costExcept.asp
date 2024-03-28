<%@LANGUAGE="VBSCRIPT"%>
<%
Response.expires=-1
Response.ContentType = "application/json"
Response.Charset = "euc-kr"
%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
On Error Resume Next

Dim result : result = "fail"
Dim sql

    emp_month = request("emp_month")
    emp_no	  = request("emp_no")
    chked  	  = request("chked")
   
	Set Dbconn = Server.CreateObject("ADODB.Connection")
	Set Rs     = Server.CreateObject("ADODB.Recordset")
	
    ' 손익제외인 경우 '2' 아니면 '0' 
    if chked then cost_except = "2" else  cost_except = "0" end if

	dbconn.open DbConnect

	sql = "UPDATE emp_master_month                     " & chr(13) &_
	      "   SET cost_except = '" & cost_except & "'  " & chr(13) &_
	      " WHERE emp_month = '" & emp_month & "'      " & chr(13) &_ 
	      "   AND emp_no    = '" & emp_no & "'         "
'Response.write sql
	Dbconn.execute sql


	sql = "UPDATE emp_master                           " & chr(13) &_
	      "   SET cost_except = '" & cost_except & "'  " & chr(13) &_
	      " WHERE emp_no = '" & emp_no & "'            "
'Response.write sql
	Dbconn.execute sql

	result = "succ"

	Dbconn.close : Set Dbconn = Nothing

If Err.number<>0 Then
	result = "error"
End IF


If Trim(result&"")<>"" Then
	result = "{""result"":""" & result & """}"
End If

Response.write result
%>