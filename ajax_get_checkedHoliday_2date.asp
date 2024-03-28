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
Dim orgColumn, sqlWhere

  work_date1 = request("work_date1")
  work_date2 = request("work_date2")

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")	
	
	holiday_memo1 = ""
	holiday_memo2 = ""
	
	dbconn.open DbConnect

	sql = " SELECT holiday_memo                    "&chr(13)&_
	      "   FROM holiday                         "&chr(13)&_
	      "  WHERE holiday = '" & work_date1 & "'  "
  'Response.write sql&chr(13)
	rs.Open sql, Dbconn, 1
	
  if not (rs.eof or rs.bof) then
  	holiday_memo1 = rs("holiday_memo")  	
  end if
  rs.close
  'Response.write holiday_memo1&chr(13)

	sql = " SELECT holiday_memo                    "&chr(13)&_
	      "   FROM holiday                         "&chr(13)&_
	      "  WHERE holiday = '" & work_date2 & "'  "
  'Response.write sql&chr(13)
	rs.Open sql, Dbconn, 1
	
  if not (rs.eof or rs.bof) then
  	holiday_memo2 = rs("holiday_memo")  	
  end if
  rs.close
  'Response.write holiday_memo2&chr(13)
  
	result = "succ"

	Dbconn.close : Set Dbconn = Nothing

If Err.number<>0 Then
	result = "error"
End IF


If Trim(result&"")<>"" Then
	result = "{""result"" : """ & result & """, ""holiday_memo1"" : """ & holiday_memo1 & """, ""holiday_memo2"" : """ & holiday_memo2 & """}"
End If

Response.write result
%>