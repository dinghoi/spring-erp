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
Dim u_type, curr_date, car_old_no, del_date, end_msg

u_type = f_Request("u_type")
car_old_no = f_Request("car_old_no")

curr_date = Mid(CStr(Now()), 1, 10)
del_date = CStr(Mid(curr_date, 1, 4)) & CStr(Mid(curr_date, 6, 2)) & CStr(Mid(curr_date, 9, 2))

DBConn.BeginTrans

objBuilder.Append "INSERT INTO car_info_del "
objBuilder.Append "SELECT '"&del_date&"' AS 'car_del_date', car_info.* "
objBuilder.Append "FROM car_info WHERE car_no ='"&car_old_no&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "DELETE FROM car_info WHERE car_no ='"&car_old_no&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = "삭제 중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = "정상적으로 삭제 되었습니다."
End If

DBConn.Close() : Set DBConn = Nothing

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	self.opener.location.reload();"
Response.Write "	window.close();"
Response.Write "</script>"
Response.End
%>
