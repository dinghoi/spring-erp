<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim pass, board_seq, page, condi, condi_value, ck_sw
Dim rsBoard, url

pass = Request.Form("pass")
board_seq = Request.Form("board_seq")
page = Request("page")
condi = Request("condi")
condi_value = Request("condi_value")
ck_sw = Request("ck_sw")

'Sql="select * from board where board_seq="&board_seq
objBuilder.Append "SELECT pass FROM board WHERE board_seq = " & board_seq

Set rsBoard = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rsBoard("pass") <> pass Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('입력하신 비밀번호가 틀립니다.');"
	Response.Write "	history.go(-1);"
	Response.Write "</script>"
Else
	'sql="delete from board where board_seq=" & board_seq
	objBuilder.Append "DELETE FROM board WHERE board_seq = " & board_seq

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	url = "/main/nkp_main.asp?page="&page&"&condi="&condi&"&condi_value="&condi_value&"&ck_sw=y"

	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('삭제 되었습니다.');"
	Response.Write "	location.replace('"&url&"');"
	Response.Write "</script>"
End If

rsBoard.Close() : Set rsBoard = Nothing
DBConn.Close() : Set DBConn = Nothing

Response.End
%>
