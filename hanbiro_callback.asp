<%@Language="VBScript" CODEPAGE="65001" %>
<%
  Response.CharSet="utf-8"
  Session.codepage="65001"
  Response.codepage="65001"
  Response.ContentType="text/html;charset=utf-8"
%>
<!--#include virtual="/include/nkpmg_dbcon_nologin.asp" -->
<%
  p_DocId      = request("DocId")
  p_date       = request("date")
  p_status     = request("status")
  p_userid     = request("userid")
  p_empno      = request("empno")
  p_comment    = request("comment")
  p_next_users = request("next_users")
  p_title      = request("title")

  Set Dbconn = Server.CreateObject("ADODB.Connection")
  Set Rs = Server.CreateObject("ADODB.Recordset")

  DbConn.Open dbconnect

  '' �ݹ� ȣ�� Ȯ��....... ȣ���� �ȵŸ� �ڷᰡ ��ϵ����ʴ´�.
	sql="insert into hanbiro_callback (DocId,date,status,userid,empno,comment,next_users,title,reg_date) "
	sql=sql & "values ('"&p_DocId&"','"&p_date&"','"&p_status&"','"&p_userid&"','"&p_empno&"','"&p_comment&"','"&p_next_users&"','"&p_title&"',now())"
	'dbconn.execute(sql)

  dbconn.Close()
	Set dbconn = Nothing
%>