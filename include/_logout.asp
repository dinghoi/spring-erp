<!--#include virtual = "/common/inc_top.asp"--><!--���� ����-->
<%
'==========================
'Author : ����ȣ
'Modify Date : 20201125
'Desc :
'	���� ���� include �߰�
'	�ڵ� ���� �� ���̵�/������ ��Ű �ʱ�ȭ �߰�
'==========================

Response.Cookies("nkpmg_user")("coo_user_id") = ""
Response.Cookies("nkpmg_user")("coo_user_name") = ""
Response.Cookies("nkpmg_user")("coo_emp_no") = ""

'��Ű ���� ����
Response.Cookies("nkpmg_user").Expires = Date - 1

Response.Redirect "index.asp"

Response.Write "<script type='text/javascript'>"
Response.Write "	window.close();"
Response.Write" 	sign_process_mg_pop.close();"
Response.Write "</script>"

Response.End
%>