<!--#include virtual = "/common/inc_top.asp"--><!--설정 파일-->
<%
'==========================
'Author : 허정호
'Modify Date : 20201125
'Desc :
'	설정 파일 include 추가
'	코드 정리 및 아이디/직원명 쿠키 초기화 추가
'==========================

Response.Cookies("nkpmg_user")("coo_user_id") = ""
Response.Cookies("nkpmg_user")("coo_user_name") = ""
Response.Cookies("nkpmg_user")("coo_emp_no") = ""

'쿠키 만료 설정
Response.Cookies("nkpmg_user").Expires = Date - 1

Response.Redirect "index.asp"

Response.Write "<script type='text/javascript'>"
Response.Write "	window.close();"
Response.Write" 	sign_process_mg_pop.close();"
Response.Write "</script>"

Response.End
%>