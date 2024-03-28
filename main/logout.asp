<!--#include virtual = "/common/inc_top.asp"-->
<%
SESSION.ABANDON

Response.Cookies("nkpmg_user").Expires = Date - 1
Response.Redirect "/index.asp"

'Response.Write "<script type='text/javascript'>"
'Response.Write "	window.close();"
'Response.Write" 	sign_process_mg_pop.close();"
'Response.Write "	location.replace('/index.asp');"
'Response.Write "</script>"

Response.End
%>