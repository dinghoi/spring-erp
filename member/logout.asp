<!--#include virtual = "/common/inc_top.asp"-->
<%
SESSION.ABANDON

Response.Cookies("nkp_member").Expires = Date - 1
Response.Redirect "/index.asp"

Response.End
%>