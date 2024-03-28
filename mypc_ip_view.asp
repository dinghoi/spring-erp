<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%

	dim userip
'	userip = request.ServerVariables("REMOTE_ADDR")
	userip = request("REMOTE_ADDR")

	response.write("MY IP :")
	response.write(userip)
%>
