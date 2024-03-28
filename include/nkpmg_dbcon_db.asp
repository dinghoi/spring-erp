<%
Dim db_host, db_name, db_user, db_pass
Dim DBConnect
Dim http_host, server_port

http_host = Request.ServerVariables("HTTP_HOST")
server_port = Request.ServerVariables("SERVER_PORT")

'기존 주석
'DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=localhost;DATABASE=nkp;UID=root;PWD=ekgus0930;"

'운영 서버 접속 정보 구분
If http_host = "intra.k-won.co.kr" And server_port = "80" Then
	db_host = "localhost"
	db_name = "nkp"
	db_user = "nkp"
	db_pass = "nkp2014"
Else
	'db_host = "211.43.210.66"
	'db_name = "nkp_dev"
	'db_user = "nkp_dev"
	'db_pass = "zpdldnjs!@3"
	'Response.write "# DB : " & db_name & "<br/>"

	db_host = "localhost"
	db_name = "nkp"
	db_user = "root"
	db_pass = "duckling"
End If

DBConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER="&db_host&";DATABASE="&db_name&";UID="&db_user&";PWD="&db_pass&";"
%>
