<%

 Dim DbConnect1
'DbConnect1 = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=211.172.241.144;DATABASE=itft2005;UID=itft_acpt;PWD=itft_acpt2015;"
 DbConnect1 = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=211.43.210.66;DATABASE=nkp;UID=nkp;PWD=nkp2014;"



 Dim DbConnect
 DbConnect = "DRIVER={MySQL ODBC 5.3 ansi Driver};SERVER=127.0.0.1;DATABASE=nkp;UID=nkp;PWD=nkp2014;"

if 	request.cookies("nkpmg_user")("coo_user_id") = "" then
	response.write"<script language=javascript>"
	response.write"location.replace('warning.asp');"
	response.write"</script>" 	
end if

set dbconn = server.CreateObject("adodb.connection")
dbconn.open dbconnect

set dbconn1 = server.CreateObject("adodb.connection")
dbconn1.open dbconnect1


%>
