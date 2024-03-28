<%

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

%>
