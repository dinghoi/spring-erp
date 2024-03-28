<%

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set rs_reside = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_ddd = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_next = Server.CreateObject("ADODB.Recordset")
Set rs_pre = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

sql_trade="select * from trade where use_sw = 'Y' and ( trade_id = '매출' or trade_id = '공용' ) order by trade_name asc"



%>
