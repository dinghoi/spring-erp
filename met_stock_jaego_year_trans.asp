<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = now()
be_date = dateadd("yyyy",-1,curr_date)
be_year = cstr(mid(be_date,1,4)) 
be_month = cstr(mid(be_date,1,4)) + cstr(mid(be_date,6,2))


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_bef = Server.CreateObject("ADODB.Recordset")
Set Rs_jae = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Dbconn.BeginTrans

'sql = "delete from met_stock_gmaster_year where emp_month ='"&be_month&"'"
'    dbconn.execute(sql)	

sql = "insert into met_stock_gmaster_year select '"&be_year&"' as stock_year,met_stock_gmaster.* from met_stock_gmaster"
    dbconn.execute(sql)


Sql = "select * from met_stock_gmaster where (stock_JJ_qty > 0 ) ORDER BY stock_code,stock_goods_type,stock_goods_code ASC"
Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

        j = j + 1
	    stock_code = rs("stock_code")
	    stock_goods_type = rs("stock_goods_type")
	    stock_goods_code = rs("stock_goods_code")
		
		stock_last_qty = rs("stock_last_qty")
		stock_in_qty = rs("stock_in_qty")
		stock_go_qty = rs("stock_go_qty")
		stock_JJ_qty = rs("stock_JJ_qty")
		
		stock_last_amt = rs("stock_last_amt")
		stock_in_amt = rs("stock_in_amt")
		stock_go_amt = rs("stock_go_amt")
		stock_jj_amt = rs("stock_jj_amt")
		
		year_qty = stock_last_qty + stock_in_qty - stock_go_qty
		year_amt = stock_last_amt + stock_in_amt - stock_go_amt
		
		stock_in_qty = 0
		stock_go_qty = 0
		stock_in_amt = 0
		stock_go_amt = 0
		
	
        sql = "update met_stock_gmaster set stock_last_qty='"&year_qty&"',stock_last_amt='"&year_amt&"',stock_in_qty='"&stock_in_qty&"',stock_in_amt='"&stock_in_amt&"',stock_go_qty='"&stock_go_qty&"',stock_go_amt='"&stock_go_amt&"',stock_JJ_qty='"&year_qty&"',stock_jj_amt='"&year_amt&"' where (stock_code = '"&stock_code&"' ) and (stock_goods_type = '"&stock_goods_type&"') and (stock_goods_code = '"&stock_goods_code&"')"
		
	    dbconn.execute(sql)
		   
	 Rs.MoveNext()
  loop		
end if





end_sw = "Y" 

if err.number <> 0 then
	Dbconn.RollbackTrans 
else    
	Dbconn.CommitTrans 
	response.write"<script language=javascript>"
	response.write"alert('"&be_year&"...재고 전기이월처리가 되었습니다...("&j&") 건');"		
	response.write"location.replace('met_stock_pum_jaego_mg.asp');"
	response.write"</script>"
	Response.End
end if

dbconn.Close()
Set dbconn = Nothing
	
%>
