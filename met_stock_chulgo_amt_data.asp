<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
'	on Error resume next

emp_user = request.cookies("nkpmg_user")("coo_user_name")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set Rs_jago = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

i = 0
j = 0

Sql = "select * from met_stock_gmaster"
Rs.Open Sql, Dbconn, 1
if not Rs.eof then
   do until Rs.eof

    if rs("stock_in_amt") > 0 then
	   j = j + 1
	   
	   stock_code = rs("stock_code")
	   stock_goods_type = rs("stock_goods_type")
	   stock_goods_code = rs("stock_goods_code")
	   
	   in_amt = 0
	   ch_amt = 0
	   in_amt = rs("stock_in_amt") / rs("stock_in_qty") 
	   if rs("stock_go_qty") > rs("stock_in_qty") then
	         ch_amt = rs("stock_in_qty") * in_amt
		  else
	         ch_amt = rs("stock_go_qty") * in_amt
		end if
	   
'	   ja_amt = rs("stock_jj_amt") - ch_amt
'	   ja_amt = rs("stock_JJ_qty") * in_amt
       ja_amt = rs("stock_in_amt") - ch_amt
		
	   sql = "update met_stock_gmaster set stock_go_amt='"&ch_amt&"',stock_jj_amt='"&ja_amt&"' where stock_code='"&stock_code&"' and stock_goods_type='"&stock_goods_type&"' and stock_goods_code='"&stock_goods_code&"'"
		
	   dbconn.execute(sql)
    end if

	Rs.MoveNext()
  loop		
		response.write"<script language=javascript>"
		response.write"alert('출고 금액이 갱신되었습니다...."&j&"');"		
		response.write"location.replace('met_stock_pum_jaego_mg.asp');"
		response.write"</script>"
		Response.End
else
		response.write"<script language=javascript>"
		response.write"alert(' 처리된 내역이없습니다...');"		
		response.write"location.replace('met_stock_pum_jaego_mg.asp');"
		response.write"</script>"
		Response.End
end if	

dbconn.Close()
Set dbconn = Nothing
	
%>
