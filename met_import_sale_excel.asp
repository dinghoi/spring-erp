<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
view_condi=Request("view_condi")
stock=request("stock")
goods_type=request("goods_type")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = " N/W ���� ��� ��Ȳ -- "+ goods_type +" (" + from_date + " �� " + to_date + ")"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

chulgo_id = "NW���"

order_Sql = " ORDER BY chulgo_date DESC"
if view_condi = "��ü" then
   if goods_type = "��ü" then
      where_sql = " WHERE (chulgo_date >= '"&from_date&"' and chulgo_date <= '"&to_date&"') and chulgo_id = '"&chulgo_id&"'" 
	  else
	  where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (chulgo_date >= '"&from_date&"' and chulgo_date <= '"&to_date&"') and chulgo_id = '"&chulgo_id&"'" 
   end if
 else  
   if goods_type = "��ü" then
      where_sql = " WHERE (chulgo_stock_company = '"&view_condi&"') and (chulgo_date >= '"&from_date&"' and chulgo_date <= '"&to_date&"') and chulgo_id = '"&chulgo_id&"'"
	  else
	  where_sql = " WHERE (chulgo_goods_type = '"&goods_type&"') and (chulgo_stock_company = '"&view_condi&"') and (chulgo_date >= '"&from_date&"' and chulgo_date <= '"&to_date&"') and chulgo_id = '"&chulgo_id&"'"
   end if
end if   

if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (chulgo_stock_name like '%"&stock&"%') "
end if

sql = "select * from met_chulgo " + where_sql + stock_sql + order_sql 
Rs.Open Sql, Dbconn, 1
	

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��ǰ������� �ý���</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">�������</th>
				                <th scope="col">����ȣ</th>
                                <th scope="col">���â��</th>
                                <th scope="col">�뵵����</th>
				                <th scope="col">��ǥ��ȣ</th>
                                <th scope="col">����</th>
                                <th scope="col">�����</th>
                                <th scope="col">���ݾ�</th>
                                <th scope="col">���VAT</th>
                                <th scope="col">����Ѿ�</th>
                                <th scope="col">����</th>

                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
                                <th scope="col">ǰ���</th>
                                <th scope="col">�԰�</th>
                                <th scope="col">Part_No.</th>
                                <th scope="col">������</th>
                                <th scope="col">�ݾ�</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   chulgo_date = rs("chulgo_date")
						   
						   chulgo_stock = rs("chulgo_stock")
						   chulgo_seq = rs("chulgo_seq")

						   k = 0
                           sql = "select * from met_chulgo_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')  ORDER BY cg_goods_seq,cg_goods_code ASC"
	                       Rs_buy.Open Sql, Dbconn, 1	
	                       while not Rs_buy.eof
		                     k = k + 1
							 if k = 1 then
							 
							     stock_goods_code = Rs_buy("cg_goods_code")
								 sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
                                 Set Rs_good = DbConn.Execute(SQL)
                                 if not Rs_good.eof then
    	                               goods_model = Rs_good("goods_model")
		                               part_number = Rs_good("part_number")
                                    else
		                               goods_model = ""
		                               part_number = ""
                                 end if
                                 Rs_good.close()
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("rele_no")%>&nbsp;<%=rs("rele_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_stock_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("service_no")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_trade_name")%>-<%=rs("chulgo_trade_dept")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_emp_name")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("chulgo_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("chulgo_cost_vat"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("chulgo_price"),0)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_memo")%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("cg_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("cg_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("cg_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("cg_standard")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=part_number%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("cg_qty"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("cg_amt"),0)%></td>
						         </tr>
            <%
			                    else
		    %>		
                                 <tr>
								    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td class="left" ><%=Rs_buy("cg_goods_gubun")%></td>
                                    <td class="left" ><%=Rs_buy("cg_goods_code")%></td>
                                    <td class="left" ><%=Rs_buy("cg_goods_name")%></td>
                                    <td class="left" ><%=Rs_buy("cg_standard")%></td>
                                    <td class="left" ><%=part_number%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("cg_qty"),0)%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("cg_amt"),0)%></td>
						         </tr>            
            <%            							
							 end if
		                         Rs_buy.movenext()
	                       Wend
                           Rs_buy.close()
							  
						   rs.movenext()
						loop
						rs.close()
		    %>						
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
