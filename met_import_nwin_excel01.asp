<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
view_condi=Request("view_condi")
goods_type=request("goods_type")
from_date=request("from_date")
to_date=request("to_date")

curr_date = datevalue(mid(cstr(now()),1,10))

title_line = goods_type + " �����԰� ��Ȳ -- "+ view_condi +" (" + from_date + " �� " + to_date + ")"

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

order_Sql = " ORDER BY stin_in_date DESC"  

if view_condi = "��ü" then
   if goods_type = "��ü" then
      where_sql = " WHERE (stin_id = '���ڱ���') and (stin_in_date >= '"&from_date&"' and stin_in_date <= '"&to_date&"')" 
	  else
	  where_sql = " WHERE (stin_id = '���ڱ���') and (stin_goods_type = '"&goods_type&"') and (stin_in_date >= '"&from_date&"' and stin_in_date <= '"&to_date&"')" 
   end if
 else  
   if goods_type = "��ü" then
      where_sql = " WHERE (stin_id = '���ڱ���') and (stin_stock_name = '"&view_condi&"') and (stin_in_date >= '"&from_date&"' and stin_in_date <= '"&to_date&"')"
	  else
	  where_sql = " WHERE (stin_id = '���ڱ���') and (stin_goods_type = '"&goods_type&"') and (stin_stock_name = '"&view_condi&"') and (stin_in_date >= '"&from_date&"' and stin_in_date <= '"&to_date&"')"
   end if
end if   

sql = "select * from met_stin " + where_sql + order_sql 
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
								<th class="first" scope="col">�԰�����</th>
                                <th scope="col">�뵵����</th>
                                <th scope="col">�԰��ȣ</th>
                                <th scope="col">�԰���</th>
                                <th scope="col">�׷��</th>
                                <th scope="col">�����</th>
                                <th scope="col">�԰�â��</th>
                                <th scope="col">�԰��Ѿ�</th>
                                <th scope="col">�԰�ݾ�</th>
                                <th scope="col">�ΰ���</th>

                                <th scope="col">���Űŷ�ó</th>
                                <th scope="col">����ڹ�ȣ</th>
                                
                                <th scope="col">PO ����</th>
                                <th scope="col">PO_Number</th>
                                <th scope="col">Parking_Number(BL)</th>
                                
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
                                <th scope="col">ǰ���</th>
                                <th scope="col">Part_Number</th>
                                <th scope="col">�԰����</th>
                                <th scope="col">�ܰ�($)</th>
                                <th scope="col">�ܰ�(��ȭ)</th>
                                <th scope="col">���ܰ�</th>
                                <th scope="col">�԰�ܰ�</th>
                                <th scope="col">�԰�ݾ�</th>
                                <th scope="col">���ȯ��</th>
                                <th scope="col">÷��_serial</th>
                                
                                <th scope="col">���</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   stin_in_date = rs("stin_in_date")
						   
						   stin_order_no = rs("stin_order_no")
						   stin_order_seq = rs("stin_order_seq")
						   
						   stin_trade_no = mid(rs("stin_trade_no"),1,3) + "-" + mid(rs("stin_trade_no"),4,2) + "-" + mid(rs("stin_trade_no"),6)

						   k = 0
                           sql = "select * from met_stin_goods where (stin_date = '"&stin_in_date&"') and (stin_order_no = '"&stin_order_no&"') and (stin_order_seq = '"&stin_order_seq&"')  ORDER BY stin_goods_seq,stin_goods_code ASC"
	                       Rs_buy.Open Sql, Dbconn, 1	
	                       while not Rs_buy.eof
						   
						     unit_wonga = Rs_buy("d_cost") * Rs_buy("w_won")
		                     k = k + 1
							 if k = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_in_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_order_no")%>-<%=rs("stin_order_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_id")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_buy_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_buy_saupbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_stock_name")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("stin_price"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("stin_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("stin_cost_vat"),0)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_trade_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=stin_trade_no%></td>
                                    
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("po_date")%></td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("po_number")%></td>
                                    <td align="left" bgcolor="#EEFFFF"><%=rs("park_bl")%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("stin_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("stin_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("stin_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("part_number")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("stin_qty"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("d_cost"),2)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(unit_wonga,0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("ex_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("stin_unit_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("stin_amt"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("w_won"),2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("excel_att_file")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("stin_memo")%></td>
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
                                    <td class="left" >&nbsp;</td>
                                    
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td class="left" ><%=Rs_buy("stin_goods_gubun")%></td>
                                    <td class="left" ><%=Rs_buy("stin_goods_code")%></td>
                                    <td class="left" ><%=Rs_buy("stin_goods_name")%></td>
                                    <td class="left" ><%=Rs_buy("part_number")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("stin_qty"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("d_cost"),2)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(unit_wonga,0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("ex_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("stin_unit_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("stin_amt"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("w_won"),2)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("excel_att_file")%></td>
                                    <td align="right">&nbsp;</td>
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
