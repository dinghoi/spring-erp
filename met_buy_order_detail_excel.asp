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

title_line = goods_type + " ���� ��Ȳ -- "+ view_condi +" (" + from_date + " �� " + to_date + ")"

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
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY order_company,order_date,order_no,order_seq DESC"
if view_condi = "��ü" then
   if goods_type = "��ü" then
      where_sql = " WHERE (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')" 
	  else
	  where_sql = " WHERE (order_goods_type = '"&goods_type&"') and (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "��ü" then
      where_sql = " WHERE (order_company = '"&view_condi&"') and (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')"
	  else
	  where_sql = " WHERE (order_goods_type = '"&goods_type&"') and (order_company = '"&view_condi&"') and (order_date >= '"+from_date+"' and order_date <= '"+to_date+"')"
   end if
end if

sql = "select * from met_order " + where_sql + order_sql
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
								<th class="first" scope="col">���Ź�ȣ</th>
                                <th scope="col">�뵵����</th>
				                <th scope="col">����ǰ����</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">�����</th>
                                <th scope="col">�μ�</th>
                                <th scope="col">���Ŵ��</th>
                                <th scope="col">���ְŷ�ó</th>
                                <th scope="col">����ڹ�ȣ</th>
                                <th scope="col">�ŷ�ó�����</th>
                                <th scope="col">�̸���</th>
                                <th scope="col">�����Ѿ�</th>
                                <th scope="col">���ֱݾ�</th>
                                <th scope="col">�ΰ���</th>
                                <th scope="col">�������</th>
                                <th scope="col">����</th>
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
                                <th scope="col">ǰ���</th>
                                <th scope="col">�԰�</th>
                                <th scope="col">����</th>
                                <th scope="col">���Դܰ�</th>
                                <th scope="col">���־�</th>
							</tr>
						</thead>
						<tbody>
			<%
						i = 0
						do until rs.eof
                           i = i + 1
						   order_no = rs("order_no")
						   order_seq = rs("order_seq")
						   order_date = rs("order_date")
						   
						   buy_no = rs("order_buy_no")
						   buy_seq = rs("order_buy_seq")
						   buy_date = rs("order_buy_date")
						   
						   order_ing = rs("order_ing")
						   buy_ing = rs("order_ing")
						   buy_ing_gubun = ""
						   if buy_ing = "0" then
						         buy_ing_gubun = "����ǰ��"
						      elseif buy_ing = "1" then
							            buy_ing_gubun = "�κй���"
									 elseif buy_ing = "2" then
							                   buy_ing_gubun = "��ü����"
										    elseif buy_ing = "3" then
							                          buy_ing_gubun = "���ֿϷ�"
												   elseif buy_ing = "4" then
							                                 buy_ing_gubun = "�԰�"
						   end if
					   
						   k = 0
                           sql = "select * from met_order_goods where (og_order_no = '"&order_no&"') and (og_order_seq = '"&order_seq&"') and (og_order_date = '"&order_date&"')  ORDER BY og_seq,og_goods_code ASC"
	                       Rs_buy.Open Sql, Dbconn, 1	
	                       while not Rs_buy.eof
		                     k = k + 1
							 if k = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("order_no")%>-<%=rs("order_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_company")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_saupbu")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_org_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_emp_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_trade_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_trade_no")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_trade_person")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_trade_email")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("order_price"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("order_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(rs("order_cost_vat"),0)%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=buy_ing_gubun%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("order_memo")%></td>
                                    
								    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("og_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("og_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("og_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_buy("og_standard")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("og_qty"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("og_unit_cost"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_buy("og_amt"),0)%></td>
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
                                    <td align="right">&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td align="right">&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    <td class="left" >&nbsp;</td>
                                    
								    <td class="left" ><%=Rs_buy("og_goods_gubun")%></td>
                                    <td class="left" ><%=Rs_buy("og_goods_code")%></td>
                                    <td class="left" ><%=Rs_buy("og_goods_name")%></td>
                                    <td class="left" ><%=Rs_buy("og_standard")%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("og_qty"),0)%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("og_unit_cost"),0)%></td>
                                    <td align="right"><%=formatnumber(Rs_buy("og_amt"),0)%></td>
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
