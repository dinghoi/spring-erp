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

title_line = " â���̵� ��� ��Ȳ -- "+ goods_type +" (" + from_date + " �� " + to_date + ")"

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_stock = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

chulgo_id = "â���̵�"

order_Sql = " ORDER BY chulgo_date,chulgo_stock,chulgo_seq DESC"

if view_condi = "��ü" then
   if goods_type = "��ü" then
          where_sql = " WHERE (chulgo_id = '"&chulgo_id&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"')" 
	  else
	      where_sql = " WHERE (chulgo_id = '"&chulgo_id&"') and (chulgo_goods_type = '"&goods_type&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"')" 
   end if
 else  
   if goods_type = "��ü" then
          where_sql = " WHERE (chulgo_id = '"&chulgo_id&"') and (chulgo_stock_company = '"&view_condi&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"')"
	  else
	      where_sql = " WHERE (chulgo_id = '"&chulgo_id&"') and (chulgo_goods_type = '"&goods_type&"') and (chulgo_stock_company = '"&view_condi&"') and (chulgo_date >= '"+from_date+"' and chulgo_date <= '"+to_date+"')"
   end if
end if   


if stock = "" then
       stock_sql = ""
   else
       stock_sql = " and (chulgo_stock_name like '%"&stock&"%') "
end if

sql = "select * from met_mv_go " + where_sql + stock_sql + order_sql
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
                                <th scope="col">�뵵����</th>
                                <th scope="col">�������</th>
                                <th scope="col">����ȣ</th>
                                <th scope="col">������</th>
                                <th scope="col">���â��</th>
                                <th scope="col">�����</th>
                                <th scope="col">��û��(No)</th>
                                <th scope="col">��ûâ��</th>
                                <th scope="col">��û���</th>
                                <th scope="col">����</th>

                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
                                <th scope="col">ǰ���</th>
                                <th scope="col">�԰�</th>
                                <th scope="col">����</th>
                                <th scope="col">�Ƿڼ���</th>
                                <th scope="col">������</th>
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
						   
						   rele_date = rs("rele_date")
						   rele_stock = rs("rele_stock")
						   rele_seq = rs("rele_seq")
					       
						   sql = "select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')"
						   Set Rs_reg=DbConn.Execute(Sql)
						   if Rs_reg.eof or Rs_reg.bof then
								rele_stock_name = ""
								rele_emp_name = ""
							  else
							  	rele_stock_name = Rs_reg("rele_stock_name")
								rele_emp_name = Rs_reg("rele_emp_name")
						   end if
						   Rs_reg.close()
						   
						   chulgo_no = mid(cstr(rs("chulgo_date")),3,2) + mid(cstr(rs("chulgo_date")),6,2) + mid(cstr(rs("chulgo_date")),9,2) 
						   rele_no = mid(cstr(rs("rele_date")),3,2) + mid(cstr(rs("rele_date")),6,2) + mid(cstr(rs("rele_date")),9,2) 

						   k = 0
                           sql = "select * from met_mv_go_goods where (chulgo_date = '"&chulgo_date&"') and (chulgo_stock = '"&chulgo_stock&"') and (chulgo_seq = '"&chulgo_seq&"')  ORDER BY cg_goods_seq,cg_goods_code ASC"
	                       Rs_reg.Open Sql, Dbconn, 1	
	                       while not Rs_reg.eof
		                     k = k + 1
							 if k = 1 then
		    %>
                                 <tr>
								    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_goods_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_date")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=chulgo_no%>&nbsp;<%=rs("chulgo_stock")%><%=rs("chulgo_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_type")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_stock_name")%>(<%=rs("chulgo_stock")%>)</td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_emp_name")%>(<%=rs("chulgo_emp_no")%>)</td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rele_no%>&nbsp;<%=rs("rele_stock")%><%=rs("rele_seq")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rele_stock_name%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rele_emp_name%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=rs("chulgo_memo")%></td>
                                    
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("cg_goods_gubun")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("cg_goods_code")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("cg_goods_name")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("cg_standard")%></td>
                                    <td class="left" bgcolor="#EEFFFF"><%=Rs_reg("cg_goods_grade")%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_reg("rl_qty"),0)%></td>
                                    <td bgcolor="#EEFFFF" align="right"><%=formatnumber(Rs_reg("cg_qty"),0)%></td>
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
                                    
								    <td class="left" ><%=Rs_reg("cg_goods_gubun")%></td>
                                    <td class="left" ><%=Rs_reg("cg_goods_code")%></td>
                                    <td class="left" ><%=Rs_reg("cg_goods_name")%></td>
                                    <td class="left" ><%=Rs_reg("cg_standard")%></td>
                                    <td class="left" ><%=Rs_reg("cg_goods_grade")%></td>
                                    <td align="right"><%=formatnumber(Rs_reg("rl_qty"),0)%></td>
                                    <td align="right"><%=formatnumber(Rs_reg("cg_qty"),0)%></td>
						         </tr>            
            <%            							
							 end if
		                         Rs_reg.movenext()
	                       Wend
                           Rs_reg.close()
							  
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
