<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim rs
Dim rs_numRows

stock_goods_code = request("stock_goods_code")
stock_code = request("stock_code")
stock_company = request("stock_company")
stock_name = request("stock_name")
stock_goods_type = request("stock_goods_type")
goods_name = request("goods_name")

title_line = " ǰ�� ��.���(�����)��Ȳ -- "+ goods_name 

savefilename = title_line + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set Rs_jae = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_goods_code where (goods_code = '"&stock_goods_code&"')"
Set rs = DbConn.Execute(SQL)
if not rs.eof then
    	goods_code = rs("goods_code")
		goods_grade = rs("goods_grade")
        goods_gubun = rs("goods_gubun")
	    goods_name = rs("goods_name")
	    goods_standard = rs("goods_standard")
	    goods_type = rs("goods_type")
   else
		goods_code = ""
		goods_grade = ""
        goods_gubun = ""
	    goods_name = ""
	    goods_standard = ""
	    goods_type = ""
end if
rs.close()

sql = "select * from met_stock_gmaster where (stock_goods_code = '"&stock_goods_code&"') and (stock_code = '"&stock_code&"') and (stock_goods_type = '"&stock_goods_type&"') ORDER BY stock_company,stock_code ASC"
Set Rs_jae = DbConn.Execute(SQL)
if not Rs_jae.eof then

   stock_level = Rs_jae("stock_level")
   goods_code = Rs_jae("stock_goods_code")
   goods_gubun = Rs_jae("stock_goods_gubun")
   goods_name = Rs_jae("stock_goods_name")
   goods_standard = Rs_jae("stock_goods_standard")
   goods_grade = Rs_jae("stock_goods_grade")
   stock_last_qty = Rs_jae("stock_last_qty")
   stock_JJ_qty = Rs_jae("stock_JJ_qty")
end if
Rs_jae.close()


sql = "select * from met_stock_inout where (stock_goods_code = '"&stock_goods_code&"') and (stock_code = '"&stock_code&"') and (stock_goods_type = '"&stock_goods_type&"') ORDER BY stock_date,id_seq,inout_no,inout_seq ASC"
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
						<tbody> 
							<tr>
                                <th>ȸ��</th>
							    <td class="left"><%=stock_company%>&nbsp;</td>
							    <th>â���</th>
							    <td class="left"><%=stock_name%>&nbsp;</td>
							    <th>â����</th>
							    <td class="left"><%=stock_level%>&nbsp;</td>
 							</tr>
                            <tr>
                                <th>ǰ���ڵ�</th>
							    <td class="left"><%=goods_code%>&nbsp;</td>
							    <th>ǰ���</th>
							    <td class="left"><%=goods_name%>&nbsp;</td>
							    <th>����</th>
							    <td class="left"><%=goods_grade%>&nbsp;</td>
 							</tr>
                            <tr>
							    <th>�뵵����</th>
							    <td class="left"><%=stock_goods_type%>&nbsp;</td>
							    <th>ǰ�񱸺�</th>
							    <td class="left"><%=goods_gubun%>&nbsp;</td>
							    <th>�԰�</th>
							    <td class="left"><%=goods_standard%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                <br>
                <h3 class="stit" style="font-size:12px;">�� �� / ��� ����(����) ��</h3>
            	<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th scope="col">����</th>
                                <th scope="col">�뵵����</th>
                                <th scope="col">����</th>
                                <th scope="col">��ȣ</th>
                                <th scope="col">��û�׷��</th>
                                <th scope="col">��û�����</th>
                                <th scope="col">�԰�â��</th>
                                
                                <th scope="col">����</th>
                                <th scope="col">����</th>
                                <th scope="col">����No/<br>��ǥ��ȣ</th>
                                <th scope="col">����<br>�̿�</th>
                                <th scope="col">�԰�</th>
                                <th scope="col">���</th>
                                <th scope="col">�����</th>
                                <th scope="col">���</th>
							</tr>
						</thead>
						<tbody>     
						    <tr>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>�����̿�</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td align="right"><%=formatnumber(stock_last_qty,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						
						<%
							i = 0
							h_last_qty = stock_last_qty
							h_in_qty = 0
							h_go_qty = 0
							h_jj_qty = stock_JJ_qty
							do until rs.eof or rs.bof
								h_in_qty = h_in_qty + rs("stock_in_qty")
								h_go_qty = h_go_qty + rs("stock_go_qty")
						%>
							<tr>
                                <td><%=rs("stock_date")%>&nbsp;</td>
                                <td><%=rs("stock_goods_type")%>&nbsp;</td>
                                <td><%=rs("stock_id")%>&nbsp;</td>
                                <td><%=rs("inout_no")%>&nbsp;<%=rs("inout_seq")%></td>
                                <td><%=rs("rele_company")%>&nbsp;</td>
                                <td><%=rs("rele_saupbu")%>&nbsp;</td>
                                <td><%=rs("rele_stock_name")%>&nbsp;</td>
                                <td><%=rs("trade_name")%>&nbsp;</td>
                                <td><%=rs("trade_dept")%>&nbsp;</td>
                                <td><%=rs("out_service_no")%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("stock_in_qty"),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(rs("stock_go_qty"),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
								<td><%=rs("chulgo_return")%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
                            <tr>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>���� ���</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td>&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td class="right">&nbsp;</td>
                                <td align="right"><%=formatnumber(stock_JJ_qty,0)%>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
                            <tr>
                                <td colspan="10" style="background:#ffe8e8;">�� ��</td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(h_last_qty,0)%>&nbsp;</td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(h_in_qty,0)%>&nbsp;</td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(h_go_qty,0)%>&nbsp;</td>
                                <td align="right" style="background:#ffe8e8;"><%=formatnumber(h_jj_qty,0)%>&nbsp;</td>
								<td style="background:#ffe8e8;">&nbsp;</td>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
