<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

order_date = request("order_date")
order_buy_no = request("order_buy_no")
order_buy_date = request("order_buy_date")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_order = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_buy where (buy_no = '"&order_buy_no&"') and (buy_date = '"&order_buy_date&"')"
Set Rs_buy = DbConn.Execute(SQL)
if not Rs_buy.eof then
    	buy_no = Rs_buy("buy_no")
		buy_date = Rs_buy("buy_date")
		buy_company = Rs_buy("buy_company")
	    buy_saupbu = Rs_buy("buy_saupbu")
	    buy_org_code = Rs_buy("buy_org_code")
	    buy_org_name = Rs_buy("buy_org_name")
	    buy_emp_no = Rs_buy("buy_emp_no")
	    buy_emp_name = Rs_buy("buy_emp_name")
	    buy_bill_collect = Rs_buy("buy_bill_collect")
        buy_collect_due_date = Rs_buy("buy_collect_due_date")
	    buy_trade_no = Rs_buy("buy_trade_no")
        buy_trade_name = Rs_buy("buy_trade_name")
        buy_trade_person = Rs_buy("buy_trade_person")
        buy_out_method = Rs_buy("buy_out_method")
        buy_out_request_date = Rs_buy("buy_out_request_date")
        buy_price = Rs_buy("buy_price")
        buy_cost = Rs_buy("buy_cost")
        buy_cost_vat = Rs_buy("buy_cost_vat")
        buy_memo = Rs_buy("buy_memo")
        buy_ing = Rs_buy("buy_ing")
		buy_att_file = Rs_buy("buy_att_file")

	    if buy_out_request_date = "0000-00-00" then
	          buy_out_request_date = ""
	    end if
   else
		buy_company = ""
	    buy_saupbu = ""
	    buy_org_code = ""
	    buy_org_name = ""
	    buy_emp_no = ""
	    buy_emp_name = ""
	    buy_bill_collect = ""
        buy_collect_due_date = ""
	    buy_trade_no = ""
        buy_trade_name = ""
        buy_trade_person = ""
        buy_out_method = ""
        buy_out_request_date = ""
        buy_price = 0
        buy_cost = 0
        buy_cost_vat = 0
        buy_memo = ""
        buy_ing = ""
		buy_att_file = ""
end if
Rs_buy.close()

sql = "select * from met_order where (order_date = '"&order_date&"') and (order_buy_no = '"&order_buy_no&"')"
Set Rs_order = DbConn.Execute(SQL)
if not Rs_order.eof then
   	order_buy_no = Rs_order("order_buy_no")
	order_date = Rs_order("order_date")
	order_buy_date = Rs_order("order_buy_date")
	order_goods_type = Rs_order("order_goods_type")
	order_company = Rs_order("order_company")
    order_saupbu = Rs_order("order_saupbu")
    order_org_code = Rs_order("order_org_code")
    order_org_name = Rs_order("order_org_name")
	order_emp_no = Rs_order("order_emp_no")
    order_emp_name = Rs_order("order_emp_name")
    order_bill_collect = Rs_order("order_bill_collect")
    order_collect_due_date = Rs_order("order_collect_due_date")
    order_trade_no = Rs_order("order_trade_no")
    order_trade_name = Rs_order("order_trade_name")
    order_trade_person = Rs_order("order_trade_person")
    order_in_date = Rs_order("order_in_date")
    order_stock_company = Rs_order("order_stock_company")
    order_stock_code = Rs_order("order_stock_code")
    order_stock_name = Rs_order("order_stock_name")
    order_out_method = Rs_order("order_out_method")
    order_out_request_date = Rs_order("order_out_request_date")
    order_price = Rs_order("order_price")
    order_cost = Rs_order("order_cost")
    order_cost_vat = Rs_order("order_cost_vat")
    order_memo = Rs_order("order_memo")
    order_ing = Rs_order("order_ing")

	if order_out_request_date = "0000-00-00" then
	      order_out_request_date = ""
	end if
	
	if order_in_date = "0000-00-00" then
	      order_in_date = ""
	end if
end if
Rs_order.close()

sql = "select * from met_order_goods where (og_order_date = '"&order_date&"') and (og_buy_no = '"&order_buy_no&"') ORDER BY og_buy_seq,og_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "���� ǰ�� ����"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>������� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}					
			function chkfrm() {
				if(document.frm.in_buy_no.value =="") {
					alert('�����Ƿڹ�ȣ�� �Է��ϼ���');
					frm.in_buy_no.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_order_detail_list.asp?order_buy_no=<%=order_buy_no%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="*" >
							<col width="12%" >
							<col width="20%" >
							<col width="12%" >
							<col width="20%" >
						</colgroup>
						<tbody> 
							<tr>
                                <th style="background:#f8f8f8;">���ſ�ûȸ��</th>
							    <td class="left"><%=buy_company%>&nbsp;</td>
							    <th style="background:#f8f8f8;">��û�����</th>
							    <td class="left"><%=buy_saupbu%>&nbsp;</td>
							    <th style="background:#f8f8f8;">���ſ�û��</th>
							    <td class="left"><%=buy_org_name%>&nbsp;<%=buy_emp_name%></td>
 							</tr>
                            <tr>
							    <th style="background:#f8f8f8;">���ſ�û��</th>
							    <td class="left"><%=buy_date%>&nbsp;</td>
                                <th style="background:#f8f8f8;">�����</th>
							    <td class="left"><%=buy_out_method%>&nbsp;</td>
							    <th style="background:#f8f8f8;">����û��</th>
							    <td class="left"><%=buy_out_request_date%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th style="background:#f8f8f8;">���</th>
							  <td class="left" colspan="5" ><textarea name="buy_memo" rows="3" id="textarea"><%=buy_memo%></textarea></td>
						    </tr>
                            <tr>
							  <th style="background:#f8f8f8;">���ֵ����</th>
							  <td class="left"><%=order_date%></td>
							  <th style="background:#f8f8f8;">���ִ����</th>
							  <td class="left"><%=order_emp_name%>&nbsp;<%=order_emp_no%></td>
                              <th style="background:#f8f8f8;">�Ҽ�</th>
							  <td class="left"><%=order_org_name%>&nbsp;</td>
						    </tr>
							<tr>
							  <th style="background:#f8f8f8;">���ְŷ�ó</th>
							  <td class="left"><%=order_trade_name%>&nbsp;</td>
							  <th style="background:#f8f8f8;">����ڹ�ȣ</th>
							  <td class="left"><%=order_trade_no%>&nbsp;</td>
							  <th style="background:#f8f8f8;">�ŷ�ó<br>�����</th>
							  <td class="left"><%=order_trade_person%>&nbsp;</td>
						    </tr>
							<tr>
							  <th style="background:#f8f8f8;">�������</th>
							  <td colspan="3" class="left">
                              <input type="radio" name="bill_collect" value="����" <% if order_bill_collect = "����" then %>checked<% end if %> style="width:40px" id="Radio3">����
  							  <input type="radio" name="bill_collect" value="����" <% if order_bill_collect = "����" then %>checked<% end if %> style="width:40px" id="Radio4">����
                              <input type="radio" name="bill_collect" value="ī��" <% if order_bill_collect = "ī��" then %>checked<% end if %> style="width:40px" id="Radio3">ī��
  							  <input type="radio" name="bill_collect" value="��ȯ" <% if order_bill_collect = "��ȯ" then %>checked<% end if %> style="width:40px" id="Radio4">��ȯ
                              </td>
							  <th style="background:#f8f8f8;">���޿�����</th>
							  <td class="left"><%=order_collect_due_date%>&nbsp;</td>
						    </tr>
                            <tr>
							  <th style="background:#f8f8f8;">�԰���â��</th>
							  <td colspan="3" class="left"><%=order_stock_company%>&nbsp;-&nbsp;<%=order_stock_name%>(<%=order_stock_code%>)</td>
                              <th style="background:#f8f8f8;">�԰�����</th>
							  <td class="left"><%=order_in_date%>&nbsp;</td>
						    </tr>
						</tbody>
					</table>
                </div>
                <br>
                <h3 class="stit" style="font-size:12px;">�� ���� ���� ���� ��</h3>
            	<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="4%" >
							<col width="8%" >
                            <col width="*" >
                            <col width="10%" >
							<col width="14%" >
							<col width="14%" >
                            <col width="6%" >
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col"><input type="checkbox" name="tot_check" id="tot_check"></th>
								<th scope="col">�뵵����</th>
                                <th scope="col">ǰ�񱸺�</th>
                                <th scope="col">ǰ���ڵ�</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�԰�</th>
                                <th scope="col">�����Ƿ�</th>
								<th scope="col">���ּ���</th>
								<th scope="col">���ִܰ�</th>
								<th scope="col">���ֱݾ�</th>
							</tr>
						</thead>
						<tbody>     
						<%
							do until rs.eof or rs.bof
						
						%>
							<tr>
								<td class="first"><input type="checkbox" name="del_check" id="del_check" value="Y"></td>
                                <td><%=rs("og_goods_type")%>&nbsp;</td>
								<td><%=rs("og_goods_gubun")%>&nbsp;</td>
                                <td><%=rs("og_goods_code")%>&nbsp;</td>
                                <td><%=rs("og_goods_name")%>&nbsp;</td>
                                <td><%=rs("og_standard")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("og_bg_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("og_qty"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("og_unit_cost"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("og_amt"),0)%>&nbsp;</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
						</tbody>
					</table>
                    <br>
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="21%" >
							<col width="13%" >
							<col width="21%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>�����Ѿ�</th>
							  <td class="right"><%=formatnumber(order_price,0)%></td>
							  <th>���ֱݾ�</th>
							  <td class="right"><%=formatnumber(order_cost,0)%></td>
							  <th>�ΰ���</th>
							  <td class="right"><%=formatnumber(order_cost_vat,0)%></td>
						    </tr>
						</tbody>
					</table>
			</div>				
	   </div>     
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >�ݱ�</a>&nbsp;&nbsp;
					</div>
                    <br>       				
	</form>
	</body>
</html>

