<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
approve_no = request("approve_no")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql="select * from card_slip where approve_no = '"&approve_no&"'"
Set rs=DbConn.Execute(Sql)

title_line = "ī�� �ŷ�ó ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>����ȸ��ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.cust_no1.value =="") {
					alert('����ڹ�ȣ1�� �Է��ϼ���.');
					frm.cust_no1.focus();
					return false;}
				if(document.frm.cust_no2.value =="") {
					alert('����ڹ�ȣ2�� �Է��ϼ���.');
					frm.cust_no2.focus();
					return false;}
				if(document.frm.cust_no3.value =="") {
					alert('����ڹ�ȣ3�� �Է��ϼ���.');
					frm.cust_no3.focus();
					return false;}
				if(document.frm.customer.value =="") {
					alert('����ŷ�ó���� �Է��ϼ���.');
					frm.account_item.focus();
					return false;}
				{
				a=confirm('�Է��Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="card_customer_mod_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="20%" >
							<col width="30%" >
							<col width="20%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">�����ŷ�ó</th>
								<td class="left"><%=rs("customer_no")%></td>
                                <th>�ŷ�ó��</th>
								<td class="left"><%=rs("customer")%></td>
							</tr>
							<tr>
								<th class="first">����ŷ�ó</th>
								<td class="left">
                                <input name="cust_no1" type="text" id="cust_no1" style="width:25px; text-align:center" maxlength="3" value="<%=cust_no1%>" onKeyUp="checkNum(this);">
                                -
                                <input name="cust_no2" type="text" id="cust_no2" style="width:20px; text-align:center" maxlength="2" value="<%=cust_no2%>" onKeyUp="checkNum(this);">
                                -
                                <input name="cust_no3" type="text" id="cust_no3" style="width:50px; text-align:center" maxlength="5" value="<%=cust_no3%>" onKeyUp="checkNum(this);">
                                </td>
								<th>����ŷ�ó��</th>
								<td class="left"><input name="customer" type="text" id="customer" style="width:150px;" onKeyUp="checklength(this,30);"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="approve_no" value="<%=approve_no%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

