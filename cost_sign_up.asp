<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
sign_month=Request("sign_month")
sign_pro=Request("sign_pro")

sign_id = "01"

from_date = cstr(mid(sign_month,1,4)) + "-" + cstr(mid(sign_month,5,2)) + "-" + "01"
to_date = cstr(mid(sign_month,1,4)) + "-" + cstr(mid(sign_month,5,2)) + "-" + "31"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_cash = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' ���Ǻ� ��ȸ.........
base_sql = "select * from general_cost where (slip_date >= '" + from_date  + "' and slip_date <= '" + to_date  + "') and "

condi_sql = " reg_id = '" + user_id + "'"
if position = "����" then
	condi_sql = "bonbu = '"&bonbu&"' and saupbu = '"&saupbu&"' and team = '"&team&"'"
end if

order_sql = " ORDER BY slip_date ASC"

sql = base_sql + condi_sql + order_sql
Rs.Open Sql, Dbconn, 1

if position = "����" then
	sign_title = team
end if
if position = "�������" then
	sign_title = saupbu
end if
if position = "����" then
	sign_title = bonbu
end if

sign_head = mid(sign_month,1,4) + "��" + mid(sign_month,5,2) + "��" + " �Ϲݰ�� �����Ȳ - ( " + sign_title + " )" 
sub_title_line = ". �ۼ��� : " + user_name + "( " + user_id + " )"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.sign_memo.value =="") {
					alert('Ư�̻����� �Է��ϼ���.');
					frm.sign_memo.focus();
					return false;}
				{
				a=confirm('���縦 ����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
		</script>

	</head>
	<body>
		<div id="form_wrap">			
			<div id="container">
				<h3 class="tit"><%=sign_head%></h3>
				<br>
				<h3 class="stit"><%=sub_title_line%></h3>
				<form action="cost_sign_up_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="11%" >
							<col width="8%" >
							<col width="8%" >
							<col width="6%" >
							<col width="7%" >
							<col width="6%" >
							<col width="*" >
							<col width="8%" >
							<col width="5%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�߻�����</th>
								<th scope="col">�Ҽ�</th>
								<th scope="col">�׸�</th>
								<th scope="col">���ڰ���NO</th>
								<th scope="col">��뱸��</th>
								<th scope="col">�ݾ�</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">�߻�����/�ŷ�ó</th>
								<th scope="col">�����</th>
								<th scope="col">����</th>
								<th scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0
						do until rs.eof
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							if rs("pay_yn") = "Y" then
								pay_yn = "����"
							  else
							  	pay_yn = "������"
							end if
							if rs("end_yn") = "Y" then
								end_yn = "����"
								end_view = "N"
							  elseif rs("end_yn") = "I" then
								end_yn = "������"
								end_view = "N"
							  else
							  	end_yn = "����"
							end if
							belong = rs("team") + " " + rs("belong")
							if rs("team") = "" then
								belong = rs("saupbu")
							end if
							if belong = "" then
								belong = rs("bonbu")
							end if
						%>
							<tr>
								<td class="first"><%=rs("slip_date")%></td>
								<td><%=belong%></td>
								<td><%=rs("account_item")%></td>
								<td><%=rs("sign_no")%></td>
								<td><%=rs("pay_method")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("customer")%></td>
								<td><%=rs("use_man")%><%=rs("user_grade")%>&nbsp;</td>
								<td><%=pay_yn%></td>
								<td><%=rs("slip_memo")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
							<tr>
								<th class="first" colspan="5">�� ��</th>
							  	<th class="right"><%=formatnumber(cost_sum,0)%></th>
							  	<th class="right"><%=formatnumber(cost_vat_sum,0)%></th>
							  	<th class="right" colspan="4">&nbsp;</th>
							</tr>
							<tr>
								<td class="first" bgcolor="#CCFFFF">Ư�̻���</td>
						  	  <td class="left" colspan="10"><textarea name="sign_memo" cols="140" rows="3" id="textarea"></textarea></td>
						  </tr>
						</tbody>
					</table>
				</div>
				<br>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
		            <div align=center>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>
					</div>                  
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="sign_month" value="<%=sign_month%>" ID="Hidden1">
				<input type="hidden" name="sign_pro" value="<%=sign_pro%>" ID="Hidden1">
				<input type="hidden" name="sign_id" value="<%=sign_id%>" ID="Hidden1">
				<input type="hidden" name="sign_head" value="<%=sign_head%>" ID="Hidden1">
				<input type="hidden" name="from_date" value="<%=from_date%>" ID="Hidden1">
				<input type="hidden" name="to_date" value="<%=to_date%>" ID="Hidden1">
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

