<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
trade_code = request("trade_code")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_acc = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql="select * from trade where trade_code = '"&trade_code&"'"

Response.write sql

Set rs=DbConn.Execute(Sql)
'Response.write Sql

title_line = "������ �ŷ�ó ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//ENrs("customer_no")http://www.w3.org/TR/html4/loose.dtd">
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
				{
				a=confirm('�����Ͻðڽ��ϱ�?')
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
				<form action="trade_cost_mod_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">����ڹ�ȣ</th>
				        <td class="left"><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%>&nbsp;<strong>��ǥ�� :</strong>&nbsp;<%=rs("trade_owner")%></td>
				        <th>��ȣ</th>
				        <td class="left"><%=rs("trade_name")%></td>
			          </tr>
				      <tr>
				        <th class="first">�׷��</th>
				        <td class="left"><input name="group_name" type="text" id="group_name" style="width:170px;" value="<%=rs("group_name")%>" onKeyUp="checklength(this,30);">
                        <a href="#" onClick="pop_Window('trade_search.asp?gubun=<%="5"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a></td>
				        <th>��꼭 ����<br>ȸ���</th>
				        <td class="left"><input name="bill_trade_name" type="text" value="<%=rs("bill_trade_name")%>" style="width:170px"></td>
			          </tr>
				      <tr>
				        <th class="first">�����</th>
				        <td colspan="3" class="left">
                        <input name="emp_no" type="text" id="emp_no" style="width:80px" value="<%=rs("emp_no")%>" readonly="true">
                        <input name="emp_name" type="text" id="emp_name" style="width:80px" value="<%=rs("emp_name")%>" readonly="true">
                        <input name="saupbu" type="text" id="saupbu" style="width:150px" value="<%=rs("saupbu")%>" readonly="true">
                        <a href="#" onClick="pop_Window('emp_search.asp?gubun=<%="2"%>','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">�����ȸ</a>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">�ŷ�ó����</th>
				        <td class="left">
							<input type="radio" name="trade_id" value="����" <% if rs("trade_id") = "����" then %>checked<% end if %> style="width:20px">��������
							<input type="radio" name="trade_id" value="�Ϲ�" <% if rs("trade_id") = "�Ϲ�" then %>checked<% end if %> style="width:20px">�Ϲݰ��
							<input type="radio" name="trade_id" value="�迭��" <% if rs("trade_id") = "�迭��" then %>checked<% end if %> style="width:20px">Kwon��ȸ��
					    </td>
				        <th>�������</th>
				        <td class="left">
							<input type="radio" name="use_sw" value="Y" <% if rs("use_sw") = "Y" then %>checked<% end if %> style="width:20px">���
							<input type="radio" name="use_sw" value="N" <% if rs("use_sw") = "N" then %>checked<% end if %> style="width:20px">�̻��
						</td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="���" onClick="javascript:frmcheck();" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="���" onClick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="trade_code" value="<%=trade_code%>" ID="Hidden1">
				</form>
		</div>
	</body>
</html>

