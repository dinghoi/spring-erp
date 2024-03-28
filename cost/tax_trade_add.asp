<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--include virtual="/include/db_create.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim approve_no, title_line, rsTax, u_type, trade_code
Dim bill_trade_code, group_name

approve_no = f_Request("approve_no")
u_type = f_Request("u_type")
trade_code = f_Request("trade_code")

title_line = "�ŷ�ó ���"

bill_trade_code = ""
group_name = ""

'sales_type = ""
'trade_no1 = ""
'trade_no2 = ""
'trade_no3 = ""
'trade_name = ""
'bill_trade_name = ""
'trade_id = "�Ϲ�"
'trade_owner = ""
'trade_addr = ""
'trade_uptae = ""
'trade_upjong = ""
'trade_tel = ""
'trade_fax = ""
'trade_email = ""
'trade_person = ""
'trade_person_tel = ""
'use_sw = "Y"

Dim trade_no1, trade_no2, trade_no3, trade_name, trade_owner, person_email
Dim trade_id, sales_type, trade_addr, trade_uptae, trade_upjong, trade_tel, trade_fax
Dim person_name, person_grade, person_tel_no, person_memo

'Sql = "select * from tax_bill where approve_no = '"&approve_no&"'"
'Set rs=DbConn.Execute(Sql)
objBuilder.Append "SELECT trade_no, trade_name, trade_owner, bill_id, send_email, receive_email "
objBuilder.Append "FROM tax_bill WHERE approve_no = '"&approve_no&"' "

Set rsTax = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

trade_no1 = Mid(rsTax("trade_no"), 1, 3)
trade_no2 = Mid(rsTax("trade_no"), 4, 2)
trade_no3 = Mid(rsTax("trade_no"), 6)
trade_name = rsTax("trade_name")
trade_owner = rsTax("trade_owner")

If rsTax("bill_id") = "1" Then
	person_email = rsTax("send_email")
Else
	person_email = rsTax("receive_email")
End If

trade_id = ""
sales_type = ""
trade_addr = ""
trade_uptae = ""
trade_upjong = ""
trade_tel = ""
trade_fax = ""
person_name = ""
person_grade = ""
person_tel_no = ""
person_memo = ""

'���� �ּ�
'trade_person = rs("trade_person")
'trade_person_tel = rs("trade_person_tel")
'bill_trade_code = rs("bill_trade_code")
'bill_trade_name = rs("bill_trade_name")

rsTax.Close() : Set rsTax = Nothing

Dim sales_saupbu, cost_year, sqlStr, rsSales

sales_saupbu = "Y"
cost_year = Year(Now())

'sql="select * from sales_org where saupbu = '"&saupbu&"' and sales_year='" & cost_year & "' "
'Set rs=DbConn.Execute(Sql)
sqlStr = "SELECT sort_seq FROM sales_org WHERE saupbu = '"&bonbu&"' AND sales_year = '"&cost_year&"' "
Set rsSales = DBConn.Execute(sqlStr)

If rsSales.EOF Or rsSales.BOF Then
	sales_saupbu = "N"
	'saupbu = ""
	bonbu = ""
End If

rsSales.Close() : Set rsSales = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.trade_no1.value ==""){
					alert('����ڹ�ȣ�� �Է��ϼ���');
					frm.trade_no1.focus();
					return false;
				}

				if(document.frm.trade_no2.value ==""){
					alert('����ڹ�ȣ�� �Է��ϼ���');
					frm.trade_no2.focus();
					return false;
				}

				if(document.frm.trade_no3.value ==""){
					alert('����ڹ�ȣ�� �Է��ϼ���');
					frm.trade_no3.focus();
					return false;
				}

				if(document.frm.trade_name.value ==""){
					alert('��ȣ�� �Է��ϼ���');
					frm.trade_name.focus();
					return false;
				}

				if(document.frm.sales_type.value ==""){
					alert('�ŷ�ó ������ �����ϼ���');
					frm.sales_type.focus();
					return false;
				}

				k = 0;

				for(j=0;j<3;j++){
					if(eval("document.frm.trade_id[" + j + "].checked")){
						k = k + 1
					}
				}

				if(k==0){
					alert ("��೻���� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}

				if(document.frm.trade_owner.value ==""){
					alert('��ǥ�ڸ��� �Է��ϼ���');

					frm.trade_owner.focus();
					return false;
				}
				/*
				if(document.frm.trade_addr.value ==""){
					alert('�ּҸ� �Է��ϼ���');

					frm.trade_addr.focus();
					return false;
				}

				if(document.frm.trade_uptae.value ==""){
					alert('���¸� �Է��ϼ���');

					frm.trade_uptae.focus();
					return false;
				}

				if(document.frm.trade_upjong.value =="") {
					alert('������ �Է��ϼ���');

					frm.trade_upjong.focus();
					return false;
				}
				*/

				if(document.frm.person_email.value !="") {
					if(document.frm.person_name.value =="") {
						alert('��꼭������ �־� ����ڸ� �Է��ؾ� �մϴ�');

						frm.person_name.focus();
						return false;
					}
				}

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
			<form action="/cost/tax_trade_add_save.asp" method="post" name="frm">
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
				        <td class="left">
                        <input name="trade_no1" type="text" id="trade_no1" style="width:25px; text-align:center" maxlength="3" value="<%=trade_no1%>" onKeyUp="checkNum(this);">
                        -
                        <input name="trade_no2" type="text" id="trade_no2" style="width:20px; text-align:center" maxlength="2" value="<%=trade_no2%>" onKeyUp="checkNum(this);">
                        -
                        <input name="trade_no3" type="text" id="trade_no3" style="width:50px; text-align:center" maxlength="5" value="<%=trade_no3%>" onKeyUp="checkNum(this);"></td>
				        <th>��ȣ</th>
				        <td class="left"><input name="trade_name" type="text" id="trade_name" style="width:200px;" value="<%=trade_name%>" onKeyUp="checklength(this,50);"></td>
			          </tr>
				      <tr>
				        <th class="first">�ŷ�ó����</th>
				        <td class="left">
							<select name="sales_type" id="sales_type" style="width:200px">
								<option value="">����</option>
								<option value="����" <%If sales_type = "����" Then %>selected<%End If %>>����</option>
								<option value="����" <%If sales_type = "����" Then %>selected<%End If %>>����</option>
								<option value="����" <%If sales_type = "����" Then %>selected<%End If %>>����</option>
							</select>
						</td>
				        <th>��೻��</th>
				        <td class="left">
							<input type="radio" name="trade_id" value="����" <%If trade_id = "����" Then %>checked<%End If %> style="width:20px">��������
							<input type="radio" name="trade_id" value="�Ϲ�" <%If trade_id = "�Ϲ�" Then %>checked<%End If %> style="width:20px">�Ϲݰ��
							<input type="radio" name="trade_id" value="�迭��" <%If trade_id = "�迭��" Then %>checked<%End If %> style="width:20px">Kwon��ȸ��
						</td>
			          </tr>
				      <tr>
				        <th class="first">�׷��</th>
				        <td class="left">
							<input name="group_name" type="text" id="group_name" style="width:170px;" value="<%=group_name%>" onKeyUp="checklength(this,30);">
							<a href="#" onClick="pop_Window('/trade_search.asp?gubun=<%="5"%>','trade_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">��ȸ</a>
						</td>
				        <th>��ǥ��</th>
				        <td class="left">
							<input name="trade_owner" type="text" id="trade_owner" style="width:200px;" value="<%=trade_owner%>" onKeyUp="checklength(this,20);">
						</td>
			          </tr>
				      <tr>
				        <th class="first">�ּ�</th>
				        <td colspan="3" class="left">
							<input name="trade_addr" type="text" id="trade_addr" style="width:500px" value="<%=trade_addr%>" onKeyUp="checklength(this,100);">
						</td>
			          </tr>
				      <tr>
				        <th class="first">����</th>
				        <td class="left">
							<input name="trade_uptae" type="text" id="trade_uptae" style="width:200px;" value="<%=trade_uptae%>" onKeyUp="checklength(this,50);">
						</td>
				        <th>����</th>
				        <td class="left">
							<input name="trade_upjong" type="text" id="trade_upjong" style="width:200px;" value="<%=trade_upjong%>" onKeyUp="checklength(this,50);">
						</td>
			          </tr>
				      <tr>
				        <th class="first">��ȭ��ȣ</th>
				        <td class="left">
							<input name="trade_tel" type="text" id="trade_tel" style="width:200px;" value="<%=trade_tel%>" onKeyUp="checklength(this,20);">
						</td>
				        <th>�ѽ�</th>
				        <td class="left">
							<input name="trade_fax" type="text" id="trade_fax" style="width:200px;" value="<%=trade_fax%>" onKeyUp="checklength(this,20);">
						</td>
			          </tr>
				      <tr>
				        <th class="first">�����</th>
				        <td class="left">
							<input name="person_name" type="text" id="person_name" style="width:200px;" value="<%=person_name%>" onKeyUp="checklength(this,20);">
						</td>
				        <th>����� ����</th>
				        <td class="left">
							<input name="person_grade" type="text" id="person_grade" style="width:200px;" value="<%=person_grade%>" onKeyUp="checklength(this,20);">
						</td>
			          </tr>
				      <tr>
				        <th class="first">��ȭ��ȣ</th>
				        <td class="left">
							<input name="person_tel_no" type="text" id="person_tel_no" style="width:200px;" value="<%=person_tel_no%>" onKeyUp="checklength(this,20);">
						</td>
				        <th>��꼭����</th>
				        <td class="left">
							<input name="person_email" type="text" id="person_email" style="width:200px;" value="<%=person_email%>" onKeyUp="checklength(this,50);">
						</td>
			          </tr>
				      <tr>
				        <th class="first">�ŷ�ó�޸�</th>
				        <td colspan="3" class="left">
							<input name="person_memo" type="text" id="person_memo" style="width:500px" value="<%=person_memo%>" onKeyUp="checklength(this,50);">
						</td>
			          </tr>
				      <tr>
				        <th class="first">���̿������</th>
				        <td class="left">
							<input name="emp_no" type="text" id="emp_no" style="width:80px;" value="<%=emp_no%>" readonly="true">
							<input name="emp_name" type="text" id="emp_name" style="width:100px;" value="<%=user_name%>" readonly="true">
                        </td>
				        <th>�������</th>
				        <td class="left">
						<%
						If sales_saupbu = "Y" Then
						%>
							<input name="saupbu" type="text" id="saupbu" style="width:150px;" value="<%'=saupbu%><%=bonbu%>" readonly="true">
						<%
						Else
						%>
							<select name="saupbu" id="saupbu" style="width:150px">
								<option value="" <%'If saupbu = "" Then %><%If bonbu = "" Then%>selected<%End If %>>����</option>
						<%
							Dim sqlOrg, rs_org

							sqlOrg = "SELECT saupbu FROM sales_org WHERE sales_year='"&cost_year&"' ORDER BY sort_seq ASC "
							'rs_org.Open Sql, Dbconn, 1
							Set rs_org = DBConn.Execute(sqlOrg)

							Do Until rs_org.EOF
						%>
                    			<option value='<%=rs_org("saupbu")%>' <%'If saupbu = rs_org("saupbu") then %><%If bonbu = rs_org("saupbu") Then%>selected<%End If %>><%=rs_org("saupbu")%></option>
						<%
                    			rs_org.MoveNext()
							Loop
							rs_org.Close() : Set rs_org = Nothing
							DBConn.Close() : Set DBConn = Nothing
                   		%>
                    		</select>
						<%
						End If
						%>
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01"><input type="button" value="���" onClick="javascript:frmcheck();" /></span>
                    <span class="btnType01"><input type="button" value="���" onClick="javascript:goAction();" /></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" />
				<input type="hidden" name="trade_code" value="<%=trade_code%>" />
				<input type="hidden" name="bill_trade_code" value="<%=bill_trade_code%>" />
			</form>
		</div>
	</body>
</html>