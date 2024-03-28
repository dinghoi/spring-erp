<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim approve_no, title_line, rs, cost_year, rs_org

approve_no = Request("approve_no")

title_line = "���� ����� ����"

objBuilder.Append "SELECT sslt.saupbu, sslt.cost_amt, sslt.vat_amt, sslt.sales_amt, sslt.sales_memo, "
objBuilder.Append "	sslt.sales_date, sslt.sales_company, sslt.emp_name, sslt.emp_no, "
objBuilder.Append "	eomt.org_company, org_name, org_bonbu "
objBuilder.Append "FROM saupbu_sales AS sslt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON sslt.emp_no = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE approve_no = '"&approve_no&"' "

Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

cost_year = Mid(rs("sales_date"), 1, 4)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function goAction(){
			   window.close ();
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
				if(document.frm.saupbu.value ==""){
					alert('�������θ� �����ϼ���');
					frm.saupbu.focus();
					return false;
				}

				{
					a=confirm('���� �Ͻðڽ��ϱ�?')
					if(a==true){
						return true;
					}
					return false;
				}
			}

			//���� ����[����ȣ_20210716]
			function sales_del(grade){
				//���� üũ([memb]account_grade:0 �� ����)
				if(grade !== "0"){
					non_grade();
					return false;
				}

				cfm = confirm("���� �����Ͻðڽ��ϱ�?");

				if(cfm === true){
					sales_del_Init();
					return;
				}
			}

			function sales_del_Init(){
				var frm = document.frm;
				var app_no = $('#approve_no').val();

				frm.action = "/sales/sales_saupbu_del.asp?approve_no="+app_no;
				frm.submit();
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/sales/sales_saupbu_mod_save.asp" method="post" name="frm">
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
				        <th class="first">��������</th>
				        <td class="left"><%=rs("sales_date")%></td>
				        <th>����ȸ��</th>
				        <td class="left">
							<%=rs("sales_company")%>
							<%'=rs("org_company")%>
						</td>
			          </tr>
				      <tr>
				        <th class="first">��������</th>
				        <td class="left">
	                        <select name="saupbu" id="saupbu" style="width:150px">
                                <!--<option value="ȸ�簣�ŷ�" <% if rs("saupbu") ="ȸ�簣�ŷ�" then %>selected<% end if %>>ȸ�簣�ŷ�</option>-->
							<%
							objBuilder.Append "SELECT saupbu FROM sales_org "
							objBuilder.Append "WHERE sales_year = '"&cost_year&"' AND saupbu <> '��Ÿ�����' "
							objBuilder.Append "ORDER BY sort_seq "

							Set rs_org = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rs_org.EOF
							%>
                                <option value='<%=rs_org("saupbu")%>'
									<%If rs("saupbu") = rs_org("saupbu") then %>selected<% end if %>><%=rs_org("saupbu")%>
								</option>
							<%
								rs_org.MoveNext()
							Loop
							rs_org.Close() : Set rs_org = Nothing

							'������ �繫�̻�, �ý��� �����ڸ� ����
							If user_id = "100359" Or user_id = "102592" Then
							%>
							<option value="��Ÿ�����" <%If rs("org_bonbu") = "��Ÿ�����" Then %>selected<%End If %>>��Ÿ�����</option>
							<%End If%>
	                        </select>
                        </td>
				        <th>�����</th>
				        <td class="left">
                        	<input name="emp_name" type="text" id="emp_name" style="width:60px" value="<%=rs("emp_name")%>" readonly="true">
                        	<input name="emp_no" type="text" id="emp_no" style="width:60px" value="<%=rs("emp_no")%>" readonly="true">
                          	<input name="emp_grade" type="hidden" id="emp_grade" style="width:60px" readonly="true">
							<a href="#" onClick="pop_Window('/insa/emp_search.asp?gubun=1','emp_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">�����ȸ</a>
						</td>
			          </tr>
				      <tr>
				        <th class="first">���ް���</th>
				        <td class="left"><%=FormatNumber(rs("cost_amt"), 0)%></td>
				        <th>����</th>
				        <td class="left"><%=FormatNumber(rs("vat_amt"), 0)%></td>
			          </tr>
				      <tr>
				        <th class="first">�հ�ݾ�</th>
				        <td class="left"><%=FormatNumber(rs("sales_amt"), 0)%></td>
				        <th>ǰ���</th>
				        <td class="left"><%=rs("sales_memo")%></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align="center">
                    <span class="btnType01">
						<input type="button" value="����" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1">
					</span>
                    <span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"></span>&nbsp;&nbsp;
				<%
				If sales_grade = "0" Or empProfitViewAll = "Y" Then
				%>
					<span class="btnType01"><input type="button" value="����" onclick="sales_del('<%=account_grade%>');"></span>
				<%
				End If
				%>
                </div>
				<input type="hidden" name="sales_date" value="<%=rs("sales_date")%>" />
				<input type="hidden" name="approve_no" value="<%=approve_no%>" id="approve_no" />
			</form>
			<%
			rs.Close() : Set rs = Nothing
			DBConn.Close() : Set DBConn = Nothing
			%>
		</div>
	</body>
</html>

