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
Dim ck_sw, page
Dim from_date, to_date, field_check, field_view, page_cnt
Dim curr_dd, bill_id
Dim pgsize, start_page, stpage
Dim rs, title_line

'ck_sw = Request("ck_sw")
page = Request.QueryString("page")

'If ck_sw = "y" Then
'	from_date = Request("from_date")
'	to_date = Request("to_date")
'	field_check = Request("field_check")
'	field_view = Request("field_view")
'	page_cnt = Request("page_cnt")
'Else
'	from_date = Request.Form("from_date")
'	to_date = Request.Form("to_date")
'	field_check = Request.Form("field_check")
'	field_view = Request.Form("field_view")
'	page_cnt = Request.Form("page_cnt")
'End If

from_date = f_Request("from_date")
to_date = f_Request("to_date")
field_check = f_Request("field_check")
field_view = f_Request("field_view")
page_cnt = f_Request("page_cnt")

If to_date = "" Or from_date = "" Then
	curr_dd = CStr(DatePart("d", Now()))
	to_date = Mid(CStr(Now()), 1, 10)
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	field_check = "total"
End If

If field_check = "total" Then
	field_view = ""
End If

bill_id = "����"

pgsize = 10 ' ȭ�� �� ������
'pgsize = page_cnt ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

objBuilder.Append "SELECT glct.customer, glct.customer_no, glct.slip_date, glct.slip_gubun, "
objBuilder.Append "	glct.account_item, glct.slip_memo, glct.price, glct.cost, glct.cost_vat, "
objbuilder.Append "	glct.emp_company, glct.org_name, glct.emp_no, glct.emp_name, glct.end_yn, "
objBuilder.Append "	eomt.org_company, eomt.org_name AS emp_org_name "
objBuilder.Append "FROM general_cost AS glct "
objBuilder.Append "INNER JOIN emp_master AS emtt ON glct.emp_no = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE tax_bill_yn = 'Y' "
objBuilder.Append "	AND (slip_date >= '" & from_date & "' AND slip_date <= '" & to_date  & "') "

' ���Ǻ� ��ȸ
If field_check <> "total" Then
	objBuilder.Append "AND (" & field_check & " LIKE '%" & field_view & "%' ) "
End If

objBuilder.Append "ORDER BY customer, slip_gubun, slip_date "

'Set rs = Server.CreateObject("ADODB.Recordset")
'rs.Open objBuilder.ToString(), DBConn, 1
Set rs = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "���Լ��ݰ�꼭 ����"
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
			function getPageCode(){
				return "1 1";
			}

			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck(){
				if (formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
//				if (document.frm.bill_id.value == "") {
//					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
//					return false;
//				}
				return true;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/tax_bill_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/finance/tax_bill_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>������ : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>������ : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<strong>������ȸ</strong>
                                <select name="field_check" id="field_check" style="width:120px">
                              		<option value="total" <%If field_check = "total" Then%>selected<%End If%>>��ü</option>
                                    <option value="customer" <%If field_check = "customer" Then%>selected<%End If%>>�ŷ�ó</option>
                                    <option value="account" <%If field_check = "account" Then%>selected<%End If%>>����</option>
                                    <option value="eomt.org_name" <%If field_check = "eomt.org_name" Then%>selected<%End If%>>����μ�</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:120px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="12%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
							<col width="*" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�ŷ�ó</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">������</th>
								<th scope="col">����</th>
								<th scope="col">�׸�</th>
								<th scope="col">���೻��</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">����ȸ��</th>
                                <th scope="col">����μ�</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim i, price_sum, cost_sum, cost_vat_sum, end_yn
						Dim customer_no

						i = 0
						price_sum = 0
						cost_sum = 0
						cost_vat_sum = 0

						Do Until rs.EOF
							price_sum = price_sum + rs("price")
							cost_sum = cost_sum + rs("cost")
							cost_vat_sum = cost_vat_sum + rs("cost_vat")
							i = i + 1

							If rs("end_yn") = "Y" Then
								end_yn = "����"
							Else
							  	end_yn = "����"
							End If

							customer_no = Mid(rs("customer_no"), 1, 3) & "-" & Mid(rs("customer_no"), 4, 2) & "-" & Right(rs("customer_no"), 5)
						%>
							<tr>
								<td class="first"><%=rs("customer")%></td>
								<td><%=customer_no%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account_item")%></td>
								<td><%=rs("slip_memo")%></td>
							  	<td class="right"><%=FormatNumber(rs("price"), 0)%></td>
							  	<td class="right"><%=FormatNumber(rs("cost"), 0)%></td>
							  	<td class="right"><%=FormatNumber(rs("cost_vat"), 0)%></td>
								<td><%=rs("emp_company")%></td>
								<td><%=rs("emp_org_name")%></td>
							</tr>
						<%
							rs.MoveNext()
						Loop

						rs.Close() : Set rs = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
							<tr>
								<th class="first">�Ѱ�</th>
								<th colspan="1"><%=i%>&nbsp;��</th>
							  	<th colspan="3">�հ� :&nbsp;<%=FormatNumber(price_sum, 0)%></th>
							  	<th colspan="3">���ް��� :&nbsp;<%=FormatNumber(cost_sum, 0)%></th>
								<th colspan="3">�ΰ��� :&nbsp;<%=FormatNumber(cost_vat_sum, 0)%></th>
							</tr>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
                    <a href="/finance/tax_bill_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&bill_id=<%=bill_id%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="user_id">
				<input type="hidden" name="pass">
			</form>
		</div>
	</div>
	</body>
</html>

