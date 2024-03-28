<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim win_sw, ck_sw
Dim slip_month, owner_company, card_type, field_check, field_view
Dim from_date, end_date, to_date
Dim Page, pgsize, start_page, stpage
Dim owner_company_sql, card_type_sql, field_sql, order_sql, sDate_sql
Dim rs_vat, vat_record
Dim rsCount, total_record, total_page
Dim rs_sum, sum_price, sum_cost, sum_cost_vat
Dim rsCard, rs_etc
Dim title_line, del_msg

Dim err_msg

win_sw = "close"

ck_sw = Request("ck_sw")
Page = Request("page")

If ck_sw = "y" Then
	slip_month = Request("slip_month")
	owner_company = Request("owner_company")
	card_type = Request("card_type")
	field_check = Request("field_check")
	field_view = Request("field_view")
Else
	slip_month = Request.Form("slip_month")
	owner_company = Request.Form("owner_company")
	card_type = Request.Form("card_type")
	field_check = Request.Form("field_check")
	field_view = Request.Form("field_view")
End if

If slip_month = "" Then
	slip_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
	field_check = "total"
	card_type = "��ü"
	owner_company = "��ü"
End If

If field_check = "total" Then
	field_view = ""
End If

from_date = Mid(slip_month, 1, 4) & "-" + Mid(slip_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

pgsize = 10 ' ȭ�� �� ������

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

sDate_sql = " (slip_date >='"&from_date&"' AND slip_date <= '"&to_date&"') "

'�˻� ��ȸ ����
If owner_company = "��ü" Then
	owner_company_sql = " "
Else
	owner_company_sql = " AND (owner_company = '" + owner_company + "') "
End If

If card_type = "��ü" Then
	card_type_sql = " "
Else
	card_type_sql = " AND (card_slip.card_type = '" + card_type + "') "
End If

If field_check <> "total" Then
	If field_check = "person_end" Then
		field_sql = " AND (card_slip." + field_check + " = 'N') "
	Else
		field_sql = " AND (card_slip." + field_check + " LIKE '%" + field_view + "%') "
	End If
Else
  	field_sql = " "
End If

order_sql = " ORDER BY slip_date ASC"

'�ΰ���
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE (cost_vat > 0) "
objBuilder.Append "	AND " & sDate_sql
objBuilder.Append owner_company_sql & card_type_sql & field_sql

Set rs_vat = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

vat_record = CInt(rs_vat(0)) 'Result.RecordCount

rs_vat.Close()
Set rs_vat = Nothing

'�� �Ǽ�
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE " & sDate_sql
objBuilder.Append owner_company_sql & card_type_sql & field_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'Result.RecordCount
total_record = CInt(RsCount(0))

rsCount.Close()
Set rsCount = Nothing

'Result.PageCount
If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize)
Else
	total_page = Int((total_record / pgsize) + 1)
End If

'�Ѱ�
objBuilder.Append "SELECT SUM(price) AS price, "
objBuilder.Append "	SUM(cost) AS cost, "
objBuilder.Append "	SUM(cost_vat) AS cost_vat "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE " & sDate_sql
objBuilder.Append owner_company_sql & card_type_sql & field_sql

Set rs_sum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_sum("price")) Then
	sum_price = 0
	sum_cost = 0
	sum_cost_vat = 0
Else
	sum_price = CDbl(rs_sum("price"))
	sum_cost = CDbl(rs_sum("cost"))
	sum_cost_vat = CDbl(rs_sum("cost_vat"))
End If

objBuilder.Append "SELECT approve_no, cancel_yn, slip_date, card_type, card_no, emp_name, "
objBuilder.Append "customer, upjong, account, account_item, price, "
objBuilder.Append "cost, cost_vat, account_end, person_end, end_sw, "
objBuilder.Append "pl_yn "
objBuilder.Append "FROM card_slip "
objBuilder.Append "WHERE " & sDate_sql
objBuilder.Append owner_company_sql & card_type_sql & field_sql & order_sql & " LIMIT " & stpage & "," & pgsize

Set rsCard = Server.CreateObject("ADODB.Recordset")
rsCard.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

title_line = "ī�� ��ǥ ����"

'del_msg = slip_month + "�� ī������ " + card_type + " �����Ͻðڽ��ϱ�??"
del_msg = Left(slip_month, 4) & "�� " & Right(slip_month, 2) & "�� ����ȸ�� " & owner_company & "�� ī������ " + card_type + "��(��) �����Ͻðڽ��ϱ�?"
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
				return "0 1";
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
				if(chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if(document.frm.slip_month.value == ""){
					alert ("������� �Է��ϼ���");
					return false;
				}
				return true;
			}

			function del_confirm(del_msg){
				result = confirm(del_msg);

				if(result == true){
					document.frm.action = "card_slip_del.asp";
               		document.frm.submit();
					return true;
				}
				return false;
			}

			function first_end(){
				a = confirm('�����Ͻðڽ��ϱ�?');

				if(a == true){
					document.frm.action = "card_first_end.asp";
               		document.frm.submit();
					return true;
				}

				return false;
			}

			function first_end_cancel(){
				a = confirm('����Ͻðڽ��ϱ�?');

				if (a == true) {
					document.frm.action = "card_first_end_cancel.asp";
               		document.frm.submit();
					return true;
				}
				return false;
			}

			function last_end(){
				a = confirm('�����Ͻðڽ��ϱ�?');

				if(a == true) {
					document.frm.action = "card_last_end.asp";
               		document.frm.submit();
					return true;
				}
				return false;
			}

			function last_end_cancel(){
				a = confirm('����Ͻðڽ��ϱ�?');

				if (a == true) {
					document.frm.action = "card_last_end_cancel.asp";
               		document.frm.submit();
					return true;
				}
				return false;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/card_slip_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="card_slip_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
							<p>
                                <label>
                                &nbsp;&nbsp;<strong>�����&nbsp;</strong>(��201401) :
                                <input name="slip_month" type="text" value="<%=slip_month%>" style="width:60px">
                                </label>
                                <strong>����ȸ��</strong>
                                <select name="owner_company" id="owner_company" style="width:120px">
                                    <option value="��ü" <% if owner_company = "��ü" then %>selected<% end if %>>��ü</option>
                                    <%

                                    ' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
									objBuilder.Append "SELECT org_company "
									objBuilder.Append "FROM emp_org_mst "
									objBuilder.Append "WHERE ISNULL(org_end_date) "
									objBuilder.Append "AND org_level = 'ȸ��' "
									objBuilder.Append "ORDER BY org_company ASC "

									Set rs_etc = Server.CreateObject("ADODB.Recordset")
                                    rs_etc.Open objBuilder.ToString(), DBConn, 1
									objBuilder.Clear()

                                    Do Until rs_etc.EOF
                                        %>
                                        <option value='<%=rs_etc("org_company")%>' <%If owner_company = rs_etc("org_company") then %>selected<% end if %>><%=rs_etc("org_company")%></option>
                                        <%
                                        rs_etc.MoveNext()
                                    Loop

                                    rs_etc.close()
                                    %>
                                </select>
                                <strong>ī������</strong>
                                <select name="card_type" id="card_type" style="width:100px">
                                    <option value="��ü" <% if card_type = "��ü" then %>selected<% end if %>>��ü</option>
                                    <%
									objBuilder.Append "SELECT etc_name "
									objBuilder.Append "FROM etc_code "
									objBuilder.Append "WHERE etc_type = '44' "
									objBuilder.Append "ORDER BY etc_name ASC "

                                    rs_etc.Open objBuilder.ToString(), DBConn, 1
									objBuilder.Clear()

                                    Do Until rs_etc.EOF
                                        %>
                                        <option value='<%=rs_etc("etc_name")%>' <%If card_type = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                        <%
                                        rs_etc.MoveNext()
                                    Loop

                                    rs_etc.close()
									Set rs_etc = Nothing
                                    %>
                                </select>
                                <strong>�ʵ�����</strong>
                                <select name="field_check" id="field_check" style="width:120px">
                                    <option value="total" <%If field_check = "total" Then %>selected<%End If %>>��ü</option>
                                    <option value="card_no" <%If field_check = "card_no" Then %>selected<%End If %>>ī���ȣ</option>
                                    <option value="emp_name" <%If field_check = "emp_name" Then %>selected<%End If %>>�����</option>
                                    <option value="customer" <%If field_check = "customer" Then %>selected<%End If %>>�ŷ�ó</option>
                                    <option value="account" <%If field_check = "account" Then %>selected<%End If %>>��������</option>
                                    <option value="account_item" <%If field_check = "account_item" Then %>selected<%End If %>>�׸�</option>
                                    <option value="person_end" <%If field_check = "person_end" Then %>selected<%End If %>>���θ����ȵ�����</option>
                                    <option value="upjong" <%If field_check = "upjong" Then %>selected<%End If %>>����</option>
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
							<col width="6%" >
							<col width="7%" >
							<col width="10%" >
							<col width="5%" >
							<col width="*" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="5%" >
							<col width="8%" >
							<col width="8%" >
							<col width="3%" >
							<col width="4%" >
							<col width="4%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">�����</th>
								<th rowspan="2" scope="col">ī������</th>
								<th rowspan="2" scope="col">ī���ȣ</th>
								<th rowspan="2" scope="col">�����</th>
								<th rowspan="2" scope="col">�ŷ�ó</th>
								<th rowspan="2" scope="col">����</th>
								<th rowspan="2" scope="col">�հ�</th>
								<th rowspan="2" scope="col">���ް���</th>
								<th rowspan="2" scope="col">�ΰ���</th>
								<th rowspan="2" scope="col">��������</th>
								<th rowspan="2" scope="col">�׸�</th>
								<th rowspan="2" scope="col">����</th>
								<th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">���� ����</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">�渮</th>
							  <th scope="col">�����</th>
							  <th scope="col">����</th>
				            </tr>
						</thead>
						<tbody>
							<tr>
                                <th colspan="2" class="first">�Ѱ�</th>
                                <th><%=total_record%>&nbsp;��</th>
                                <th colspan="3"><%=err_msg%>&nbsp;�հ� :&nbsp;<%=FormatNumber(sum_price,0)%></th>
                                <th colspan="4">���ް��� :&nbsp;<%=FormatNumber(sum_cost,0)%></th>
                                <th colspan="5">�ΰ��� :&nbsp;<%=FormatNumber(sum_cost_vat,0)%>&nbsp;(<%=vat_record%>��)</th>
							</tr>
						<%
						Dim i, j
						Dim account_end, person_end, end_sw, end_count, account_view
						Dim person_view

						i = 0
						j = 0
						account_end = ""
						person_end = ""
						end_sw = ""
						end_count = 0

						Do Until rsCard.EOF
							account_end = rsCard("account_end")
							person_end = rsCard("person_end")
							end_sw = rsCard("end_sw")

							If end_sw = "Y" Then
								end_count = end_count + 1
							End If

							i = i + 1
							If rsCard("cost_vat") <> 0 Then
								j = j + 1
							End If

							If rsCard("account_end") = "Y" Then
								account_view = "����"
							Else
							  	account_view = "����"
							End If

							If rsCard("person_end") = "Y" Then
								person_view = "����"
							Else
							  	person_view = "����"
							End If
						%>
							<tr>
                                <td class="first"><%=rsCard("slip_date")%><input type="hidden" name="approve_no" value="<%=rsCard("approve_no")%>"></td>
                                <td><%=rsCard("card_type")%></td>
                                <td><%=rsCard("card_no")%></td>
                                <td><%=rsCard("emp_name")%></td>
                                <td class="left"><a href="#" onClick="pop_Window('card_customer_mod.asp?approve_no=<%=rsCard("approve_no")%>','ī��ŷ�ó����','scrollbars=yes,width=700,height=200')"><%=rsCard("customer")%></a></td>
                                <td class="left"><%=rsCard("upjong")%></td>
                                <td class="right"><%=FormatNumber(rsCard("price"),0)%></td>
                                <td class="right"><%=FormatNumber(rsCard("cost"),0)%></td>
                                <td class="right"><%=FormatNumber(rsCard("cost_vat"),0)%></td>
                                <td><%=rsCard("account")%>&nbsp;</td>
                                <td><%=rsCard("account_item")%>&nbsp;</td>
                                <td><%=rsCard("pl_yn")%></td>
                                <td><%=account_view%></td>
                                <td><%=person_view%></td>
                                <td>
                                <% If rsCard("end_sw") = "Y" Then %>
                                    ����
                                <% Else %>
                                    <a href="#" onClick="pop_Window('card_slip_mod.asp?approve_no=<%=rsCard("approve_no")%>&cancel_yn=<%=rsCard("cancel_yn")%>','ī����ǥ����','scrollbars=yes,width=800,height=300')">����</a>
                                <% End If %>
                                </td>
							</tr>
						<%
							rsCard.MoveNext()
						Loop

						rsCard.close()
						Set rsCard = Nothing

						Dim price_sum, cost_sum, cost_vat_sum

						If price_sum <> ( cost_sum + cost_vat_sum ) Then
							err_msg = "�ݾ�Ȯ�� ���"
						Else
						  	err_msg = " "
						End If
						%>
						</tbody>
					</table>
				</div>
				<%
				Dim intstart, intend, first_page

                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="24%">
					<div class="btnCenter">
					<a href = "card_slip_up_excel.asp?slip_month=<%=slip_month%>&owner_company=<%=owner_company%>&card_type=<%=card_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">���Դٿ�</a>
					<a href = "card_slip_excel.asp?slip_month=<%=slip_month%>&owner_company=<%=owner_company%>&card_type=<%=card_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">�����ٿ�</a>
					<% If end_count = 0 Then %>
						<a href="#" onClick="del_confirm('<%=del_msg%>')" class="btnType04">����</a>
					<% End If %>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="card_slip_mg.asp?page=<%=first_page%>&slip_month=<%=slip_month%>&owner_company=<%=owner_company%>&card_type=<%=card_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[ó��]</a>
                  		<% If intstart > 1 Then %>
                            <a href="card_slip_mg.asp?page=<%=intstart -1%>&slip_month=<%=slip_month%>&owner_company=<%=owner_company%>&card_type=<%=card_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a>
                        <% End If %>
                        <% For i = intstart To intend %>
                            <% If i = int(page) Then %>
                                <b>[<%=i%>]</b>
                            <% Else %>
                                <a href="card_slip_mg.asp?page=<%=i%>&slip_month=<%=slip_month%>&owner_company=<%=owner_company%>&card_type=<%=card_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% End If %>
                        <% Next %>
						<% If intend < total_page Then %>
                            <a href="card_slip_mg.asp?page=<%=intend+1%>&slip_month=<%=slip_month%>&owner_company=<%=owner_company%>&card_type=<%=card_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a>
                            <a href="card_slip_mg.asp?page=<%=total_page%>&slip_month=<%=slip_month%>&owner_company=<%=owner_company%>&card_type=<%=card_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[������]</a>
                        <%	Else %>
                            [����]&nbsp;[������]
                        <% End If %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnCenter">
                    <% If account_end = "N" And end_sw = "N" Then	%>
                        <a href="#" onClick="first_end()" class="btnType04">�渮1������</a>
                    <% End If	%>
                    <% If account_end = "Y" And end_sw = "N" Then	%>
                        <a href="#" onClick="first_end_cancel()" class="btnType04">�渮1�����</a>
                    <% End If	%>
                    <% If account_end = "Y" And end_sw = "N" Then	%>
                        <a href="#" onClick="last_end()" class="btnType04">��������</a>
                    <% End If	%>
                    <% If account_end = "Y" And end_sw = "Y" Then	%>
                        <a href="#" onClick="last_end_cancel()" class="btnType04">�����������</a>
                    <% End If	%>
					</div>
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
