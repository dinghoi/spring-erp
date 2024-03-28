<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
	Dim Rs
	Dim Repeat_Rows
	Dim from_date
	Dim to_date
	Dim win_sw

	win_sw = "close"

	ck_sw=Request("ck_sw")
	Page=Request("page")

	if ck_sw = "y" Then
		bill_month = request("bill_month")
		owner_company = request("owner_company")
		field_check = request("field_check")
		field_view = request("field_view")
	else
		bill_month = request.form("bill_month")
		owner_company = request.form("owner_company")
		field_check = request.form("field_check")
		field_view = request.form("field_view")
	end if

	if bill_month = "" then
		bill_month = mid(now(),1,4) + mid(now(),6,2)
		owner_company = "��ü"
		field_check = "total"
		field_view = ""
	end if

	if field_check = "total" then
		field_view = ""
	end if

	from_date = mid(bill_month,1,4) + "-" + mid(bill_month,5,2) + "-01"
	end_date = datevalue(from_date)
	end_date = dateadd("m",1,from_date)
	to_date = cstr(dateadd("d",-1,end_date))

	pgsize = 10 ' ȭ�� �� ������

	If Page = "" Then
		Page = 1
		start_page = 1
	End If
	stpage = int((page - 1) * pgsize)

	base_sql = "select * from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (end_yn = 'Y') and (cost_reg_yn = 'N') and (bill_id ='1') "

	if field_check = "total" then
		field_sql = " "
	  else
		field_sql = " and ("&field_check&" like '%"&field_view&"%') "
	end if
	if owner_company = "��ü" then
		owner_sql = " "
	  else
		owner_sql = " and (owner_company = '"&owner_company&"') "
	end if

	order_sql = " ORDER BY bill_date ASC"

	sql = "select count(*) from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (end_yn = 'Y') and (cost_reg_yn = 'N') and (bill_id = '1') " + field_sql + owner_sql
	Set RsCount = Dbconn.Execute (sql)

	tottal_record = cint(RsCount(0)) 'Result.RecordCount

	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF

	sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from tax_bill where (bill_date >='"&from_date&"' and bill_date <='"&to_date&"') and (end_yn = 'Y') and (cost_reg_yn = 'N') and (bill_id = '1') " + field_sql + owner_sql
	Set rs_sum = Dbconn.Execute (sql)
	if isnull(rs_sum("price")) then
		sum_price = 0
		sum_cost = 0
		sum_cost_vat = 0
	  else
		sum_price = cdbl(rs_sum("price"))
		sum_cost = cdbl(rs_sum("cost"))
		sum_cost_vat = cdbl(rs_sum("cost_vat"))
	end if

	sql = base_sql & field_sql & owner_sql & order_sql & " limit "& stpage & "," &pgsize
	Rs.Open Sql, Dbconn, 1
'Response.write Sql

	title_line = "�̼��� ���� ���ݰ�꼭 ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	    <script src="/java/jquery-1.9.1.js"></script>
	    <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.bill_month.value == "") {
					alert ("����� �����ϼ���");
					return false;
				}
				return true;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_esero_in_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>��ȸ����</dt>
                        <dd>
                            <p>
								<label>
								<strong>��꼭 ������ : </strong>
                                	<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
								<strong>ȸ��</strong>
                                <select name="owner_company" id="owner_company" style="width:150px">
                                  <option value="��ü" <% if owner_company = "��ü" then %>selected<% end if %>>��ü</option>
                                  <%
									' 2019.02.22 ������ ��û ȸ�縮��Ʈ�� ������ �ҽ� org_end_date�� null �� �ƴ� �������ڸ� �����ϸ� ����Ʈ�� ��Ÿ���� �ʴ´�.
									'Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = 'ȸ��'  ORDER BY org_company ASC"
									sql = "SELECT org_name from emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') AND org_level = 'ȸ��' ORDER BY org_company ASC"
                                    rs_org.Open Sql, Dbconn, 1
                                    do until rs_org.eof
                                    %>
                                  <option value='<%=rs_org("org_name")%>' <%If owner_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                                  <%
                                        rs_org.movenext()
                                    loop
                                    rs_org.close()
                                    %>
                                </select>
                                </label>
                                <label>
														<strong>��������</strong>
                                <select name="field_check" id="field_check" style="width:100px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                    <option value="trade_name" <% if field_check = "trade_name" then %>selected<% end if %>>��ȣ��</option>
                                    <option value="tax_bill_memo" <% if field_check = "tax_bill_memo" then %>selected<% end if %>>�ŷ�����</option>
                                    <option value="receive_email" <% if field_check = "receive_email" then %>selected<% end if %>>�̸���</option>                                </select>
								</label>
                                <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px" id="field_view" >
								</label>
            					<a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="8%" >
							<col width="6%" >
							<col width="13%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="*" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">������</th>
								<th scope="col">��꼭����ȸ��</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">��ȣ��</th>
								<th scope="col">�����</th>
								<th scope="col">�����</th>
								<th scope="col">���޹޴����̸���</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">û��</th>
								<th scope="col">�ŷ�����</th>
								<th scope="col">��ϱ���</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>�Ǽ�</strong></td>
								<td><%=formatnumber(tottal_record,0)%>&nbsp;��</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right">&nbsp;</td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
								<td class="right"><%=formatnumber(sum_cost_vat,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
						end_sw = "N"
						do until rs.eof
							Sql="select * from trade where trade_no = '"&rs("trade_no")&"'"
							Set rs_trade=DbConn.Execute(Sql)
							'Response.write Sql
							trade_sw = "Y"
							if rs_trade.eof or rs_trade.bof then
								trade_sw = "N"
							end if
							if rs("receive_email") = "" or isnull(rs("receive_email")) then
								emp_name = "-"
								emp_saupbu = "-"
							  else
								k = instr(1,rs("receive_email"),"@")
								if k < 2 or isnull(k) then
									k = 2
								end if
								Sql="select * from emp_master where emp_email = '"&mid(trim(rs("receive_email")),1,k-1)&"'"
								Set rs_emp=DbConn.Execute(Sql)
								if rs_emp.eof then
									emp_name = "-"
									emp_saupbu = "-"
								  else
									emp_name = rs_emp("emp_name")
									emp_saupbu = rs_emp("emp_saupbu")
								end if
							end if
						%>
							<tr>
								<td class="first"><%=rs("bill_date")%></td>
								<td><%=rs("owner_company")%></td>
								<td><%=mid(rs("trade_no"),1,3)%>-<%=mid(rs("trade_no"),4,2)%>-<%=right(rs("trade_no"),5)%></td>
								<td><%=rs("trade_name")%></td>
								<td><%=emp_saupbu%>&nbsp;</td>
								<td><%=emp_name%></td>
								<td><%=rs("receive_email")%>&nbsp;</td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("bill_collect")%>&nbsp;</td>
								<td class="left"><%=rs("tax_bill_memo")%></td>
								<td>
 						<% if trade_sw = "Y" then	%>
							<a href="#" onClick="pop_Window('tax_esero_in_detail_add.asp?approve_no=<%=rs("approve_no")%>','tax_esero_in_detail_add_pop','scrollbars=yes,width=1000,height=280')">�����</a>
						<%   else	%>
							<a href="#" onClick="pop_Window('tax_trade_add.asp?approve_no=<%=rs("approve_no")%>','tax_trade_add_pop','scrollbars=yes,width=800,height=450')">�ŷ�ó���</a>
                        <% end if	%>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="24%">
					<div class="btnCenter">
                    <a href="tax_esero_in_excel.asp?bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="tax_esero_in_mg.asp?page=<%=first_page%>&bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="tax_esero_in_mg.asp?page=<%=intstart -1%>&bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="tax_esero_in_mg.asp?page=<%=i%>&bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="tax_esero_in_mg.asp?page=<%=intend+1%>&bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a>
                        <a href="tax_esero_in_mg.asp?page=<%=total_page%>&bill_month=<%=bill_month%>&owner_company=<%=owner_company%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnCenter">
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>

