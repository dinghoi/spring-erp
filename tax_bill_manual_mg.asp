<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
		slip_gubun = request("slip_gubun")
	else
		bill_month = request.form("bill_month")
		slip_gubun = request.form("slip_gubun")
	end if

	if bill_month = "" then
		bill_month = mid(now(),1,4) + mid(now(),6,2)
		slip_gubun = "��ü"
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

	Set Dbconn=Server.CreateObject("ADODB.Connection")
	Set Rs = Server.CreateObject("ADODB.Recordset")
	Set rs_etc = Server.CreateObject("ADODB.Recordset")
	Set RsCount = Server.CreateObject("ADODB.Recordset")
	Set Rscost = Server.CreateObject("ADODB.Recordset")
	dbconn.open DbConnect

' �����Ǻ�
	posi_sql = " and (emp_no = '"&user_id&"' or reg_id = '"&user_id&"') "

	if position = "����" then
		view_condi = "����"
	end if

	if position = "��Ʈ��" then
		if org_name = "��ȭ����ȣ��" then
			posi_sql = " and (org_name = '��ȭ����ȣ��' or org_name = '��ȭ��������') "
		  else
			posi_sql = " and org_name = '"&org_name&"'"
		end if
	end if

	if position = "����" then
		posi_sql = " and team = '"&team&"'"
	end if

	if position = "�������" or cost_grade = "2" then
		posi_sql = " and saupbu = '"&saupbu&"'"
	end if

	if position = "������" or cost_grade = "1" then
		posi_sql = " and bonbu = '"&bonbu&"'"
	end if

	view_grade = position

	if cost_grade = "0" then
		posi_sql = ""
	end if

	if slip_gubun = "��ü" then
		gubun_sql = ""
	  else
	  	gubun_sql = " and slip_gubun = '"&slip_gubun&"' "
	end if

	base_sql = "select * from general_cost where (tax_bill_yn = 'Y') and (manual_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
	order_sql = " ORDER BY org_name, emp_name, slip_date ASC"

	sql = "select count(*) from general_cost where (tax_bill_yn = 'Y') and (manual_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
	Set RsCount = Dbconn.Execute (sql)

	tottal_record = cint(RsCount(0)) 'Result.RecordCount

	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF

	sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from general_cost where (tax_bill_yn = 'Y') and (manual_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
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

	sql = base_sql + posi_sql + gubun_sql + order_sql + " limit "& stpage & "," &pgsize
	Rs.Open Sql, Dbconn, 1

	title_line = "���۾� ���� ���ݰ�꼭 ����"
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
				<form action="tax_bill_manual_mg.asp" method="post" name="frm">
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
                                <strong>������� : </strong>
                                <select name="slip_gubun" id="slip_gubun" style="width:120px">
                                  <option value='��ü' <%If slip_gubun = "��ü" then %>selected<% end if %>>��ü</option>
                                  <%
                                    Sql="select * from type_code where etc_seq = '4' and etc_id = 'T' order by type_name asc"
                                    rs_etc.Open Sql, Dbconn, 1
                                    do until rs_etc.eof
                                    %>
                                  <option value='<%=rs_etc("type_name")%>' <%If slip_gubun = rs_etc("type_name") then %>selected<% end if %>><%=rs_etc("type_name")%></option>
                                  <%
                                        rs_etc.movenext()
                                    loop
                                    rs_etc.close()
                                    %>
                                  <option value='���' <%If slip_gubun = "���" then %>selected<% end if %>>���</option>
                                </select>
                                </label>
            					<a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="5%" >
							<col width="7%" >
							<col width="11%" >
							<col width="12%" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
							<col width="4%" >
							<col width="7%" >
							<col width="12%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�������</th>
								<th scope="col">��翵�������</th>
								<th scope="col">�����</th>
								<th scope="col">��������</th>
								<th scope="col">����</th>
								<th scope="col">���־�ü</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">����</th>
								<th scope="col">��������</th>
								<th scope="col">���೻��</th>
								<th scope="col">����</th>
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
								<td class="right"><%=formatnumber(sum_price,0)%></td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
								<td class="right"><%=formatnumber(sum_cost_vat,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
						do until rs.eof
							if rs("end_yn") = "Y" then
								end_yn = "����"
								end_view = "N"
							  elseif rs("end_yn") = "I" then
								end_yn = "������"
								end_view = "N"
							  else
							  	end_yn = "����"
							end if
							org_name = rs("emp_company") + "/" + rs("org_name")
							customer_no = mid(rs("customer_no"),1,3) + "-" + mid(rs("customer_no"),4,2) + "-" + mid(rs("customer_no"),6)
						%>
							<tr>
								<td class="first"><%=rs("org_name")%></td>
								<td><%=rs("mg_saupbu")%>&nbsp;</td>
								<td><%=rs("emp_name")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("customer")%></td>
							  	<td class="right"><%=formatnumber(rs("price"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							  	<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td><%=rs("slip_gubun")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
								<td>
							<% if rs("end_yn") = "C" or rs("end_yn") = "N" then %>
							<%   if (rs("reg_id") = user_id) or (rs("emp_no") = user_id) or cost_grade = "0" or position ="�������" or position = "������"  then	%>
                                <a href="#" onClick="pop_Window('tax_bill_manual_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','tax_bill_manual_add_pop','scrollbars=yes,width=1000,height=310')">����</a>
							<%     else	%>
								�Ұ�
                            <%	 end if	%>
							<%  else	%>
								����
                        	<% end if %>
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
                    <a href="tax_bill_manual_excel.asp?bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="tax_bill_manual_mg.asp?page=<%=first_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="tax_bill_manual_mg.asp?page=<%=intstart -1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="tax_bill_manual_mg.asp?page=<%=i%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="tax_bill_manual_mg.asp?page=<%=intend+1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[����]</a>
                        <a href="tax_bill_manual_mg.asp?page=<%=total_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnRight">
					<a href="#" onClick="pop_Window('tax_bill_manual_add.asp','tax_bill_manual_add_pop','scrollbars=yes,width=1000,height=310')" class="btnType04">���� ���ݰ�꼭 ���</a>
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>

