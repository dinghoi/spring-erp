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
        view_c = request("view_c")
        view_d = request("view_d")
		emp_name = request("emp_name")
	else
		bill_month = request.form("bill_month")
		slip_gubun = request.form("slip_gubun")
        view_c = request.form("view_c")
        view_d = request.form("view_d")
		emp_name = request.form("emp_name")
    end if

    if view_d = "" then
        view_d = "slip"
	end if

	if bill_month = "" then
		bill_month = mid(now(),1,4) + mid(now(),6,2)
		slip_gubun = "��ü"
        view_c = "total"
        view_d = "slip"
		emp_name = ""
	end if

	if view_c = "total" then
		emp_name = ""
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

	if view_c = "total" then
		emp_sql = ""
	elseif view_c = "emp_name" then
	  	emp_sql = " and emp_name like '%"&emp_name&"%'"
	else
	  	emp_sql = " and customer like '%"&emp_name&"%'"
	end if

    base_sql = "select * from general_cost where (tax_bill_yn = 'Y') "
    if view_d = "slip" then
        base_sql = base_sql & " and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
        order_sql = " ORDER BY org_name, emp_name, slip_date ASC"
    end if
    if view_d = "reg" then
        base_sql = base_sql & " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
        order_sql = " ORDER BY org_name, emp_name, reg_date ASC"
    end if

    sql = "select count(*) from general_cost where (tax_bill_yn = 'Y') "
    if view_d = "slip" then
        sql = sql & " and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')  "
    end if
    if view_d = "reg" then
        sql = sql &  " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
    end if
    sql = sql + posi_sql + gubun_sql + emp_sql
	Set RsCount = Dbconn.Execute (sql)

	tottal_record = cint(RsCount(0)) 'Result.RecordCount

	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF

    sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from general_cost where (tax_bill_yn = 'Y') "
    if view_d = "slip" then
        sql = sql & "and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
    end if
    if view_d = "reg" then
        sql = sql &  " and (reg_date >='"&from_date&" 00:00:00' and reg_date <='"&to_date&" 23:59:59')"
    end if

    sql = sql +  posi_sql + gubun_sql + emp_sql

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

	sql = base_sql + posi_sql + gubun_sql + emp_sql + order_sql + " limit "& stpage & "," &pgsize
	Rs.Open Sql, Dbconn, 1

	title_line = "���� ���ݰ�꼭 ����"
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
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('emp_name_view').style.display = 'none';
				}
				if (eval("document.frm.view_c[1].checked") || eval("document.frm.view_c[2].checked")) {
					document.getElementById('emp_name_view').style.display = '';
				}
			}
		</script>
	</head>
	<body onLoad="condi_view()">
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="tax_bill_in_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>��ȸ����</dt>
                        <dd>
                            <p>
								<label>
                                    <input type="radio" name="view_d" value="slip" <% if view_d = "slip" then %>checked<% end if %> style="width:25px">
                                    <strong>�߻����&nbsp;</strong>
                                    <input type="radio" name="view_d" value="reg" <% if view_d = "reg" then %>checked<% end if %> style="width:25px">
                                    <strong>�߱޳��&nbsp;</strong>

                                    : <input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
                                    (��201401)
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
								<label>
								<strong>��ȸ���� : </strong>
                              	<input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px" onClick="condi_view()">��ü
                                <input type="radio" name="view_c" value="emp_name" <% if view_c = "emp_name" then %>checked<% end if %> style="width:25px" onClick="condi_view()">���κ�
                                <input type="radio" name="view_c" value="customer" <% if view_c = "customer" then %>checked<% end if %> style="width:25px" onClick="condi_view()">���־�ü
								</label>
								<label>
                                	<input name="emp_name" type="text" value="<%=emp_name%>" style="width:100px; display:none" id="emp_name_view">
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
							<col width="7%" >
							<col width="4%" >
							<col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="11%" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
							<col width="4%" >
							<col width="7%" >
							<col width="11%" >
							<col width="3%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�������</th>
								<th scope="col">��翵��<br>�����</th>
								<th scope="col">�����</th>
								<th scope="col">��������</th>
								<th scope="col">�߱�����</th>
								<th scope="col">����</th>
								<th scope="col">���־�ü</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">����</th>
								<th scope="col">��������</th>
								<th scope="col">���೻��</th>
								<th scope="col">����</th>
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
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(sum_price,0)%></td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
								<td class="right"><%=formatnumber(sum_cost_vat,0)%></td>
								<td>&nbsp;</td>
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
                                <%
                                ' 5�� ���� ���� �Է°� ����...
                                chk_slip_month = mid(rs("slip_date"),1,7)
                                chk_reg_month = mid(rs("reg_date"),1,7)
                                chk_reg_day = mid(rs("reg_date"),9,2)

                                if ((chk_slip_month < chk_reg_month) and (chk_reg_day > "05")) then
                                    bgcolor = "burlywood"
                                else
                                    bgcolor = "#f8f8f8"
                                end if
                                %>
                                <tr style="background-color: <%=bgcolor%>;">
                                    <td class="first"><%=rs("org_name")%></td>
                                    <td><%=rs("mg_saupbu")%>&nbsp;</td>
                                    <td><%=rs("emp_name")%></td>
                                    <td><%=rs("slip_date")%></td>
                                    <td><%=mid(rs("reg_date"),1,10)%></td>
                                    <td><%=rs("company")%></td>
                                    <td><%=rs("customer")%></td>
                                    <td class="right"><%=formatnumber(rs("price"),0)%></td>
                                    <td class="right"><%=formatnumber(rs("cost"),0)%></td>
                                    <td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
                                    <td><%=rs("slip_gubun")%></td>
                                    <td><%=rs("account")%></td>
                                    <td><%=rs("slip_memo")%></td>
                                    <td><%=rs("pl_yn")%></td>
                                    <td>
                                    <% if rs("end_yn") = "C" or rs("end_yn") = "N" then %>
                                        <% if (rs("reg_id") = user_id) or (rs("emp_no") = user_id) or cost_grade = "0" or position ="�������" or position = "������"  then	%>
                                            <a href="#" onClick="pop_Window('tax_bill_in_mod.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','tax_bill_in_mod_pop','scrollbars=yes,width=1000,height=300')">����</a>
                                        <% else	%>
                                            �Ұ�
                                        <% end if %>
                                    <% else	%>
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
                    <a href="tax_bill_in_excel.asp?bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="tax_bill_in_mg.asp?page=<%=first_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[ó��]</a>
                        <% if intstart > 1 then %>
                            <a href="tax_bill_in_mg.asp?page=<%=intstart -1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[����]</a>
                        <% end if %>
                        <% for i = intstart to intend %>
                            <% if i = int(page) then %>
                                <b>[<%=i%>]</b>
                            <% else %>
                                <a href="tax_bill_in_mg.asp?page=<%=i%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% end if %>
                        <% next %>
                        <% if intend < total_page then %>
                            <a href="tax_bill_in_mg.asp?page=<%=intend+1%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[����]</a>
                            <a href="tax_bill_in_mg.asp?page=<%=total_page%>&bill_month=<%=bill_month%>&slip_gubun=<%=slip_gubun%>&view_c=<%=view_c%>&view_d=<%=view_d%>&emp_name=<%=emp_name%>&ck_sw=<%="y"%>">[������]</a>
                        <% else %>
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

