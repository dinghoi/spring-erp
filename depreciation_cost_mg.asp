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
		slip_month = request("slip_month")
		account = request("account")
	else
		slip_month = request.form("slip_month")
		account = request.form("account")
	end if

	if slip_month = "" then
		slip_month = mid(now(),1,4) + mid(now(),6,2)
		account = "��ü"
	end if

	from_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
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

	if cost_grade = "0" then
		posi_sql = ""
	end if

	if account = "��ü" then
		gubun_sql = ""
	  else
	  	gubun_sql = " and account = '"&account&"' "
	end if

	base_sql = "select * from general_cost where (slip_gubun = '�󰢺�') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
	order_sql = " ORDER BY org_name, emp_name, slip_date ASC"

	sql = "select count(*) from general_cost where (slip_gubun = '�󰢺�') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
	Set RsCount = Dbconn.Execute (sql)

	tottal_record = cint(RsCount(0)) 'Result.RecordCount

	IF tottal_record mod pgsize = 0 THEN
		total_page = int(tottal_record / pgsize) 'Result.PageCount
	  ELSE
		total_page = int((tottal_record / pgsize) + 1)
	END IF

	sql = "select sum(price) as price,sum(cost) as cost,sum(cost_vat) as cost_vat from general_cost where (slip_gubun = '�󰢺�') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + posi_sql + gubun_sql
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

	title_line = "�󰢺� ����"
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
				if (document.frm.slip_month.value == "") {
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
				<form action="depreciation_cost_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>��ȸ����</dt>
                        <dd>
                            <p>
								<label>
								<strong>����� : </strong>
                                	<input name="slip_month" type="text" value="<%=slip_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
                                <strong>�󰢺����� : </strong>
                                <select name="account" id="account" style="width:120px">
                                  <option value='��ü' <%If account = "��ü" then %>selected<% end if %>>��ü</option>
                                  <option value='��ջ󰢺�' <%If account = "��ջ󰢺�" then %>selected<% end if %>>��ջ󰢺�</option>
                                  <option value='�����ڻ�' <%If account = "�����ڻ�" then %>selected<% end if %>>�����ڻ�</option>
                                  <option value='�����ڻ�' <%If account = "�����ڻ�" then %>selected<% end if %>>�����ڻ�</option>
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
							<col width="15%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="*" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���ȸ��</th>
								<th scope="col">�������</th>
								<th scope="col">�����</th>
								<th scope="col">�ݾ�</th>
								<th scope="col">�󰢺�����</th>
								<th scope="col">�󰢺� ���γ���</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>�Ǽ�</strong></td>
								<td><%=formatnumber(tottal_record,0)%>&nbsp;��</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(sum_cost,0)%></td>
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
						%>
							<tr>
								<td class="first"><%=rs("emp_company")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("emp_name")%></td>
							  	<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("slip_memo")%></td>
								<td>
							<% if rs("end_yn") = "C" or rs("end_yn") = "N" then %>
							<%   if (rs("reg_id") = user_id) or (rs("emp_no") = user_id) or cost_grade = "0" or position ="�������" or position = "������"  then	%>
                                <a href="#" onClick="pop_Window('depreciation_cost_add.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','depreciation_cost_add_pop','scrollbars=yes,width=800,height=200')">����</a>
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
                    <a href="depreciation_cost_excel.asp?slip_month=<%=slip_month%>&account=<%=account%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="depreciation_cost_mg.asp?page=<%=first_page%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="depreciation_cost_mg.asp?page=<%=intstart -1%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="depreciation_cost_mg.asp?page=<%=i%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="depreciation_cost_mg.asp?page=<%=intend+1%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[����]</a>
                        <a href="depreciation_cost_mg.asp?page=<%=total_page%>&slip_month=<%=slip_month%>&account=<%=account%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="24%">
					<div class="btnRight">
					<a href="#" onClick="pop_Window('depreciation_cost_add.asp','depreciation_cost_add_pop','scrollbars=yes,width=800,height=200')" class="btnType04">�󰢺���</a>
					</div>
                    </td>
			      </tr>
				  </table>
				</form>
		</div>
	</div>
	</body>
</html>

