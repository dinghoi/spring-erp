<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim from_date
Dim to_date

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	slip_month=Request("slip_month")
	confirm=Request("confirm")
	view_condi=Request("view_condi")
	condi=Request("condi")
  else
	slip_month=Request.form("slip_month")
	confirm=Request.form("confirm")
	view_condi=Request.form("view_condi")
	condi=Request.form("condi")
end if

if slip_month = "" then
	slip_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
	view_condi = "total"
	condi = ""
	confirm = "N"
end If

if view_condi = "total" then
	condi = ""
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

' ���Ǻ�
confirm_sql = " and confirm_yn = '" + confirm + "'"

if view_condi = "total" then
	condi_sql = ""
  else
  	condi_sql = " and " + view_condi + " = '" + condi + "'"
end if

base_sql = "select * from general_cost where (slip_gubun = '���') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"

Sql = "SELECT count(*) FROM general_cost where (slip_gubun = '���') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') " + confirm_sql + condi_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

order_sql = " ORDER BY slip_date ASC"

sql = base_sql + confirm_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "�Ϲݰ�� üũ"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ȸ�� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.slip_month.value == "") {
					alert ("�߻������ �Է��ϼ���.");
					return false;
				}	
				return true;
			}
			function frm1check () {
				if (chkfrm1()) {
					document.frm1.submit ();
				}
			}
			
			function chkfrm1() {
				{
				alert ("�����Ͻðڽ��ϱ�?");
					return true;
				}	
				return false;
			}
		</script>
		<script>
		function checkAll() {
			// üũ�ڽ����� �����´�.
			var checkObjs = $("input[type='checkbox']");
		
		 
			// ��ü�� ���õǾ��� ������ ���� �������� ������.
			if(checkObjs.length == $("input[type='checkbox']:checked").length) {
				checkObjs.prop("checked", false);
			}
			// ��ü�� ���õǾ��� ���� ������ ��ü ����~
			else {
				checkObjs.prop("checked", true);
			}
		} 
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/account_cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="general_cost_check.asp" method="post" name="frm">
				<fieldset class="srch">
				  <legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�߻����&nbsp;</strong>(��201401) : 
                                	<input name="slip_month" type="text" value="<%=slip_month%>" style="width:70px">
								</label>
								<label>
								<input name="confirm" type="radio" value="N"  <% if confirm = "N" then %>checked<% end if %> style="width:25px">
								��Ȯ��
                                <input name="confirm" type="radio" value="Y"  <% if confirm = "Y" then %>checked<% end if %> style="width:25px">
                                Ȯ��
								</label>
                                <label>
 								<strong>��ȸ����</strong>
                                <select name="view_condi" id="view_condi" style="width:150px">
                              		<option value="total" <% if view_condi = "total" then %>selected<% end if %>>��ü</option>
                                    <option value="emp_company" <% if view_condi = "emp_company" then %>selected<% end if %>>ȸ�纰</option>
                                    <option value="bonbu" <% if view_condi = "bonbu" then %>selected<% end if %>>���κ�</option>
                                    <option value="saupbu" <% if view_condi = "saupbu" then %>selected<% end if %>>����κ�</option>
                                    <option value="team" <% if view_condi = "team" then %>selected<% end if %>>����</option>
                                    <option value="reside_place" <% if view_condi = "reside_place" then %>selected<% end if %>>����ó��</option>
                                    <option value="emp_name" <% if view_condi = "emp_name" then %>selected<% end if %>>�����</option>
                                </select>
								</label>
                                <label>
								<strong>���� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<form action="general_cost_check_ok.asp" method="post" name="frm1">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="7%" >
							<col width="11%" >
							<col width="7%" >
							<col width="8%" >
							<col width="8%" >
							<col width="5%" >
							<col width="6%" >
							<col width="*" >
							<col width="5%" >
							<col width="4%" >
							<col width="14%" >
							<col width="4%" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">�߻�����</th>
								<th scope="col">�Ҽ�</th>
								<th scope="col">�����</th>
								<th scope="col">��뱸��</th>
								<th scope="col">����׸�</th>
								<th scope="col">����NO</th>
								<th scope="col">��û�ݾ�</th>
								<th scope="col">���ó</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">���</th>
								<td colspan="2" scope="col"><a href="javascript:;" onclick="checkAll()" class="btnType04">��ü</a></td>
							</tr>
						</thead>
						<tbody>
						<%
						cost_sum = 0
						pay_sum = 0
						mi_pay_sum = 0
						cancel_sum = 0
						i = 0
    					seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
							i = i + 1
							cost_sum = cost_sum + rs("cost")
							if rs("cancel_yn") = "Y" then
								cancel_sum = cancel_sum + rs("cost")
							end if
							if rs("cancel_yn") <> "Y" then
								if rs("pay_yn") = "Y" then
									pay_sum = pay_sum + rs("cost")
								  else
									mi_pay_sum = mi_pay_sum + rs("cost")
								end if
							end if							  

							if rs("pay_yn") = "Y" then
								pay_yn = "����"
							  else
							  	pay_yn = "������"
							end if
							if rs("cancel_yn") = "Y" then
								cancel_yn = "���"
							  else
							  	cancel_yn = "����"
							end if
							if rs("confirm_yn") = "Y" then
								confirm_yn = "Ȯ��"
							  else
							  	confirm_yn = "��Ȯ��"
							end if
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("emp_name")%>&nbsp;<%=rs("emp_grade")%></td>
								<td><%=rs("account")%></td>
								<td><%=rs("account_item")%></td>
								<td><%=rs("sign_no")%>&nbsp;</td>
							  	<td class="right">
								<a href="#" onClick="pop_Window('general_cost_mod.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>','common_cost_add_pop','scrollbars=yes,width=800,height=280')"><%=formatnumber(rs("cost"),0)%></a>
                                </td>
								<td><%=rs("customer")%></td>
								<td><%=pay_yn%></td>
								<td><%=cancel_yn%></td>
								<td><%=rs("slip_memo")%></td>
								<td><%=confirm_yn%></td>
							  	<td>
							<% if rs("confirm_yn") = "Y" then	%>
                                &nbsp;
                            <%   else	%>
                                <input name="confirm_yn" type="checkbox" id="confirm_yn" value="<%=i%>">
                            <% end if	%>
					            <input type="hidden" name="slip_date" value="<%=rs("slip_date")%>" ID="Hidden1">
					            <input type="hidden" name="slip_seq" value="<%=rs("slip_seq")%>" ID="Hidden1">
                                </td>
							</tr>
						<%
							rs.movenext()
  							seq = seq -1
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
				    <td width="20%">
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="general_cost_check.asp?page=<%=first_page%>&slip_month=<%=slip_month%>&confirm=<%=confirm%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="general_cost_check.asp?page=<%=intstart -1%>&slip_month=<%=slip_month%>&confirm=<%=confirm%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="general_cost_check.asp?page=<%=i%>&slip_month=<%=slip_month%>&confirm=<%=confirm%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="general_cost_check.asp?page=<%=intend+1%>&slip_month=<%=slip_month%>&confirm=<%=confirm%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a> <a href="general_cost_check.asp?page=<%=total_page%>&slip_month=<%=slip_month%>&confirm=<%=confirm%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnRight">
                    <span class="btnType01"><input type="button" value="Ȯ������" onclick="javascript:frm1check();" NAME="Button1"></span>
					</div>                  
                    </td>
			      </tr>
				  </table>
				<br>
	            <input type="hidden" name="tot_seq" value="<%=i%>" ID="Hidden1">
	            <input type="hidden" name="slip_month" value="<%=slip_month%>" ID="Hidden1">
	            <input type="hidden" name="acpt_confirm" value="<%=confirm%>" ID="Hidden1">
	            <input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">
	            <input type="hidden" name="condi" value="<%=condi%>" ID="Hidden1">
	            <input type="hidden" name="page" value="<%=page%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

