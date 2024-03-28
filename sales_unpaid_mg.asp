<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim as_process
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(160)

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	field_check=Request("field_check")
	field_view=Request("field_view")
	view_sw=Request("view_sw")
	curr_date=Request("curr_date")
  else
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	view_sw=Request.form("view_sw")
	curr_date=Request.form("curr_date")
End if

If field_check = "" Then
	field_check = "total"
	view_sw = 0
	curr_date = mid(cstr(now()),1,10)
End If

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' ȭ�� �� ������

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql="select max(collect_date) as max_date from sales_collect"
set rs=dbconn.execute(sql)

if	isnull(rs("max_date"))  then
	max_date = "2015-11-01"
  else
	max_date = rs("max_date")
end if
rs.close()


base_sql = "select * from saupbu_sales where (sales_amt <> collect_tot_amt) "

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

if view_sw = "1" then
	view_sql = " and ( collect_due_date < '"&curr_date&"' ) "
  elseif view_sw = "2" then
	view_sql = " and ( collect_due_date >= '"&curr_date&"' ) "
  else
  	view_sql = " "
end if

order_sql = " ORDER BY emp_name, company, sales_date ASC"

Sql = "SELECT count(*) FROM saupbu_sales where (sales_amt <> collect_tot_amt) " + field_sql + view_sql
Set RsCount = Dbconn.Execute (sql)

total_record = CLng(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

sql = "select sum(sales_amt) as price,sum(collect_tot_amt) as collect from saupbu_sales where (sales_amt <> collect_tot_amt) " + field_sql + view_sql
Set rs_sum = Dbconn.Execute (sql)
if isnull(rs_sum("price")) then
	tot_sales_amt = 0
	tot_collect_tot_amt = 0
  else
	tot_sales_amt = cdbl(rs_sum("price"))
	tot_collect_tot_amt = cdbl(rs_sum("collect"))
end if

sql = base_sql + field_sql + view_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1

title_line = "�̼��� ����"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=curr_date%>" );
			});
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.field_check.value == "") {
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}
				if (document.frm.max_date.value > document.frm.curr_date.value) {
					alert ("�������ڰ� �Ա����ں��� �۽��ϴ�.");
					frm.unpaid_memo.focus();
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/sales_unpaid_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="sales_unpaid_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
							<label>
                               <select name="field_check" id="field_check" style="width:80px">
                           		<option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                <option value="slip_no" <% if field_check = "slip_no" then %>selected<% end if %>>��ǥ��ȣ</option>
                                <option value="company" <% if field_check = "company" then %>selected<% end if %>>�ŷ�ó��</option>
                                <option value="emp_name" <% if field_check = "emp_name" then %>selected<% end if %>>�������</option>
                               </select>
							</label>
                            <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:120px" id="field_view" >
							</label>
							<label>
                                <input type="radio" name="view_sw" value="0" <% if view_sw = "0" then %>checked<% end if %> style="width:30px" id="Radio3"><strong>��ü</strong>
                                <input type="radio" name="view_sw" value="1" <% if view_sw = "1" then %>checked<% end if %> style="width:30px" id="Radio3"><strong>����</strong>
                                <input type="radio" name="view_sw" value="2" <% if view_sw = "2" then %>checked<% end if %> style="width:30px" id="Radio3"><strong>�̵���</strong>
							</label>
							<label>
								<strong>�������� : </strong>
                                	<input name="curr_date" type="text" style="width:70px" id="datepicker">
                            </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="14%" >
							<col width="5%" >
							<col width="8%" >
							<col width="7%" >
							<col width="8%" >
							<col width="10%" >
							<col width="*" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��ǥ��ȣ</th>
								<th scope="col">��������</th>
								<th scope="col">����<br>������</th>
								<th scope="col">�̼���<br>������</th>
								<th scope="col">�ŷ�ó��</th>
								<th scope="col">�������</th>
								<th scope="col">�����Ѿ�</th>
								<th scope="col">�Ѽ��ݾ�</th>
								<th scope="col">�ܾ�</th>
								<th scope="col">��������</th>
								<th scope="col">�̼��� ����</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>�Ǽ�</strong></td>
								<td><strong><%=formatnumber(total_record,0)%>��<strong></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=formatnumber(tot_sales_amt,0)%></td>
								<td class="right"><%=formatnumber(tot_collect_tot_amt,0)%></td>
								<td class="right"><%=formatnumber(tot_sales_amt - tot_collect_tot_amt,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
    					seq = total_record - ( page - 1 ) * pgsize
						do until rs.eof
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=mid(rs("slip_no"),1,17)%>&nbsp;</td>
								<td><%=rs("sales_date")%></td>
								<td><%=rs("collect_due_date")%>&nbsp;</td>
								<td><%=rs("unpaid_due_date")%>&nbsp;</td>
								<td><%=rs("company")%></td>
								<td><%=rs("emp_name")%></td>
								<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("collect_tot_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_amt")-rs("collect_tot_amt"),0)%></td>
								<td><%=rs("change_memo")%>&nbsp;</td>
								<td><%=rs("unpaid_memo")%>&nbsp;</td>
							  	<td>
                                <a href="#" onClick="pop_Window('sales_collect_add.asp?approve_no=<%=rs("approve_no")%>&u_type=<%="U"%>','sales_collect_add_pop','scrollbars=yes,width=1000,height=700')">���</a>
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
				    <td width="15%">
					<div class="btnCenter">
                    <a href="sales_unpaid_excel.asp?field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&curr_date=<%=curr_date%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="sales_unpaid_mg.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&curr_date=<%=curr_date%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="sales_unpaid_mg.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&curr_date=<%=curr_date%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
               	  <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sales_unpaid_mg.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&curr_date=<%=curr_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
               	  <% if 	intend < total_page then %>
                        <a href="sales_unpaid_mg.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&curr_date=<%=curr_date%>&ck_sw=<%="y"%>">[����]</a>
                        <a href="sales_unpaid_mg.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&curr_date=<%=curr_date%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="max_date" value="<%=max_date%>">
			</form>
		</div>
	</div>
	</body>
</html>

