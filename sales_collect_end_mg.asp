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
dim sum_tab(4,2)

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	field_check=Request("field_check")
	field_view=Request("field_view")
	view_sw=Request("view_sw")
  else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	view_sw=Request.form("view_sw")
End if

If to_date = "" or from_date = "" Then
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-31),1,10)
	field_check = "total"
	view_sw = 0
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
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

base_sql = "select sales_collect.*, saupbu_sales.sales_date, saupbu_sales.company, saupbu_sales.sales_amt, saupbu_sales.collect_tot_amt, saupbu_sales.emp_name from saupbu_sales INNER JOIN sales_collect ON saupbu_sales.approve_no = sales_collect.approve_no where (collect_id = '4') and (collect_date >='"&from_date&"' and collect_date <= '"&to_date&"') "

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = " ORDER BY emp_name, company, sales_date,collect_date, slip_no, collect_seq ASC"

Sql = "SELECT count(*) FROM saupbu_sales INNER JOIN sales_collect ON saupbu_sales.approve_no = sales_collect.approve_no where (collect_id = '4') and (collect_date >='"&from_date&"' and collect_date <= '"&to_date&"') " + field_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

for i = 0 to 4
	sum_tab(i,1) = 0
	sum_tab(i,2) = 0
next

sql = "select sum(sales_amt) as price,sum(collect_tot_amt) as collect from saupbu_sales INNER JOIN sales_collect ON saupbu_sales.approve_no = sales_collect.approve_no where (collect_id = '4') and (collect_date >='"&from_date&"' and collect_date <= '"&to_date&"') " + field_sql + " group by bill_collect"

Set rs_sum = Dbconn.Execute (sql)

If rs_sum.EOF Or rs_sum.BOF Then
	tot_sales_amt = 0
	tot_collect_tot_amt = 0
Else
	if isnull(rs_sum("price")) then
		tot_sales_amt = 0
		tot_collect_tot_amt = 0
	else
		tot_sales_amt = cdbl(rs_sum("price"))
		tot_collect_tot_amt = cdbl(rs_sum("collect"))
	end if
End If

sql = base_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1

title_line = "�ԱݿϷ� ó�� ����"
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
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
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
				<form action="sales_collect_end_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
								<strong>ó������  </strong>
								<label>
								������
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								������
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
							����
							<label>
                               <select name="field_check" id="field_check" style="width:80px">
                           		<option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                <option value="bill_collect" <% if field_check = "bill_collect" then %>selected<% end if %>>���ݹ��</option>
                                <option value="saupbu_sales.slip_no" <% if field_check = "saupbu_sales.slip_no" then %>selected<% end if %>>��ǥ��ȣ</option>
                                <option value="company" <% if field_check = "company" then %>selected<% end if %>>�ŷ�ó��</option>
                                <option value="emp_name" <% if field_check = "emp_name" then %>selected<% end if %>>�������</option>
                               </select>
							</label>
                            <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:120px" id="field_view" >
							</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="*" >
							<col width="5%" >
							<col width="7%" >
							<col width="7%" >
							<col width="12%" >
							<col width="15%" >
							<col width="5%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">ó������</th>
								<th scope="col">��ǥ��ȣ</th>
								<th scope="col">��������</th>
								<th scope="col">�ŷ�ó��</th>
								<th scope="col">�������</th>
								<th scope="col">�����Ѿ�</th>
								<th scope="col">�ܾ�</th>
								<th scope="col">��������</th>
								<th scope="col">�̼��� ����</th>
								<th scope="col">�����</th>
								<th scope="col">��ȸ</th>
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
								<td class="right"><%=formatnumber(tot_sales_amt,0)%></td>
								<td class="right"><%=formatnumber(tot_sales_amt - tot_collect_tot_amt,0)%></td>
								<td>&nbsp;</td>
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
								<td><%=rs("collect_date")%></td>
								<td><%=mid(rs("slip_no"),1,17)%></td>
								<td><%=rs("sales_date")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("emp_name")%></td>
								<td class="right"><%=formatnumber(rs("sales_amt"),0)%></td>
								<td class="right"><%=formatnumber(rs("sales_amt")-rs("collect_tot_amt"),0)%></td>
								<td><%=rs("change_memo")%>&nbsp;</td>
								<td><%=rs("unpaid_memo")%>&nbsp;</td>
								<td><%=rs("reg_name")%></td>
							  	<td>
                                <a href="#" onClick="pop_Window('sales_collect_view.asp?approve_no=<%=rs("approve_no")%>','sales_collect_view_pop','scrollbars=yes,width=700,height=400')">��ȸ</a>
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
                    <a href="sales_collect_end_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="sales_collect_mg.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="sales_collect_mg.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sales_collect_mg.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
<% if 	intend < total_page then %>
                        <a href="sales_collect_mg.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[����]</a>
                        <a href="sales_collect_mg.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&view_sw=<%=view_sw%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
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

