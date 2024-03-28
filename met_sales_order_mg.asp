<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(160)

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	slip_id=Request("slip_id")
	view_date=Request("view_date")
	field_check=Request("field_check")
	field_view=Request("field_view")
  else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	slip_id=Request.form("slip_id")
	view_date=Request.form("view_date")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If view_date = "" Then
'	to_date = mid(cstr(now()),1,10)
'	from_date = mid(cstr(now()-31),1,10)
	field_check = "total"
	slip_id = "T"
	view_date = "total"
	to_date = ""
	from_date = ""
End If

If view_date = "total" Then
	to_date = ""
	from_date = ""
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
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_buy = Server.CreateObject("ADODB.Recordset")
Set Rs_sale = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

base_sql = "select * from sales_slip "
if view_date = "total" then
	date_sql = "where (sign_yn = 'Y') "
  else
  	date_sql = "where (sign_yn = 'Y') and (" + view_date + " >='" + from_date  + "' and " + view_date + " <= '" + to_date  + "') "
end if

if slip_id = "T" then
	slip_sql = " "
  else
	slip_sql = " and slip_id = '"+ slip_id + "'"
end if

if field_check = "total" then
  	field_sql = " "
  else
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
end if

order_sql = " ORDER BY slip_no DESC"

Sql = "SELECT count(*) FROM sales_slip " + date_sql + slip_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + slip_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
'response.write(sql)
Rs.Open Sql, Dbconn, 1

title_line = "�������� ���Ź��� ����"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��ǰ������� �ý���</title>
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
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
				if (document.frm.view_date.value != "total") {
					if (document.frm.from_date.value == "" || document.frm.to_date.value == "") {
						alert ("���ڸ� �Է��ϼ���");
						return false;
				}}	
				return true;
			}

		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/meterials_control_header.asp" -->
            <!--#include virtual = "/include/meterials_purchase_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_sales_order_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
								<strong>��ǥ���� : </strong>
                                <select name="slip_id" id="slip_id" style="width:80px">
                              		<option value="T" <% if slip_id = "T" then %>selected<% end if %>>��ü</option>
                                    <option value="2" <% if slip_id = "2" then %>selected<% end if %>>������ǥ</option>
                                    <option value="1" <% if slip_id = "1" then %>selected<% end if %>>�����ǥ</option>
                                </select>
								<strong>���ں� �˻� : </strong>
                                <select name="view_date" id="view_date" style="width:150px">
                                    <option value="total" <% if view_date = "total" then %>selected<% end if %>>��ü</option>
                                    <option value="sales_date" <% if view_date = "sales_date" then %>selected<% end if %>>��������</option>
                                    <option value="bill_issue_date" <% if view_date = "bill_issue_date" then %>selected<% end if %>>��꼭������</option>
                                    <option value="bill_due_date" <% if view_date = "bill_due_date" then %>selected<% end if %>>��꼭�����Ͽ�����</option>
                                    <option value="out_request_date" <% if view_date = "out_request_date" then %>selected<% end if %>>����û��</option>
                                    <option value="collect_due_date" <% if view_date = "collect_due_date" then %>selected<% end if %>>���ݿ�����</option>
                                    <option value="collect_date" <% if view_date = "collect_date" then %>selected<% end if %>>���ݿϷ���</option>
                                </select>
								<label>
								����
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								~
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <label>
								<strong>���� : </strong>
                                <select name="field_check" id="field_check" style="width:80px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                    <option value="slip_no" <% if field_check = "slip_no" then %>selected<% end if %>>��ǥ��ȣ</option>
                                    <option value="trade_name" <% if field_check = "trade_name" then %>selected<% end if %>>�ŷ�ó��</option>
                                    <option value="emp_name" <% if field_check = "emp_name" then %>selected<% end if %>>�������</option>
                                </select>
								</label>
                                <label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:100px" id="field_view" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
							<col width="6%" >
							<col width="8%" >
							<col width="6%" >
                            <col width="8%" >
                            <col width="8%" >
							<col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="*" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��ǥ����</th>
								<th scope="col">�����û</th>
								<th scope="col">��ǥ��ȣ</th>
								<th scope="col">��������</th>
								<th scope="col">�ŷ�ó��</th>
								<th scope="col">�������</th>
                                <th scope="col">�����Ҽ�</th>
								<th scope="col">����ǰ��</th>
                                <th scope="col">�����Ѿ�</th>
                                <th scope="col">����û��</th>
                                <th scope="col">�������</th>
                                <th scope="col">����</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
						<%
    					seq = total_record - ( page - 1 ) * pgsize
						do until rs.eof
						 	slip_id = rs("slip_id")
						    slip_no = rs("slip_no")
						    slip_seq = rs("slip_seq")
							sales_date = rs("sales_date")
							
							if rs("slip_id") = "2" then
								slip_id_view = "������ǥ"
							end if
							if rs("slip_id") = "1" then
								slip_id_view = "�����ǥ"
							end if

							if rs("sign_yn") = "Y" then
								sign_view = "����Ϸ�"
							  elseif rs("sign_yn") = "N" then 
								sign_view = "�̰���"
							  else
								sign_view = "������"
							end if
							
							sql = "select * from sales_slip_detail where slip_no = '"&slip_no&"' and slip_id = '"&slip_id&"' and slip_seq = '"&slip_seq&"'"
						    Set Rs_buy=DbConn.Execute(Sql)
						    if Rs_buy.eof or Rs_buy.bof then
							 	   bg_goods_name = ""
							   else
							  	   bg_goods_name = Rs_buy("pummok")
						    end if
						    Rs_buy.close()
							
							sql = "select * from met_order where order_id = '"&slip_id&"' and order_buy_no = '"&slip_no&"' and order_buy_seq = '"&slip_seq&"' and order_buy_date = '"&sales_date&"'"
						    Set Rs_buy=DbConn.Execute(Sql)
						    if Rs_buy.eof or Rs_buy.bof then
							 	   sales_ing_name = "��û��"
							   else
							  	   sales_ing_name = "���ֿϷ�"
						    end if
						    Rs_buy.close()
							
							sql = "select * from met_chulgo where (chulgo_id = '�����ǥ' or chulgo_id = '������ǥ') and rele_no = '"&slip_no&"' and rele_seq = '"&slip_seq&"' and rele_date = '"&sales_date&"'"
						    Set Rs_buy=DbConn.Execute(Sql)
						    if not Rs_buy.eof or not Rs_buy.bof then
							  	   sales_ing_name = "���Ϸ�"
						    end if
						    Rs_buy.close()
							
							task_memo = replace(rs("slip_memo"),chr(34),chr(39))
							view_memo = task_memo
							if len(task_memo) > 16 then
								view_memo = mid(task_memo,1,16) + ".."
							end if
							
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=slip_id_view%></td>
								<td><%=sign_view%></td>
								<td>
								<a href="#" onClick="pop_Window('sales_slip_view.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>','sales_slip_view_pop','scrollbars=yes,width=930,height=650')"><%=rs("slip_no")%>-<%=rs("slip_seq")%></a>
                                </td>
								<td><%=rs("sales_date")%>&nbsp;</td>
								<td><%=rs("trade_name")%></td>
								<td><%=rs("emp_name")%></td>
                                <td><%=rs("sales_saupbu")%>&nbsp;</td>
                                <td><%=bg_goods_name%>&nbsp;��</td>
								<td class="right"><%=formatnumber(rs("buy_cost"),0)%></td>
                                <td><%=rs("out_request_date")%></td>
                                <td><%=sales_ing_name%></td>
                                <td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
							  	<td>
                        <% if sales_ing_name = "��û��" then	%>
                                <a href="#" onClick="pop_Window('met_sales_order_add.asp?slip_id=<%=rs("slip_id")%>&slip_no=<%=rs("slip_no")%>&slip_seq=<%=rs("slip_seq")%>&sales_date=<%=rs("sales_date")%>&u_type=<%=""%>','met_sales_order_pop','scrollbars=yes,width=1230,height=650')">���</a>
						<%   else	%>
								-
						<% end if	%>
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
                    <a class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="met_sales_order_mg.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="met_sales_order_mg.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="met_sales_order_mg.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="met_sales_order_mg.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a> 
                        <a href="met_sales_order_mg.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&slip_id=<%=slip_id%>&view_date=<%=view_date%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="20%">
					<div class="btnCenter">
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

