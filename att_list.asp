<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(150)

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	company=Request("company")
	as_type=Request("as_type")
	field_check=Request("field_check")
	field_view=Request("field_view")

Else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	company=Request.form("company")
	as_type=Request.form("as_type")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If to_date = "" or from_date = "" Then
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-7),1,10)
	field_check = "total"
	company = "��ü"
	as_type = "��ü"
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

' ���Ǻ� ��ȸ.........
' ��¥�� ��ȸ(1)
base_sql = "select *  from att_file "
date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "')"
if company = "��ü" then
	company_sql = ""
  else
	company_sql = " and ( company = '" + company + "') "
end if
if as_type = "��ü" then
	type_sql = ""
  else
	type_sql = " and ( as_type = '" + as_type + "') "
end if

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if
order_sql = " ORDER BY visit_date DESC"

Sql = "SELECT count(*) FROM att_file " + date_sql + company_sql + type_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + company_sql + type_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1


title_line = "��ġ/���� ÷�ΰ���"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
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
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="att_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
                                <label>
								<strong>ȸ��</strong>
								<%
                                if c_grade = "7" or (c_grade = "5" and c_reside = "1") then
                                    sql_trade="select * from trade where use_sw = 'Y' and mg_group = '"+mg_group+"' and trade_name = '"+user_name+"' order by etc_name asc"
                                end if
                                rs_trade.Open sql_trade, Dbconn, 1
                                %>
                                <select name="company" id="company">
 									<option value="��ü">��ü</option> 
          					<% 
								While not rs_trade.eof 
							%>
          							<option value='<%=rs_trade("trade_name")%>' <%If rs_trade("trade_name") = company  then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
          					<%
									rs_trade.movenext()  
								Wend 
								rs_trade.Close()
							%>
                                </select>
								</label>
								<label>
								<strong>�����&nbsp;&nbsp;���� : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>���� : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<label>
								<strong>ó������</strong>
                                <select name="as_type" id="as_type" style="width:100px">
                                  <option value="��ü" <%If as_type = "��ü" then %>selected<% end if %>>��ü</option>
                                  <option value="�űԼ�ġ" <%If as_type = "�űԼ�ġ" then %>selected<% end if %>>�űԼ�ġ</option>
                                  <option value="�űԼ�ġ����" <%If as_type = "�űԼ�ġ����" then %>selected<% end if %>>�űԼ�ġ����</option>
                                  <option value="������ġ" <%If as_type = "������ġ" then %>selected<% end if %>>������ġ</option>
                                  <option value="������ġ����" <%If as_type = "������ġ����" then %>selected<% end if %>>������ġ����</option>
                                  <option value="������" <%If as_type = "������" then %>selected<% end if %>>������</option>
                                  <option value="����������" <%If as_type = "����������" then %>selected<% end if %>>����������</option>
                                  <option value="���ȸ��" <%If as_type = "���ȸ��" then %>selected<% end if %>>���ȸ��</option>
                                  <option value="��������" <%If as_type = "��������" then %>selected<% end if %>>��������</option>
                                </select>
								</label>
                                <label>
								<strong>���ǰ˻�</strong>
                                <select name="field_check" id="field_check" style="width:80px">
                                    <option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                    <option value="acpt_no" <% if field_check = "acpt_no" then %>selected<% end if %>>������ȣ</option>
                                    <option value="mg_ce" <% if field_check = "mg_ce" then %>selected<% end if %>>���CE</option>
                                    <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>�õ�</option>
                                    <option value="gugun" <% if field_check = "gugun" then %>selected<% end if %>>����</option>
                                    <option value="dept" <% if field_check = "dept" then %>selected<% end if %>>������</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:80px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="7%" >
							<col width="12%" >
							<col width="18%" >
							<col width="13%" >
							<col width="6%" >
							<col width="6%" >
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">ó������</th>
								<th scope="col">ó������</th>
								<th scope="col">ȸ��</th>
								<th scope="col">�μ�</th>
								<th scope="col">����</th>
								<th scope="col">���CE</th>
								<th scope="col">������ȣ</th>
								<th scope="col">÷������</th>
								<th scope="col">÷�κ���</th>
								<th scope="col">���γ���</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							path = "/att_file/" + rs("company")
						%>
							<tr>
								<td class="first"><%=rs("as_type")%></td>
								<td><%=rs("visit_date")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=rs("mg_ce")%></td>
								<td><%=rs("acpt_no")%></td>
								<td>&nbsp;
								<%
                                    if rs("att_file1") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file1")%>">÷��1</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file2") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file2")%>">÷��2</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file3") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file3")%>">÷��3</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file4") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file4")%>">÷��4</a>&nbsp;
                                <%
                                    end if
                                    if rs("att_file5") <> "" then		
                                %>
                                        <a href="download.asp?path=<%=path%>&att_file=<%=rs("att_file5")%>">÷��5</a>&nbsp;
                                <%
                                    end if
                                %>
                                </td>
								<td><a href="#" onClick="pop_Window('att_file_mod.asp?acpt_no=<%=rs("acpt_no")%>','att_file_mod_pop','scrollbars=yes,width=800,height=410')">����</a></td>
								<td><a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=rs("acpt_no")%>','asview_pop','scrollbars=yes,width=800,height=700')">��ȸ</a></td>
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
				    <td width="15%"></td>
				    <td>
                    <div id="paging">
                        <a href = "att_list.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="att_list.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="att_list.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="att_list.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[����]</a> <a href="att_list.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&company=<%=company%>&as_type=<%=as_type%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%"></td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

