<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_emp_org_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	view_condi=Request("view_condi")
  else
	view_condi=Request.form("view_condi")
End if

If view_condi = "" Then
	view_condi = "��ü"
End If

pgsize = 10 ' ȭ�� �� ������

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "ASC"
end if

order_Sql = " ORDER BY emp_bonbu,emp_saupbu,emp_team,emp_org_name,emp_no,emp_in_date " + view_sort

if view_condi = "��ü" then
       where_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_no < '900000')"
   else
	   where_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date = '0000-00-00') and (emp_company = '"&view_condi&"') and (emp_no < '900000')"
end if

Sql = "SELECT count(*) FROM emp_master " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_master " + where_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1

title_line = " ������ü ���� ��Ȳ "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_org_list.asp" method="post" name="frm">

				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>ȸ�� �˻�</dt>
                        <dd>
                            <p>
                               <strong>ȸ�� : </strong>
                              <%
								Sql="select * from emp_org_mst where (org_level = 'ȸ��') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
                                    <option value="��ü" <%If view_condi = "��ü" then %>selected<% end if %>>��ü</option>

                			  <%
								do until rs_org.eof
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()
								loop
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>

				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
                            <col width="5%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="12%" >
                            <col width="16%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>

								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�Ի���</th>
                                <th scope="col">�Ҽ�</th>
								<th scope="col">�Ҽӹ߷���</th>

                                <th scope="col">�������</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <th scope="col">����ȸ��</th>
                                <th scope="col">�Ǳٹ���</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof

						if rs("emp_org_baldate") = "1900-01-01" then
						   emp_org_baldate = ""
						   else
						   emp_org_baldate = rs("emp_org_baldate")
						end if
						%>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
                                <td class="left"><%=rs("emp_reside_company")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_stay_name")%>&nbsp;</td>
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
				    <td width="20%">
					<div class="btnCenter">
                    <a href="insa_emp_org_list_excel.asp?view_condi=<%=view_condi%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_emp_org_list.asp?page=<%=first_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_emp_org_list.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_emp_org_list.asp?page=<%=i%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_emp_org_list.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_emp_org_list.asp?page=<%=total_page%>&view_sort=<%=view_sort%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[������]</a>
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
			</form>
		</div>
	</div>
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

