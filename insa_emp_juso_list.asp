<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_emp_juso_list.asp"

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "��ü"
	condi_sql = " "
	condi = ""
end if

if view_condi = "�Ҽ�������" then
	condi_sql = "emp_org_name like '%" + condi + "%' and "
end if
if view_condi = "����" then
	condi_sql = "emp_name like '%" + condi + "%' and "
end if
if view_condi = "ȸ�纰" then
	condi_sql = "emp_company like '%" + condi + "%' and "
end if
if view_condi = "���κ�" then
	condi_sql = "emp_bonbu like '%" + condi + "%' and "
end if
if view_condi = "����κ�" then
	condi_sql = "emp_saupbu like '%" + condi + "%' and "
end if
if view_condi = "����" then
	condi_sql = "emp_team like '%" + condi + "%' and "
end if
if view_condi = "����ó ȸ�纰" then
	condi_sql = "emp_reside_company like '%" + condi + "%' and "
end if

pgsize = 10 ' ȭ�� �� ������ 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


Sql = "SELECT count(*) FROM emp_master where "+condi_sql+" (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM emp_master where "+condi_sql+" (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - ���� �ּҷ� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				return "0 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('�����Ͻðڽ��ϱ�?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_emp_juso_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="��ü" <%If view_condi = "��ü" then %>selected<% end if %>>��ü</option>
                                  <option value="�Ҽ�������" <%If view_condi = "�Ҽ�������" then %>selected<% end if %>>�Ҽ�������</option>
                                  <option value="����" <%If view_condi = "����" then %>selected<% end if %>>����</option>
                                  <option value="ȸ�纰" <%If view_condi = "ȸ�纰" then %>selected<% end if %>>ȸ�纰</option>
                                  <option value="���κ�" <%If view_condi = "���κ�" then %>selected<% end if %>>���κ�</option>
                                  <option value="����κ�" <%If view_condi = "����κ�" then %>selected<% end if %>>����κ�</option>
                                  <option value="����" <%If view_condi = "����" then %>selected<% end if %>>����</option>
                                  <option value="����ó ȸ�纰" <%If view_condi = "����ó ȸ�纰" then %>selected<% end if %>>����ó ȸ�纰</option>
                                </select>
								<strong>���� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="6%" >
							<col width="7%" >
							<col width="8%" >
							<col width="15%" >
                            <col width="11%" >
                            <col width="11%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th scope="col" class="first">�Ҽ�</th>
                                <th scope="col">��  ��</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�����ּ�</th>
                                <th scope="col">������ȣ</th>
                                <th scope="col">�޴���ȭ</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                        emp_email = rs("emp_email") + "@k-won.co.kr"
	           			%>
							<tr>
                                <td class="first"><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('insa_emp_card.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_emp_card_pop','scrollbars=yes,width=500,height=500')"><%=rs("emp_name")%></a>&nbsp;</td>

                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td class="left"><%=emp_email%>&nbsp;</td>
                                <td><%=rs("emp_extension_no")%>&nbsp;</td>
                                <td><%=rs("emp_hp_ddd")%>-<%=rs("emp_hp_no1")%>-<%=rs("emp_hp_no2")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
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
				    <td>
                    <div id="paging">
                        <a href = "insa_emp_juso_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_emp_juso_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_emp_juso_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_emp_juso_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_emp_juso_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

