<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_emp_yryc_list.asp"

user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

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

condi_sql = " 1=1 " 

if view_condi = "���" then
	condi_sql = condi_sql & " and (emp_no = '" + condi + "')  "
end if
if view_condi = "����" then
	condi_sql = condi_sql & " and (emp_name like '%" + condi + "%') "
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

Sql = "SELECT count(*) FROM emp_use_yryc where "&condi_sql&" "
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM emp_use_yryc where "+condi_sql+" ORDER BY  emp_end_date desc, yryc_sn desc limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " �ټ�1��̸�"
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
				return "1 1";
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
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="<%=be_pg%>?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="����" <%If view_condi = "����" then %>selected<% end if %>>����</option>
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
							<col width="15%" >
							<col width="15%" >
							<col width="15%" >
							<col width="20%" >
							<col width="20%" >
							<col width="15%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��  ��</th>
								<th scope="col">�������</th>
								<th scope="col">�Ի���</th>
								<th scope="col">�����</th>
                                <th scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
								<td class="first"><%=rs("yryc_sn")%>&nbsp;</td>
                                <td><%=rs("emp_name")%></td>
                                <td><%=rs("emp_person1")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_end_date")%>&nbsp;</td>
                                <td>
								<a href="#" onclick="pop_Window('yryc_certificate_print.asp?yryc_sn=<%=rs("yryc_sn")%>','yryc_certificate_print','scrollbars=yes,width=800,height=700');return false;" >���</a>
                                <!-- input type="image" id="btnPrint" src="/image/b_certifi.jpg" alt="���������ް�����ϼ�Ȯ�� ���" onclick="pop_Window('yryc_certificate_print.asp?yryc_sn=<%=rs("yryc_sn")%>','yryc_certificate_print','scrollbars=yes,width=1250,height=480');return false;" style="border-width:0px;" //-->
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
				    <td>
                  <div id="paging">
                        <a href = "<%=be_pg%>?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="<%=be_pg%>?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="<%=be_pg%>?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="<%=be_pg%>?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a> <a href="<%=be_pg%>?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="emp_no" value="<%=emp_no%>" ID="Hidden1">
                  <input type="hidden" name="emp_name" value="<%=emp_name%>" ID="Hidden1">
                  <input type="hidden" name="emp_company" value="<%=emp_company%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

