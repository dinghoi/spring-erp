<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
	use_yn = request.form("use_yn")
	emp_yn = request.form("emp_yn")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
	use_yn = request("use_yn")  
	emp_yn = request("emp_yn")  
end if

if view_condi = "" then
	view_condi = "��ü"
	condi_sql = " "
	condi = ""
	use_yn = "Y"
	emp_yn = "Y"
end if

if view_condi = "��ü" then
	condi = ""
end if


if view_condi = "CE��" then
	condi_sql = " and user_name like '%" + condi + "%'"
end if
if view_condi = "CE��" then
	condi_sql = "and user_name like '%" + condi + "%'"
end if
if view_condi = "�ҼӺ�" then
	condi_sql = "and team like '%" + condi + "%'"
end if
if view_condi = "����ó��" then
	condi_sql = "and reside_place like '%" + condi + "%'"
end if
if view_condi = "�����̵�" then
	condi_sql = "and old_user_id like '%" + condi + "%'"
end if

if use_yn = "Y" then
	use_sql = " and grade < '6'"
  else
  	use_sql = " and grade = '6'" 
end if
if emp_yn = "Y" then
	emp_sql = "(emp_no < '200000') "
  else
  	emp_sql = "(emp_no = '999999') " 
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

Sql = "SELECT count(*) FROM memb where "+emp_sql+condi_sql+use_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "SELECT * FROM memb where "+emp_sql+condi_sql+use_sql+" ORDER BY user_name ASC limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "CE ����"
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
				return "4 1";
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
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/ce_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="ce_mg_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<strong>������� : </strong>
                                <label>
                              	<input type="radio" name="use_yn" value="Y" <% if use_yn="Y" then %>checked<% end if %> style="width:30px">���
                              	<input type="radio" name="use_yn" value="N" <% if use_yn ="N" then %>checked<% end if %> style="width:30px">�̻��
								</label>
								<strong>�������� : </strong>
                                <label>
                              	<input type="radio" name="emp_yn" value="Y" <% if emp_yn ="Y" then %>checked<% end if %> style="width:30px">����
                              	<input type="radio" name="emp_yn" value="N" <% if emp_yn ="N" then %>checked<% end if %> style="width:30px">����
								</label>
								<strong>�׸����� : </strong>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="��ü" <%If view_condi = "��ü" then %>selected<% end if %>>��ü</option>
                                  <option value="CE��" <%If view_condi = "CE��" then %>selected<% end if %>>CE��</option>
                                  <option value="����" <%If view_condi = "����" then %>selected<% end if %>>����</option>
                                  <option value="����ó��" <%If view_condi = "����ó��" then %>selected<% end if %>>����ó��</option>
                                  <option value="�����̵�" <%If view_condi = "�����̵�" then %>selected<% end if %>>�����̵�</option>
                                </select>
								<strong>���� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�̸�</th>
								<th scope="col">���̵�</th>
								<th scope="col">�����̵�</th>
								<th scope="col">�Ҽ�</th>
								<th scope="col">�ڵ���</th>
								<th scope="col">����</th>
								<th scope="col">�����׷�</th>
								<th scope="col">����ó</th>
								<th scope="col">��纯��</th>
								<th scope="col">����</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							if rs("grade") = 0 then
								grade_view = "������"
							end if
							if rs("grade") = 1 then
								grade_view = "������"
							end if
							if rs("grade") = 2 then
								grade_view = "���ְ�����"
							end if
							if rs("grade") = 3 then
								grade_view = "����CE"
							end if
							if rs("grade") = 4 then
								grade_view = "CE"
							end if
							if rs("grade") = 5 then
								grade_view = "�����"
							end if
							if rs("grade") > 5 or rs("grade") < 0 then
								grade_view = "���Ѿ���"
							end if
							i = i + 1

							if rs("mg_group") = "2" then
								mg_group = "�����׷�"
							  elseif rs("mg_group") = "1" then
							  	mg_group = "�Ϲݱ׷�"
							  else
							  	mg_group = "Error"
							end if
	           			%>
							<tr>
								<td class="first"><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>
								<td><a href="#" onClick="pop_Window('pass_init.asp?user_id=<%=rs("user_id")%>','pass_init_pop','scrollbars=no,width=400,height=200')"><%=rs("user_id")%></a></td>
								<td><%=rs("old_user_id")%>&nbsp;</td>
								<td class="left"><%=rs("bonbu")%>&nbsp;<%=rs("saupbu")%>&nbsp;<%=rs("team")%></td>
								<td><%=rs("hp")%></td>
								<td><%=grade_view%></td>
								<td><%=mg_group%></td>
								<td><%=rs("reside_place")%>&nbsp;</td>
								<td>
							<% if rs("org_name") = "�����" or rs("org_name") = "���ְ���" then	%>
								&nbsp;
                            <%   else	%>
                                <a href="#" onClick="pop_Window('ce_exchange.asp?ce_id=<%=rs("user_id")%>&team=<%=rs("team")%>','ce_change','scrollbars=yes,width=750,height=600')">�ް�/��ü</a>
                            <% end if	%>
                                </td>
								<td>
							<% if rs("org_name") <> "�����" then	%>
                                <a href="#" onClick="pop_Window('ce_reg.asp?user_id=<%=rs("user_id")%>&u_type=<%="U"%>','ce_pop','scrollbars=no,width=800,height=300')">����</a><input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=rs("user_id")%>">
                            <%   else	%>
								&nbsp;
                            <% end if	%>
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
				    <td width="15%"></td>
				    <td>
                  <div id="paging">
                        <a href = "ce_mg_list.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="ce_mg_list.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="ce_mg_list.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="ce_mg_list.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[����]</a> <a href="ce_mg_list.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%" align="center"><a href="#" onclick="javascript:pop_ce()" class="btnType04">CE ���</a></td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

