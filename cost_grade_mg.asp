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
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "��ü"
	condi_sql = " "
	condi = ""
end if

if view_condi = "�����" then
	condi_sql = " and user_name like '%" + condi + "%'"
end if
if view_condi = "���޺�" then
	condi_sql = " and user_grade like '%" + condi + "%'"
end if
if view_condi = "������" then
	condi_sql = " and position like '%" + condi + "%'"
end if
if view_condi = "����" then
	condi_sql = "and team like '%" + condi + "%'"
end if
if view_condi = "����ó��" then
	condi_sql = "and reside_place like '%" + condi + "%'"
end if

use_sql = " and grade < '5'"
emp_sql = "(emp_no < '200000') "

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

title_line = "����ں� ��� ���� ����"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
		</script>
	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_grade_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<strong>�׸����� : </strong>
                                <select name="view_condi" id="select3" style="width:150px">
                                  <option value="��ü" <%If view_condi = "��ü" then %>selected<% end if %>>��ü</option>
                                  <option value="�����" <%If view_condi = "�����" then %>selected<% end if %>>�����</option>
                                  <option value="���޺�" <%If view_condi = "���޺�" then %>selected<% end if %>>���޺�</option>
                                  <option value="������" <%If view_condi = "������" then %>selected<% end if %>>������</option>
                                  <option value="����" <%If view_condi = "����" then %>selected<% end if %>>����</option>
                                  <option value="����ó��" <%If view_condi = "����ó��" then %>selected<% end if %>>����ó��</option>
                                </select>
								<strong>���� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="8%" >
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
							<col width="12%" >
							<col width="8%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�̸�</th>
								<th scope="col">���̵�</th>
								<th scope="col">�Ҽ�</th>
								<th scope="col">�ڵ���</th>
								<th scope="col">���񽺱���</th>
								<th scope="col">�����׷�</th>
								<th scope="col">����ó</th>
								<th scope="col">������</th>
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

							if rs("cost_grade") = 0 then
								cost_grade_view = "������"
							end if
							if rs("cost_grade") = 1 then
								cost_grade_view = "���������"
							end if
							if rs("cost_grade") = 2 then
								cost_grade_view = "����������"
							end if
							if rs("cost_grade") = 3 then
								cost_grade_view = "������"
							end if
							if rs("cost_grade") = 4 then
								cost_grade_view = "�����װ���"
							end if
							if rs("cost_grade") = 5 then
								cost_grade_view = "�Ϲ�CE/����"
							end if
							if rs("cost_grade") = 6 then
								cost_grade_view = "�Ϲ�CE"
							end if
							if rs("cost_grade") = 7 then
								cost_grade_view = "���Ѿ���"
							end if

							if rs("mg_group") = "2" then
								mg_group = "�����׷�"
							  elseif rs("mg_group") = "1" then
							  	mg_group = "�Ϲݱ׷�"
							  else
							  	mg_group = "Error"
							end if
							i = i + 1
	           			%>
							<tr>
								<td class="first"><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>
								<td><a href="#" onClick="pop_Window('pass_init.asp?user_id=<%=rs("user_id")%>','pass_init_pop','scrollbars=no,width=400,height=200')"><%=rs("user_id")%></a></td>
								<td class="left"><%=rs("bonbu")%>&nbsp;<%=rs("saupbu")%>&nbsp;<%=rs("team")%></td>
								<td><%=rs("hp")%></td>
								<td><%=grade_view%></td>
								<td><%=mg_group%></td>
								<td><%=rs("reside_place")%>&nbsp;</td>
								<td><%=cost_grade_view%></td>
								<td><a href="#" onClick="pop_Window('cost_grade_mod.asp?user_id=<%=rs("user_id")%>&u_type=<%="U"%>','cost_grade_pop','scrollbars=no,width=800,height=170')">����</a></td>
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
                        <a href = "cost_grade_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="cost_grade_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
       	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                    <a href="cost_grade_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
       	<% if 	intend < total_page then %>
                        <a href="cost_grade_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[����]</a> <a href="cost_grade_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&use_yn=<%=use_yn%>&emp_yn=<%=emp_yn%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%" align="center"></td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

