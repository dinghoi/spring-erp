<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim be_pg, view_condi, condi, ck_sw, condi_sql
Dim Page, pgsize, start_page, stpage
Dim rsCount, rsMaster
Dim tot_record, total_page
Dim title_line

Dim emp_org_baldate, emp_grade_date
Dim page_cnt
Dim intstart, intend, first_page, i
Dim emp_name

be_pg = "insa_master_modify.asp"

'user_id = request.cookies("nkpmg_user")("coo_user_id")
'insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

Page = Request("page")
view_condi = Request("view_condi")
condi = Request("condi")

ck_sw = Request("ck_sw")

If ck_sw = "n" Then
	view_condi = Request.Form("view_condi")
	condi = Request.Form("condi")
Else
	view_condi = Request("view_condi")
	condi = Request("condi")
End If

If view_condi = "" Then
	condi = ""
	condi_sql = "(emp_no = '" + condi + "') AND "
End If

If view_condi = "���" Then
	condi_sql = "(emp_no = '" + condi + "') AND "
End If

If view_condi = "����" Then
	condi_sql = "(emp_name like '%" + condi + "%') AND "
End If

pgsize = 10 ' ȭ�� �� ������
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)

'Sql = "SELECT count(*) FROM emp_master where "+condi_sql+" (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM emp_master "
objBuilder.Append "WHERE "&condi_sql&" "
objBuilder.Append "	(isNull(emp_end_date) or emp_end_date = '1900-01-01') "
objBuilder.Append "	AND (emp_no < '900000') "
'Set RsCount = Server.CreateObject("ADODB.Recordset")
Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

tot_record = CInt(RsCount(0)) 'Result.RecordCount

rsCount.Close()
'Set rsCount = Nothing

If tot_record MOD pgsize = 0 Then
	total_page = Int(tot_record / pgsize) 'Result.PageCount
Else
	total_page = Int((tot_record / pgsize) + 1)
End If

'Sql = "SELECT * FROM emp_master where "+condi_sql+" (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY  emp_no,emp_name ASC limit "& stpage & "," &pgsize
objBuilder.Append "SELECT emp_no, emp_name, emp_first_date, emp_in_date, emp_company, "
objBuilder.Append "	emp_bonbu, emp_saupbu, emp_team, emp_org_name, emp_org_baldate, "
objBuilder.Append "	emp_reside_place, emp_grade, emp_grade_date, emp_position, emp_birthday "
objBuilder.Append "FROM emp_master "
objBuilder.Append "WHERE "+condi_sql+" "
objBuilder.Append "	(isNull(emp_end_date) or emp_end_date = '1900-01-01') "
objBuilder.Append "	AND (emp_no < '900000') "
objBuilder.Append "ORDER BY  emp_no,emp_name ASC "
objBuilder.Append "LIMIT "& stpage & "," & pgsize & " "

Set rsMaster = Server.CreateObject("ADODB.Recordset")
rsMaster.Open objBuilder.ToString(), Dbconn, 1
objBuilder.Clear()

title_line = " �λ�⺻ ���� "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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

			function emp_master_del(val, val2, val3, val4) {

            if (!confirm("���� �����Ͻðڽ��ϱ� ?")) return;
            var frm = document.frm;
			document.frm.emp_no.value = val;
			document.frm.emp_name.value = val2;
			document.frm.emp_company.value = val3;
			document.frm.view_condi.value = val4;

            document.frm.action = "insa_emp_master_del.asp";
            document.frm.submit();
            }
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_master_modify.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="����" <%If view_condi = "����" then %>selected<% end if %>>����</option>
                                  <option value="���" <%If view_condi = "���" then %>selected<% end if %>>���</option>
                                </select>
								<strong>���� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left; ime-mode:active" >
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="8%" >
							<col width="*" >
                            <col width="3%" >
                            <col width="3%" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>
								<th scope="col">�������</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�Ի���</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">�����Ի���</th>
								<th scope="col">�Ҽӹ߷���</th>
								<th scope="col">����ó</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <th scope="col">��ȸ</th>
                                <th colspan="2" scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsMaster.EOF

							If rsMaster("emp_org_baldate") = "1900-01-01" Then
							   emp_org_baldate = ""
							Else
							   emp_org_baldate = rsMaster("emp_org_baldate")
							End If

							If rsMaster("emp_grade_date") = "1900-01-01" Then
 							   emp_grade_date = ""
							Else
							   emp_grade_date = rsMaster("emp_grade_date")
							End If
	           			%>
							<tr>
								<td class="first"><%=rsMaster("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rsMaster("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rsMaster("emp_name")%></a>
								</td>
                                <td><%=rsMaster("emp_birthday")%>&nbsp;</td>
                                <td><%=rsMaster("emp_grade")%>&nbsp;</td>
                                <td><%=rsMaster("emp_position")%>&nbsp;</td>
                                <td><%=rsMaster("emp_in_date")%>&nbsp;</td>
                                <td><%=rsMaster("emp_org_name")%>&nbsp;</td>
                                <td><%=rsMaster("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rsMaster("emp_reside_place")%>&nbsp;</td>
                                <td class="left"><%=rsMaster("emp_company")%>-<%=rsMaster("emp_bonbu")%>-<%=rsMaster("emp_saupbu")%>-<%=rsMaster("emp_team")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_emp_master_view.asp?view_condi=<%=rsMaster("emp_company")%>&emp_no=<%=rsMaster("emp_no")%>&u_type=<%=""%>','insa_emp_modify_popup','scrollbars=yes,width=1250,height=480')">��ȸ</a></td>

                          <%
						  	 '�λ� ���� ���� ���� ����
							 If InsaMasterModYn = "Y" Then
						  %>
                                <td><a href="#" onClick="pop_Window('insa_emp_master_modify.asp?view_condi=<%=rsMaster("emp_company")%>&emp_no=<%=rsMaster("emp_no")%>&u_type=<%="U"%>','insa_emp_modify_popup','scrollbars=yes,width=1250,height=600')">����</a></td>
                          <% Else %>
                                <td>&nbsp;</td>
                          <% End If %>
                          <%
						  	'�λ� ���� ���� ���� ����
							 If InsaMasterDelYn = "Y" Then
						   %>
                              <td>
                              <a href="#" onClick="emp_master_del('<%=rsMaster("emp_no")%>', '<%=rsMaster("emp_name")%>', '<%=rsMaster("emp_company")%>', '<%=view_condi%>');return false;">����</a></td>
                         <%     Else %>
                              <td>&nbsp;</td>
                         <% End If %>
							</tr>
						<%
							rsMaster.MoveNext()
						Loop

						rsMaster.Close()
						Set rsMaster = Nothing

						DBConn.Close()
						Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
				<%

                intstart = (Int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                If intend > total_page Then
                    intend = total_page
                End If
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                  <div id="paging">
                        <a href = "insa_master_modify.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% If intstart > 1 Then %>
                        <a href="insa_master_modify.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a>
                    <% End If %>
                    <% For i = intstart To intend %>
		           		<% If i = Int(page) Then %>
                        <b>[<%=i%>]</b>
						<% Else %>
                        <a href="insa_master_modify.asp?page=<%=i%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
						<% End If %>
                    <% Next %>
				 	<% If intend < total_page Then %>
                        <a href="insa_master_modify.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_master_modify.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&condi=<%=condi%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                    <% End If %>
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

