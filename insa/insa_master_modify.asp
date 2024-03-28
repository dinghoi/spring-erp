<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim page, pgsize, start_page, stpage
Dim rsCount, rsMaster
Dim tot_record, total_page
Dim title_line, pg_url

Dim emp_org_baldate, emp_grade_date
Dim page_cnt
Dim intstart, intend, first_page, i
Dim emp_name

be_pg = "/insa/insa_master_modify.asp"

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")

title_line = " �λ�⺻ ���� "

Select Case view_condi
	Case "���"
		condi_sql = "AND emp_no = '"&condi&"' "
	Case "����"
		condi_sql = "AND emp_name LIKE '%"&condi&"%' "
	Case Else
		condi = ""
		condi_sql = "AND emp_no = '"&condi&"' "
End Select

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&condi="&condi

objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM emp_master "
objBuilder.Append "WHERE (isNull(emp_end_date) OR emp_end_date = '1900-01-01' OR emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emp_no < '900000' " & condi_sql

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

tot_record = CInt(RsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_first_date, emtt.emp_in_date, emtt.emp_company, "
objBuilder.Append "	emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, emtt.emp_org_name, emtt.emp_org_baldate, "
objBuilder.Append "	emtt.emp_reside_place, emtt.emp_grade, emtt.emp_grade_date, emtt.emp_position, emtt.emp_birthday, "
objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, eomt.org_name, eomt.org_reside_place, "
objBuilder.Append "	eomt.org_code "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emp_end_date) OR emp_end_date = '1900-01-01' OR emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emp_no < '900000' "&condi_sql
objBuilder.Append "ORDER BY  emp_no,emp_name ASC "
objBuilder.Append "LIMIT "& stpage & "," & pgsize & " "

Set rsMaster = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
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

			function frmcheck(){
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}

			function delcheck(){
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
			}

			function emp_master_del(val, val2, val3, val4){
				var frm = document.frm;

				if (!confirm("���� �����Ͻðڽ��ϱ� ?")) return;

				document.frm.emp_no.value = val;
				document.frm.emp_name.value = val2;
				document.frm.emp_company.value = val3;
				document.frm.view_condi.value = val4;

				document.frm.action = "/insa/insa_emp_master_del.asp";
				document.frm.submit();
            }
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_master_modify.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="����" <%If view_condi = "����" Then %>selected<%End If %>>����</option>
                                  <option value="���" <%If view_condi = "���" Then %>selected<%End If %>>���</option>
                                </select>
								<strong>���� : </strong>
								<input type="text" name="condi" value="<%=condi%>" style="width:150px; text-align:left; ime-mode:active;"/>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"/></a>
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
						If rsMaster.EOF Or rsMaster.BOF Then
							Response.Write "<tr><td colspan='13' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						Else
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
									<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsMaster("emp_no")%>','�λ� ���ī��','scrollbars=yes,width=1250,height=670')"><%=rsMaster("emp_name")%></a>
								</td>
                                <td><%=rsMaster("emp_birthday")%>&nbsp;</td>
                                <td><%=rsMaster("emp_grade")%>&nbsp;</td>
                                <td><%=rsMaster("emp_position")%>&nbsp;</td>
                                <td><%=rsMaster("emp_in_date")%>&nbsp;</td>
                                <td><%=rsMaster("org_name")%>&nbsp;</td>
                                <td><%=rsMaster("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rsMaster("org_reside_place")%>&nbsp;</td>
                                <td class="left">
								<%
									Call EmpOrgCodeSelect(rsMaster("org_code"))
								%>(<%=rsMaster("org_code")%>)
								</td>
                                <td>
                                <a href="#" onClick="pop_Window('/insa/insa_emp_master_view.asp?view_condi=<%=rsMaster("emp_company")%>&emp_no=<%=rsMaster("emp_no")%>&u_type=<%=""%>','�λ�⺻���� ��ȸ','scrollbars=yes,width=1250,height=500')">��ȸ</a></td>

                          <%
						  	 '�λ� ���� ���� ���� ����
							 If InsaMasterModYn = "Y" Then
						  %>
                                <td><a href="#" onClick="pop_Window('/insa/insa_emp_master_modify.asp?view_condi=<%=rsMaster("emp_company")%>&emp_no=<%=rsMaster("emp_no")%>&u_type=<%="U"%>','�λ�⺻���� ����','scrollbars=yes,width=1250,height=610')">����</a></td>
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
							rsMaster.Close() : Set rsMaster = Nothing
						End If
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, tot_record, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
				</table>
				<input type="hidden" name="emp_no" value="<%=emp_no%>"/>
				<input type="hidden" name="emp_name" value="<%=emp_name%>"/>
				<input type="hidden" name="emp_company" value="<%=emp_company%>"/>
			</form>
		</div>
	</div>
	</body>
</html>