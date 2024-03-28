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
Dim be_pg, page, view_condi, condi, ck_sw
Dim condi_sql, pgsize, start_page, stpage
Dim title_line, total_page, rsReport
Dim page_cnt, pg_url
Dim rsCount, totRecord
Dim emp_org_baldate, emp_grade_date, emp_birthday

be_pg = "/insa/insa_report_mg.asp"

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")
'view_condi = f_Request("view_condi")

Select Case view_condi
	Case "�Ҽ�������"
		condi_sql = "AND eomt.org_name LIKE '%"&condi&"%' "
	Case "����"
		condi_sql = "AND emtt.emp_name LIKE '%"&condi&"%' "
	Case "���޺�"
		condi_sql = "AND emtt.emp_grade LIKE '%"&condi&"%' "
	Case "������"
		condi_sql = "AND emtt.emp_job LIKE '%" & condi & "%' "
	Case "��å��"
		condi_sql = "AND emtt.emp_position LIKE '%" & condi & "%' "
	Case "ȸ�纰"
		condi_sql = "AND eomt.org_company LIKE '%" & condi & "%' "
	Case "���κ�"
		condi_sql = "AND eomt.org_bonbu LIKE '%" & condi & "%' "
	Case "����κ�"
		condi_sql = "AND eomt.org_saupbu LIKE '%" & condi & "%' "
	Case "����"
		condi_sql = "AND eomt.org_team LIKE '%" & condi & "%' "
	Case "����óȸ�纰"
		condi_sql = "AND eomt.org_reside_company LIKE '%" & condi & "%' "
	Case "�Ի����ں�"
		condi_sql = "AND emp_in_date LIKE '%" & condi & "%' "
	Case Else
		view_condi = "��ü"
		condi_sql = ""
End Select

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&condi="&condi

'�� ī��Ʈ ��ȸ
objBuilder.Append "SELECT COUNT(*) FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append condi_sql

Set rsCount = Dbconn.Execute(objBuilder.ToString())
objBuilder.Clear()

totRecord = CInt(RsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_birthday, emtt.emp_no, "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, emtt.emp_in_date, "
objBuilder.Append "	emtt.emp_org_name, emtt.emp_first_date, emtt.emp_reside_place, emtt.emp_company, "
objBuilder.Append "	emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, eomt.org_name, "
objBuilder.Append "	eomt.org_reside_place, eomt.org_code "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND emtt.emp_no < '900000' "
objBuilder.Append condi_sql
objBuilder.Append "ORDER BY emp_no,emp_name ASC "
objBuilder.Append "LIMIT "& stpage & ", " &pgsize

Set rsReport = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = view_condi&" - ���� ��Ȳ "
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ� ���� �ý���</title>
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

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}
			/*
			function delcheck(){
				if(form_chk(document.frm_del)){
					document.frm_del.submit();
				}
			}

			function form_chk(){
				var result = confirm('�����Ͻðڽ��ϱ�?');

				if(result == true){
					return true;
				}
				return false;
			}*/
		</script>
	</head>
	<!--<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">-->
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_report_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px;">
                                  <option value="��ü" <%If view_condi = "��ü" Then %>selected<%End If %>>��ü</option>
                                  <option value="�Ҽ�������" <%If view_condi = "�Ҽ�������" then %>selected<% end if %>>�Ҽ�������</option>
                                  <option value="����" <%If view_condi = "����" Then %>selected<%End If %>>����</option>
                                  <option value="���޺�" <%If view_condi = "���޺�" Then %>selected<%End If %>>���޺�</option>
                                  <option value="������" <%If view_condi = "������" Then %>selected<%End If %>>������</option>
                                  <option value="��å��" <%If view_condi = "��å��" Then %>selected<%End If %>>��å��</option>
                                  <option value="ȸ�纰" <%If view_condi = "ȸ�纰" Then %>selected<%End If %>>ȸ�纰</option>
                                  <option value="���κ�" <%If view_condi = "���κ�" Then %>selected<%End If %>>���κ�</option>
                                  <option value="����κ�" <%If view_condi = "����κ�" Then %>selected<%End If %>>����κ�</option>
                                  <option value="����" <%If view_condi = "����" Then %>selected<%End If %>>����</option>
                                  <option value="����óȸ�纰" <%If view_condi = "����óȸ�纰" Then %>selected<%End If %>>����óȸ�纰</option>
                                  <option value="�Ի����ں�" <%If view_condi = "�Ի����ں�" Then %>selected<%End If %>>�Ի����ں�</option>
                                </select>
								<strong>���� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left;"/>
                                <strong> (�Ի��� ������ yyyy-mm-dd ���·� �Է�)</strong>
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
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="9%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">����</th>
								<th scope="col">�������</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�Ի���</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">�����Ի���</th>
								<th scope="col">�Ҽӹ߷���</th>
								<th scope="col">����ó</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsReport.EOF
							If rsReport("emp_org_baldate") = "1900-01-01" Then
							   emp_org_baldate = ""
							Else
							   emp_org_baldate = rsReport("emp_org_baldate")
							End If

							If rsReport("emp_grade_date") = "1900-01-01" Then
							   emp_grade_date = ""
							Else
							   emp_grade_date = rsReport("emp_grade_date")
							End If

							If rsReport("emp_birthday") = "1900-01-01" Then
							   emp_birthday = ""
							Else
							   emp_birthday = rsReport("emp_birthday")
							End If
	           			%>
							<tr>
								<td class="first"><%=rsReport("emp_no")%>&nbsp;</td>
                                <td>
									<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsReport("emp_no")%>','�λ� ��� ī��','scrollbars=yes,width=1250,height=670')"><%=rsReport("emp_name")%></a>
								</td>
                                <td><%=emp_birthday%>&nbsp;</td>
                                <td><%=rsReport("emp_grade")%>&nbsp;</td>
                                <td><%=rsReport("emp_job")%>&nbsp;</td>
                                <td><%=rsReport("emp_position")%>&nbsp;</td>
                                <td><%=rsReport("emp_in_date")%>&nbsp;</td>
                                <td><%=rsReport("org_name")%>&nbsp;</td>
                                <td><%=rsReport("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rsReport("org_reside_place")%>&nbsp;</td>
                                <td class="left">
								<%
								Call EmpOrgCodeSelect(rsReport("org_code"))
								%>
								</td>
							</tr>
						<%
							rsReport.MoveNext()
						Loop
						rsReport.close() : Set rsReport = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, totRecord, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                  	</td>
			      </tr>
				  </table>
			</form>
		</div>
	</div>
	</body>
</html>