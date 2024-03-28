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
Dim be_pg, page, view_condi, condi, ck_sw, condi_sql
Dim pgsize, start_page, stpage, rsCount, totRecord
Dim total_page, title_line, rsEndMem, pg_url
Dim emp_org_baldate, emp_grade_date, page_cnt

be_pg = "/insa/insa_emp_end_list.asp"

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")

title_line = " ������ ��ȸ "

If f_toString(condi, "") <> "" Then
	Select Case view_condi
		Case "���"
			condi_sql = "AND emp_no = '"&condi&"' "
		Case "����"
			condi_sql = "AND emp_name LIKE '%"&condi&"%' "
		Case Else
			condi = ""
			condi_sql = "AND emp_no = '"&condi&"' "
	End Select
End If

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&condi="&condi

objBuilder.Append "SELECT COUNT(*) FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emp_end_date > '1900-01-01' AND emp_no < '900000' "& condi_sql

Set rsCount = Dbconn.Execute(objBuilder.ToString())
objBuilder.Clear()

totRecord = cint(RsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emtt.emp_org_baldate, emtt.emp_grade_date, emtt.emp_no, emtt.emp_name, "
objBuilder.Append "	emtt.emp_birthday, emtt.emp_grade, emtt.emp_position, emtt.emp_in_date, "
objBuilder.Append "	emtt.emp_org_name, emtt.emp_first_date, emtt.emp_end_date, emtt.emp_company, "
objBuilder.Append "	emtt.emp_bonbu, emtt.emp_saupbu, emtt.emp_team, "
'objBuilder.Append "	eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team "
objBuilder.Append "	eomt.org_name, eomt.org_code "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emp_end_date > '1900-01-01' AND emp_no < '900000' "&condi_sql
objBuilder.Append "ORDER BY emtt.emp_end_date DESC, emtt.emp_no, emp_name ASC "
objBuilder.Append "LIMIT "& stpage & ", " &pgsize

Set rsEndMem = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
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
				return "0 1";
			}

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}
			/*
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
			}*/
			//-->
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/insa/insa_emp_end_list.asp" method="post" name="frm">
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
								<input type="text" name="condi" value="<%=condi%>" style="width:150px; text-align:left;"/>
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
								<th scope="col">��������</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
                                <th scope="col">��ȸ</th>
                                <th colspan="2" scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsEndMem.EOF
							If rsEndMem("emp_org_baldate") = "1900-01-01" Then
							   emp_org_baldate = ""
							Else
							   emp_org_baldate = rsEndMem("emp_org_baldate")
							End If

							If rsEndMem("emp_grade_date") = "1900-01-01" Then
							   emp_grade_date = ""
							Else
							   emp_grade_date = rsEndMem("emp_grade_date")
							End If
	           			%>
							<tr>
								<td class="first"><%=rsEndMem("emp_no")%>&nbsp;</td>
                                <td>
									<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsEndMem("emp_no")%>','�λ� ��� ī��','scrollbars=yes,width=1250,height=670')"><%=rsEndMem("emp_name")%></a>
								</td>
                                <td><%=rsEndMem("emp_birthday")%>&nbsp;</td>
                                <td><%=rsEndMem("emp_grade")%>&nbsp;</td>
                                <td><%=rsEndMem("emp_position")%>&nbsp;</td>
                                <td><%=rsEndMem("emp_in_date")%>&nbsp;</td>
                                <td><%=rsEndMem("org_name")%>&nbsp;</td>
                                <td><%=rsEndMem("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rsEndMem("emp_end_date")%>&nbsp;</td>
                                <td class="left">
								<%
								Call EmpOrgCodeSelect(rsEndMem("org_code"))
								%>
								</td>
                                <td>
									<a href="#" onClick="pop_Window('/insa/insa_emp_master_view.asp?view_condi=<%=rsEndMem("emp_company")%>&emp_no=<%=rsEndMem("emp_no")%>','insa_emp_modify_popup','scrollbars=yes,width=1250,height=480')">��ȸ</a>
								</td>
                                <td colspan="2">&nbsp;</td>
							</tr>
						<%
							rsEndMem.MoveNext()
						Loop
						rsEndMem.Close() : Set rsEndMem = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
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