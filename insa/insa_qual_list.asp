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
Dim view_company, title_line, condi_sql, pgsize
Dim pasize, start_page, stpage, pg_url
Dim rsCount, totRecord, rsQual
Dim end_view, rs_org
Dim qual_empno, emp_name, emp_grade, emp_job, emp_position
Dim emp_org_code, emp_org_name, page_cnt
Dim rs_emp

be_pg = "/insa/insa_qual_list.asp"

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")
view_company = f_Request("view_company")

title_line = " �ڰ��� ���� ��Ȳ "

If f_toString(view_condi, "") = "" Then
	view_company = "���̿�"
	view_condi = "��ü"
	condi_sql = ""
	condi = ""
End If

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&condi="&condi&"&view_company="&view_company

objBuilder.Append "SELECT COUNT(*) FROM emp_qual AS emqt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emqt.qual_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "AND eomt.org_company LIKE '"&view_company&"' "

If view_condi = "����óȸ��" Then
	objBuilder.Append "AND eomt.org_reside_place LIKE '%" & condi & "%' "
ElseIf view_condi = "�ڰ�����" Then
	objBuilder.Append "AND emqt.qual_type LIKE '%" & condi & "%' "
Else
	objBuilder.Append "AND emtt.emp_name LIKE '%"&condi&"%' "
End If

Set rsCount = Dbconn.Execute(objBuilder.ToString())
objBuilder.Clear()

totRecord = CInt(rsCount(0))

rsCount.Close() : Set rsCount = Nothing

objBuilder.Append "SELECT emqt.qual_empno, emqt.qual_type, emqt.qual_grade, emqt.qual_org, "
objBuilder.Append "	emqt.qual_no, emqt.qual_pass_date, emqt.qual_empno,  "
objBuilder.Append "	emtt.emp_name, emtt.emp_grade, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_org_code, emtt.emp_org_name, emtt.emp_company, "
objBuilder.Append "	eomt.org_name, eomt.org_company "
objBuilder.Append "FROM emp_qual AS emqt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emqt.qual_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE (isNull(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date = '0000-00-00') "
objBuilder.Append "	AND eomt.org_company LIKE '%"&view_company&"%' "

If view_condi = "����óȸ��" Then
	objBuilder.Append "AND eomt.org_reside_place "
ElseIf view_condi = "�ڰ�����" Then
	objBuilder.Append "AND emqt.qual_type "
Else
	objBuilder.Append "AND emtt.emp_name "
End If

objBuilder.Append "LIKE '%"&condi&"%' "

objBuilder.Append "ORDER BY emqt.qual_empno ASC "
objBuilder.Append "LIMIT "& stpage & ", " &pgsize

Set rsQual = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
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

			/*function delcheck(){
				if(form_chk(document.frm_del)){
					document.frm_del.submit();
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
				<form action="/insa/insa_qual_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>�˻�</dt>
                        <dd>
                            <p>
                               <strong>ȸ�� : </strong>
                              <%
								objBuilder.Append "SELECT org_name FROM emp_org_mst "
								objBuilder.Append "WHERE (isNull(org_end_date) OR org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = 'ȸ��' AND org_code <> '6272' "
								objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

								Set rs_org = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  %>
                                <label>
									<select name="view_company" id="view_company" type="text" style="width:150px;">
								  <%
									do until rs_org.eof
								  %>
                						<option value='<%=rs_org("org_name")%>' <%If view_company = rs_org("org_name") Then %>selected<%End If%>><%=rs_org("org_name")%></option>
								  <%
										rs_org.MoveNext()
									Loop
									rs_org.Close() : Set rs_org = Nothing
								  %>
            					</select>
                                </label>
                                <strong>���� : </strong>
                                <label>
									<select name="view_condi" id="select3" style="width:100px;">
										<option value="��ü" <%If view_condi = "��ü" Then %>selected<%End If %>>��ü</option>
										<option value="�ڰ�����" <%If view_condi = "�ڰ�����" Then %>selected<%End If %>>�ڰ�����</option>
										<option value="����óȸ��" <%If view_condi = "����óȸ��" Then %>selected<%End If %>>����óȸ��</option>
									</select>
                                </label>
								<strong>�˻��� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left;"/>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="14%" >
							<col width="6%" >
							<col width="*" >
							<col width="12%" >
							<col width="8%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="6%" >
							<col width="10%" >
							<col width="10%" >
							<col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�ڰ�����</th>
								<th scope="col">���</th>
								<th scope="col">�߱ޱ��</th>
								<th scope="col">�ڰݵ�Ϲ�ȣ</th>
								<th scope="col">�����</th>
								<th scope="col">���</th>
                                <th scope="col">����</th>
                                <th scope="col">����</th>
								<th scope="col">ȸ��</th>
								<th scope="col">�Ҽ�(����ó)</th>
								<th scope="col">��</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsQual.EOF
							qual_empno = rsQual("qual_empno")
							emp_name = rsQual("emp_name")
							emp_grade = rsQual("emp_grade")
							emp_job = rsQual("emp_job")
							emp_position = rsQual("emp_position")
							emp_org_code = rsQual("emp_org_code")
							emp_org_name = rsQual("org_name")
							emp_company = rsQual("org_company")
	           			%>
							<tr>
								<td class="first"><%=rsQual("qual_type")%>&nbsp;</td>
                                <td><%=rsQual("qual_grade")%>&nbsp;</td>
                                <td><%=rsQual("qual_org")%>&nbsp;</td>
                                <td><%=rsQual("qual_no")%>&nbsp;</td>
                                <td><%=rsQual("qual_pass_date")%>&nbsp;</td>
                                <td><%=rsQual("qual_empno")%>&nbsp;</td>
                                <td>
									<a href="#" onClick="pop_Window('/insa/insa_card00.asp?emp_no=<%=rsQual("qual_empno")%>','�λ� ��� ī��','scrollbars=yes,width=1250,height=670')"><%=emp_name%></a>
								</td>
                                <td><%=emp_job%>&nbsp;</td>
                                <td><%=emp_company%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td>
									<a href="#" onClick="pop_Window('/insa/insa_qual_view.asp?emp_no=<%=rsQual("qual_empno")%>&emp_name=<%=emp_name%>','�ڰ��� ����','scrollbars=yes,width=800,height=400')">��ȸ</a>&nbsp;
								</td>
							</tr>
						<%
							rsQual.MoveNext()
						Loop
						rsQual.close() : Set rsQual = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="/insa/insa_excel_quallist.asp?view_condi=<%=view_condi%>&condi=<%=condi%>&view_company=<%=view_company%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
					<%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, totRecord, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>
			      </tr>
				</table>
		</div>
	</div>
	</body>
</html>