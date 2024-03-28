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
Dim page, view_condi, condi, be_pg, condi_sql, pgsize
Dim start_page, stpage, rsCount, total_record, total_page
Dim title_line, rsEmp, str_param

page = f_Request("page")
view_condi = f_Request("view_condi")
condi = f_Request("condi")

be_pg = "/person/insa_plist_mg.asp"

If view_condi = "" Then
	view_condi = "��ü"
	condi_sql = " "
	condi = ""
End If

Select Case view_condi
	Case "�Ҽ�������"
		condi_sql = "AND emp_org_name LIKE '%"&condi&"%' "
	Case "����"
		condi_sql = "AND emp_name LIKE '%"&condi&"%' "
	Case "ȸ�纰"
		condi_sql = "AND emp_company LIKE '%"&condi&"%' "
	Case "���κ�"
		condi_sql = "AND emp_bonbu LIKE '%"&condi&"%' "
	Case "����κ�"
		condi_sql = "AND emp_saupbu LIKE '%"&condi&"%' "
	Case "����"
		condi_sql = "AND emp_team LIKE '%"&condi&"%' "
	Case "����ó ȸ�纰"
		condi_sql = "AND emp_reside_company LIKE '%"&condi&"%' "
End Select

pgsize = 10 ' ȭ�� �� ������

If page = "" Then
	page = 1
	start_page = 1
End If
stpage = Int((page - 1) * pgsize)

str_param = "&view_condi="&view_condi&"&condi="&condi

objBuilder.Append "SELECT COUNT(*) FROM emp_master "
objBuilder.Append "WHERE 1=1 "&condi_sql&" AND (ISNULL(emp_end_date) OR emp_end_date = '1900-01-01') AND (emp_no < '900000');"

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing

If total_record Mod pgsize = 0 Then
	total_page = Int(total_record / pgsize) 'Result.PageCount
Else
	total_page = Int((total_record / pgsize) + 1)
End If

objBuilder.Append "SELECT emp_no, emp_name, emp_email, emp_org_name, emp_job, "
objBuilder.Append "	emp_position, emp_extension_no, emp_hp_ddd, emp_hp_no1, emp_hp_no2, "
objBuilder.Append "	emp_company, emp_bonbu, emp_saupbu, emp_team "
objBuilder.Append "FROM emp_master "
objBuilder.Append "WHERE 1=1 "&condi_sql&" AND (isNull(emp_end_date) OR emp_end_date = '1900-01-01') AND (emp_no < '900000') "
objBuilder.Append "ORDER BY emp_no,emp_name ASC "
objBuilder.Append "LIMIT "&stpage&","&pgsize

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "���� �ּҷ� - "&view_condi
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
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

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}

			function delcheck(){
				if(form_chk(document.frm_del)){
					document.frm_del.submit();
				}
			}

			function form_chk(){
				var result = confirm('�����Ͻðڽ��ϱ�?');

				if(result){
					return true;
				}
				return false;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_plist_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_plist_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                                <select name="view_condi" id="select3" style="width:100px">
                                  <option value="��ü" <%If view_condi = "��ü" Then %>selected<%End If %>>��ü</option>
                                  <option value="�Ҽ�������" <%If view_condi = "�Ҽ�������" Then %>selected<%End If %>>�Ҽ�������</option>
                                  <option value="����" <%If view_condi = "����" Then %>selected<%End If %>>����</option>
                                  <option value="ȸ�纰" <%If view_condi = "ȸ�纰" Then %>selected<%End If %>>ȸ�纰</option>
                                  <option value="���κ�" <%If view_condi = "���κ�" Then %>selected<%End If %>>���κ�</option>
                                  <option value="����κ�" <%If view_condi = "����κ�" Then %>selected<%End If %>>����κ�</option>
                                  <option value="����" <%If view_condi = "����" Then %>selected<%End If %>>����</option>
                                  <option value="����ó ȸ�纰" <%If view_condi = "����ó ȸ�纰" Then %>selected<%End If %>>����ó ȸ�纰</option>
                                </select>
								<strong>���� : </strong>
								<input name="condi" type="text" value="<%=condi%>" style="width:150px; text-align:left"/>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"></a>
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
						Dim emp_email

						If rsEmp.EOF Or rsEmp.BOF Then
							Response.Write "<tr><td colspan='8' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						Else
							Do Until rsEmp.EOF
								emp_email = rsEmp("emp_email")&"@k-one.co.kr"
	           			%>
							<tr>
                                <td class="first"><%=rsEmp("emp_org_name")%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('/person/insa_emp_card.asp?emp_no=<%=rsEmp("emp_no")%>&emp_name=<%=rsEmp("emp_name")%>&u_type=U','insa_emp_card_pop','scrollbars=yes,width=500,height=540')"><%=rsEmp("emp_name")%></a>&nbsp;</td>
                                <td><%=rsEmp("emp_job")%>&nbsp;</td>
                                <td><%=rsEmp("emp_position")%>&nbsp;</td>
                                <td class="left"><%=emp_email%>&nbsp;</td>
                                <td><%=rsEmp("emp_extension_no")%>&nbsp;</td>
                                <td><%=rsEmp("emp_hp_ddd")%>-<%=rsEmp("emp_hp_no1")%>-<%=rsEmp("emp_hp_no2")%>&nbsp;</td>
                                <td class="left">
									<%Call EmpOrgInSaupbuText(rsEmp("emp_company"), rsEmp("emp_bonbu"), rsEmp("emp_saupbu"), rsEmp("emp_team"))%>
								</td>
							</tr>
						<%
								rsEmp.MoveNext()
							Loop
						End If
						rsEmp.Close() : Set rsEmp = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    <%
					'page navigator[����ȣ_20210720]
					Call Page_Navi(page, be_pg, str_param, total_page)

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