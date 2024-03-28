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
Dim cfm_use, cfm_use_dept, cfm_comment
Dim rsEmp, title_line

cfm_use = ""
cfm_use_dept = ""
cfm_comment = ""

objBuilder.Append "select emtt.emp_no, emtt.emp_name, emtt.emp_job, emtt.emp_position, "
objBuilder.Append "	emtt.emp_in_date, emtt.emp_birthday, emtt.emp_company, "
objBuilder.Append "	emtt.emp_org_name, eomt.org_name, eomt.org_company, eomt.org_bonbu, "
objBuilder.Append "	eomt.org_saupbu, eomt.org_team "
objBuilder.Append "FROM emp_master AS emtt "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE emtt.emp_no < '900000' "
objBuilder.Append "	AND (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date  = '0000-00-00') "
objBuilder.Append "	AND emp_no = '"&user_id&"' "

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = "������ ��û/�߱�"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ�������</title>
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
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.in_empno.value == ""){
					alert ("����� �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}

            function s_sinchung(val, val2, val3, val4, val5){
				var frm = document.frm;

				document.frm.in_empno.value = val;
				document.frm.in_name.value = val2;

				if(document.getElementById(val3).value == ""){
					alert("��û �뵵�� �������ּ���!"); return;
				}

				if(document.getElementById(val4).value == ""){
					alert("���ó�� �Է��Ͻʽÿ�!"); return;
				}

				var result = confirm("���������� ��û�Ͻðڽ��ϱ�?");

				if(result){
					document.frm.cfm_use.value = document.getElementById(val3).value;
					document.frm.action = "/person/insa_certificate_print.asp";
					document.frm.submit();
				}
				return false;
			}

			function s_sinchung2(val, val2, val3, val4, val5){
				var frm = document.frm;

				document.frm.in_empno.value = val;
				document.frm.in_name.value = val2;

				if(document.getElementById(val3).value == ""){
					alert("��û �뵵�� �������ּ���!"); return;
				}

				if(document.getElementById(val4).value == ""){
					alert("���ó�� �Է��Ͻʽÿ�!"); return;
				}

				var result = confirm("��������� ��û�Ͻðڽ��ϱ�?");

				if(result){
					document.frm.cfm_use.value = document.getElementById(val3).value;
					document.frm.action = "/person/insa_certificate_career.asp";
					document.frm.submit();
				}
				return false;
            }
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psawo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_individual_confirm.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
                        <dd>
							<strong>��� : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=user_id%>" style="width:80px;" class="no-input" readonly/>
								</label>
                            <strong>���� : </strong>
							<label>
								<input name="in_name" type="text" id="in_name" value="<%=user_name%>" style="width:80px;" class="no-input" readonly/>
							</label>
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
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="*" >
                            <col width="6%" >
                            <col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�Ի���</th>
								<th scope="col">�������</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">�Ҽ�</th>
								<th scope="col" style="background:#FFC">�뵵</th>
								<th scope="col" style="background:#FFC">���ó</th>
                                <th scope="col" style="background:#FFC">���</th>
                                <th scope="col" style="background:#FFC">����</th>
                                <th scope="col" style="background:#FFC">���</th>
							</tr>
						</thead>
						<tbody>
						<%
						Do Until rsEmp.EOF
	           			%>
							<tr>
								<td class="first"><%=rsEmp("emp_no")%>&nbsp;</td>
                                <td><%=rsEmp("emp_name")%>&nbsp;</td>
                                <td><%=rsEmp("emp_job")%>&nbsp;</td>
                                <td><%=rsEmp("emp_position")%>&nbsp;</td>
                                <td><%=rsEmp("emp_in_date")%>&nbsp;</td>
                                <td><%=rsEmp("emp_birthday")%>&nbsp;</td>
                                <td><%=rsEmp("org_company")%>&nbsp;</td>
                                <td><%=rsEmp("org_name")%>&nbsp;</td>
                                <td class="left">
                                <select name="cfm_use" id="cfm_use" value="<%=cfm_use%>" style="width:110px">
			            	        <option value="" <%If cfm_use = "" Then %>selected<%End If %>>����</option>
				                    <option value='�����' <%If cfm_use = "�����" Then %>selected<%End If %>>�����</option>
                                    <option value='������' <%If cfm_use = "������" Then %>selected<%End If %>>������</option>
                                    <option value='�������������' <%If cfm_use = "�������������" Then %>selected<%End If %>>�������������</option>
                                    <option value='������������' <%If cfm_use = "������������" Then %>selected<%End If %>>������������</option>
                                    <option value='�����������' <%If cfm_use = "�����������" Then %>selected<%End If %>>�����������</option>
                                    <option value='���������' <%If cfm_use = "���������" Then %>selected<%End If %>>���������</option>
                                    <option value='ȸ�������' <%If cfm_use = "ȸ�������" Then %>selected<%End If %>>ȸ�������</option>
                                    <option value='���ڹ߱޿�' <%If cfm_use = "���ڹ߱޿�" Then %>selected<%End If %>>���ڹ߱޿�</option>
                                    <option value='�����' <%If cfm_use = "�����" Then %>selected<%End If %>>�����</option>
                                    <option value='�뵿��(û)�����' <%If cfm_use = "�뵿��(û)�����" Then %>selected<%End If %>>�뵿��(û)�����</option>
                                    <option value='���Ӱ���Ȯ�ο�' <%If cfm_use = "���Ӱ���Ȯ�ο�" Then %>selected<%End If %>>���Ӱ���Ȯ�ο�</option>
                                    <option value='������������' <%If cfm_use = "������������" Then %>selected<%End If %>>������������</option>
                                    <option value='��ȸ�����' <%If cfm_use = "��ȸ�����" Then %>selected<%End If %>>��ȸ�����</option>
                                    <option value='���Ȯ�ο�' <%If cfm_use = "���Ȯ�ο�" Then %>selected<%End If %>>���Ȯ�ο�</option>
                                    <option value='���������' <%If cfm_use = "���������" Then %>selected<%End If %>>���������</option>
                                    <option value='�б������' <%If cfm_use = "�б������" Then %>selected<%End If %>>�б������</option>
                                    <option value='����������' <%If cfm_use = "����������" Then %>selected<%End If %>>����������</option>
                                    <option value='ī��������' <%If cfm_use = "ī��������" Then %>selected<%End If %>>ī��������</option>
                                    <option value='���ǻ������' <%If cfm_use = "���ǻ������" Then %>selected<%End If %>>���ǻ������</option>
                                </select>
                                </td>
                                <td class="left">
									<input type="text" name="cfm_use_dept" id="cfm_use_dept" style="width:100px" onKeyUp="checklength(this,30)" value="<%=cfm_use_dept%>"/>
                                </td>
                                <td class="left">
									<input type="text" name="cfm_comment" id="cfm_comment" style="width:170px" onKeyUp="checklength(this,50)" value="<%=cfm_comment%>"/>
                                </td>
                                <td>
									<input type="image" name="rptCert$ctl00$btnRequest" id="rptCert_ctl00_btnRequest" src="/image/b_certifi.jpg" alt="�������� ��û" onclick="s_sinchung('<%=rsEmp("emp_no")%>','<%=rsEmp("emp_name")%>', 'cfm_use', 'cfm_use_dept', 'cfm_comment');return false;" style="border-width:0px;" />
                                </td>
                                <%If insa_grade = "0" Then %>
                                <td>
									<input type="image" name="rptCert$ctl01$btnRequest" id="rptCert_ctl01_btnRequest" src="/image/b_certifi.jpg" alt="������� ��û" onclick="s_sinchung2('<%=rsEmp("emp_no")%>','<%=rsEmp("emp_name")%>', 'cfm_use', 'cfm_use_dept', 'cfm_comment');return false;" style="border-width:0px;" />
                                </td>
                                <%End If %>
							</tr>
						<%
							rsEmp.MoveNext()
						Loop
						rsEmp.Close() : Set rsEmp = Nothing
						DBConn.Close() : Set DBConn = Nothing
						%>
						</tbody>
					</table>
				</div>
                  <input type="hidden" name="emp_empno" value="<%=user_id%>"/>

		</div>
	</div>
	</body>
</html>