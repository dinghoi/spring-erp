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
Dim view_condi, owner_view, title_line
Dim from_date, to_date, sch_seq, sch_emp_name
Dim rsSch
Dim sch_empno, emp_name
Dim emp_org_code, emp_org_name, school_yn, rsEmp

view_condi = f_Request("view_condi")
owner_view = f_Request("owner_view")

title_line = " �з� ���� "

If view_condi = "" Then
	owner_view = "T"
End If

objBuilder.Append "SELECT emst.sch_empno, emst.sch_start_date, emst.sch_end_date, emst.sch_school_name, "
objBuilder.Append "	emst.sch_dept, emst.sch_major, emst.sch_sub_major, emst.sch_degree, emst.sch_finish, "
objBuilder.Append "	emst.sch_seq, emtt.emp_name, emtt.emp_org_code, eomt.org_name "
objBuilder.Append "FROM emp_school AS emst "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emst.sch_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

If owner_view = "C" Then
	objBuilder.Append "WHERE emtt.emp_name LIKE '%" & view_condi & "%' "
Else
	objBuilder.Append "WHERE emst.sch_empno = '"&view_condi&"' "
End If

objBuilder.Append "ORDER BY emst.sch_empno, emst.sch_seq ASC;"

Set rsSch = DBConn.Execute(objBuilder.ToString())
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
				return "1 1";
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if(document.frm.view_condi.value == ""){
					alert ("������ �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}

			function school_del(val, val2, val3, val4){

				if (!confirm("���� �����Ͻðڽ��ϱ� ?")) return;

				var frm = document.frm;

				document.frm.sch_empno.value = val;
				document.frm.sch_seq.value = val2;
				document.frm.sch_emp_name.value = val3;
				document.frm.owner_view.value = val4;

				document.frm.action = "/insa/insa_school_del.asp";
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
				<form action="/insa/insa_school_mg.asp" method="post" name="frm">
					<input type="hidden" name="sch_empno" value="<%=sch_empno%>">
					<input type="hidden" name="sch_seq" value="<%=sch_seq%>">
					<input type="hidden" name="sch_emp_name" value="<%=sch_emp_name%>">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>������ �˻���</dt>
                        <dd>
                            <p>
                                <label>
									<input type="radio" name="owner_view" value="T" <%If owner_view = "T" Then %>checked<%End If %> style="width:25px;"/>���
									<input type="radio" name="owner_view" value="C" <%If owner_view = "C" Then %>checked<%End If %> style="width:25px;"/>����
                                </label>
							<strong>���� : </strong>
								<label>
        							<input type="text" name="view_condi" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left;"/>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="�˻�"/></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
                            <col width="6%" >
                            <col width="12%" >
                            <col width="12%" >
							<col width="*" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>���</th>
                                <th>����</th>
                                <th>�Ҽ�</th>
                                <th>�Ⱓ</th>
                                <th>�б���</th>
                                <th>�а�</th>
                                <th>����</th>
                                <th>������</th>
                                <th>����</th>
                                <th>����</th>
                                <th>�з�</th>
                                <th>����</th>
                                <th>���</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsSch.EOF Or rsSch.BOF Then
							school_yn = "N"	'�з»��� ��� ����
							Response.Write "<tr><td colspan='13' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						Else
							Do Until rsSch.EOF
								sch_empno = rsSch("sch_empno")
								emp_name = rsSch("emp_name")
								emp_org_code = rsSch("emp_org_code")
								emp_org_name = rsSch("org_name")
							%>
							<tr>
								<td><%=rsSch("sch_empno")%>&nbsp;</td>
								<td><%=emp_name%>&nbsp;</td>
								<td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
								<td><%=rsSch("sch_start_date")%>��<%=rsSch("sch_end_date")%>&nbsp;</td>
								<td><%=rsSch("sch_school_name")%>&nbsp;</td>
								<td><%=rsSch("sch_dept")%>&nbsp;</td>
								<td><%=rsSch("sch_major")%>&nbsp;</td>
								<td><%=rsSch("sch_sub_major")%>&nbsp;</td>
								<td><%=rsSch("sch_degree")%>&nbsp;</td>
								<td><%=rsSch("sch_finish")%>&nbsp;</td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_school_add.asp?sch_empno=<%=rsSch("sch_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>','�з»��� ���','scrollbars=yes,width=750,height=300')">���</a>
								</td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_school_add.asp?sch_empno=<%=rsSch("sch_empno")%>&sch_seq=<%=rsSch("sch_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=U','�з»��� ����','scrollbars=yes,width=750,height=300')">����</a>
								</td>
							 <%If insa_grade = "0" Then %>
								<td>
									<a href="#" onClick="school_del('<%=rsSch("sch_empno")%>', '<%=rsSch("sch_seq")%>', '<%=emp_name%>', '<%=owner_view%>');return false;">����</a>
								</td>
							 <%End If %>

							</tr>
							<%
								rsSch.MoveNext()
							Loop
						End If
						rsSch.close() : Set rsSch = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<div class="btnRight">
							<%'��ϵ� �з� ������ ���� ���
							If owner_view = "T" And f_toString(view_condi, "") <> "" And school_yn = "N" Then
								objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no = '"&view_condi&"';"

								Set rsEmp = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If Not rsEmp.EOF Then
							%>
								<a href="#" onClick="pop_Window('/insa/insa_school_add.asp?sch_empno=<%=view_condi%>&emp_name=<%=rsEmp("emp_name")%>','�зµ��','scrollbars=yes,width=750,height=300')" class="btnType04">�зµ��</a>
							<%
								End If
								rsEmp.Close() : Set rsEmp = Nothing
							End If
							DBConn.Close() : Set DBConn = Nothing
							%>
							</div>
						</td>
					</tr>
				</table>
			</form>
		</div>
	</div>
	</body>
</html>