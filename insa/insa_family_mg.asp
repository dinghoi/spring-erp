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
Dim rsFamily
Dim family_empno, family_person2
Dim emp_name
Dim emp_org_code, emp_org_name, family_seq, family_yn
Dim rsEmp

view_condi = f_Request("view_condi")
owner_view = f_Request("owner_view")

title_line = " ���� ���� "

If view_condi = "" Then
	owner_view = "T"
End If

objBuilder.Append "SELECT emft.family_empno, emft.family_person2, emft.family_rel, emft.family_name, emft.family_birthday, "
objBuilder.Append "	emft.family_birthday_id, emft.family_job, emft.family_tel_ddd, emft.family_tel_no1, "
objBuilder.Append "	emft.family_tel_no2, emft.family_person1, emft.family_live, emft.family_seq, "
objBuilder.Append "	emtt.emp_name, emtt.emp_org_code, eomt.org_name "
objBuilder.Append "FROM emp_family AS emft "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emft.family_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "
objBuilder.Append "WHERE 1=1 "

If owner_view = "C" Then
	objBuilder.Append "AND emtt.emp_name LIKE '%"&view_condi&"%' "
Else
	objBuilder.Append "AND emft.family_empno = '"&view_condi&"' "
End If
objBuilder.Append "ORDER BY emft.family_empno, emft.family_seq ASC;"

Set rsFamily = DBConn.Execute(objBuilder.ToString())
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
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.view_condi.value == ""){
					alert("������ �Է��Ͻñ� �ٶ��ϴ�");
					return false;
				}
				return true;
			}

			function family_del(val, val2, val3, val4){
				if(!confirm("���� �����Ͻðڽ��ϱ�?")) return;

				var frm = document.frm;

				document.frm.family_empno.value = val;
				document.frm.family_seq.value = val2;
				document.frm.family_name.value = val3;
				document.frm.owner_view.value = val4;

				document.frm.action = "/insa/insa_family_del.asp";
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
				<form action="/insa/insa_family_mg.asp" method="post" name="frm">
				<input type="hidden" name="family_empno" id="family_empno"/>
				<input type="hidden" name="family_seq" id="family_seq"/>
				<input type="hidden" name="family_name" id="family_name"/>
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
                            <col width="6%" >
							<col width="6%" >
							<col width="10%" >
                            <col width="*" >
                            <col width="12%" >
                            <col width="12%" >
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
                                <th>����</th>
                                <th>����<br>����</th>
                                <th>�������</th>
                                <th>����</th>
                                <th>��ȭ��ȣ</th>
                                <th>�ֹι�ȣ</th>
                                <th>���ſ���</th>
                                <th>����</th>
                                <th>����</th>
                                <th>���</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsFamily.EOF Or rsFamily.BOF Then
							family_yn = "N"
							Response.Write "<tr><td colspan='13' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						Else
							Do Until rsFamily.EOF
								family_empno = rsFamily("family_empno")

								If f_toString(rsFamily("family_person2"), "") = "" Then
									family_person2 = rsFamily("family_person2")
								Else
									family_person2 = "*******"
								End If

								emp_name = rsFamily("emp_name")
								emp_org_code = rsFamily("emp_org_code")
								emp_org_name = rsFamily("org_name")
							%>
								<tr>
									<td><%=rsFamily("family_empno")%>&nbsp;</td>
									<td><%=emp_name%>&nbsp;</td>
									<td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
									<td><%=rsFamily("family_rel")%>&nbsp;</td>
									<td ><%=rsFamily("family_name")%>&nbsp;</td>
									<td class="left"><%=rsFamily("family_birthday")%>&nbsp;(<%=rsFamily("family_birthday_id")%>)&nbsp;</td>
									<td class="left"><%=rsFamily("family_job")%>&nbsp;</td>
									<td ><%=rsFamily("family_tel_ddd")%>-<%=rsFamily("family_tel_no1")%>-<%=rsFamily("family_tel_no2")%>&nbsp;</td>
									<td ><%=rsFamily("family_person1")%>-<%=family_person2%>&nbsp;</td>
									<td ><%=rsFamily("family_live")%>&nbsp;</td>
									<td >
										<a href="#" onClick="pop_Window('/insa/insa_family_add.asp?family_empno=<%=rsFamily("family_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%=""%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')">���</a>
									</td>
									<td>
										<a href="#" onClick="pop_Window('/insa/insa_family_add.asp?family_empno=<%=rsFamily("family_empno")%>&family_seq=<%=rsFamily("family_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%="U"%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')">����</a>
									</td>
									<td>
									<%If insa_grade = "0" Then %>
										<a href="#" onClick="family_del('<%=rsFamily("family_empno")%>', '<%=rsFamily("family_seq")%>', '<%=emp_name%>', '<%=owner_view%>');return false;">����</a>
									<%End If %>
									</td>
								</tr>
							<%
									rsFamily.MoveNext()
								Loop
							End If
							rsFamily.Close() : Set rsFamily = Nothing

							%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<div class="btnRight">
							<%
							'��ϵ� ���������� ���� ���
							If owner_view = "T" And f_toString(view_condi, "") <> "" And family_yn = "N" Then
								objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no = '"&view_condi&"';"

								Set rsEmp = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If Not rsEmp.EOF Then
							%>
								<a href="#" onClick="pop_Window('/insa/insa_family_add.asp?family_empno=<%=view_condi%>&emp_name=<%=rsEmp("emp_name")%>','�������','scrollbars=yes,width=750,height=400')" class="btnType04">�������</a>
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