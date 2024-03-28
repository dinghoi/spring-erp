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
Dim view_condi, owner_view, ck_sw, title_line
Dim qual_empno, qual_seq, qual_name
Dim rsQual, qual_yn
Dim rs_emp, emp_name, emp_bonbu, emp_saupbu, emp_team
Dim emp_org_code, emp_org_name, qual_pay_id, rsEmp

view_condi = f_Request("view_condi")
owner_view = f_Request("owner_view")

title_line = " �ڰ� ���� "

If view_condi = "" Then
	owner_view = "T"
End If

objBuilder.Append "SELECT emqt.qual_empno, emqt.qual_pay_id, emqt.qual_type, emqt.qual_grade, "
objBuilder.Append "	emqt.qual_pass_date, emqt.qual_org, emqt.qual_no, emqt.qual_passport, emqt.qual_seq, "
objBuilder.Append "	emtt.emp_name, emtt.emp_org_code, eomt.org_name "
objBuilder.Append "FROM emp_qual AS emqt "
objBuilder.Append "INNER JOIN emp_master AS emtt ON emqt.qual_empno = emtt.emp_no "
objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emtt.emp_org_code = eomt.org_code "

If owner_view = "C" Then
	objBuilder.Append "WHERE emtt.emp_name LIKE '%"&view_condi&"%' "
Else
	objBuilder.Append "WHERE emqt.qual_empno = '"&view_condi&"';"
End If

Set rsQual = DBConn.Execute(objBuilder.ToString())
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

			function qual_del(val, val2, val3, val4){
				if (!confirm("���� �����Ͻðڽ��ϱ� ?")) return;

					var frm = document.frm;

					document.frm.qual_empno.value = val;
					document.frm.qual_seq.value = val2;
					document.frm.qual_name.value = val3;
					document.frm.owner_view.value = val4;

					document.frm.action = "/insa/insa_qual_del.asp";
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
				<form action="/insa/insa_qual_mg.asp" method="post" name="frm">
					<input type="hidden" name="qual_empno" value="<%=qual_empno%>"/>
					<input type="hidden" name="qual_seq" value="<%=qual_seq%>"/>
					<input type="hidden" name="qual_name" value="<%=qual_name%>"/>
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
                            <col width="11%" >
                            <col width="15%" >
							<col width="4%" >
							<col width="6%" >
							<col width="15%" >
							<col width="*" >
							<col width="10%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>���</th>
                                <th>����</th>
                                <th>�Ҽ�</th>
                                <th>�ڰ��� ����</th>
                                <th>���</th>
                                <th>�հݳ����</th>
                                <th>�߱� �����</th>
                                <th>�ڰ� ��Ϲ�ȣ</th>
                                <th>��¼�øNo.</th>
                                <th>����</th>
                                <th>�ڰ�</th>
                                <th>����</th>
                                <th>���</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						If rsQual.EOF Or rsQual.BOF Then
							qual_yn = "N"	'��»��� ��� ����
							Response.Write "<tr><td colspan='13' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						Else
							Do Until rsQual.EOF
								qual_empno = rsQual("qual_empno")
								emp_name = rsQual("emp_name")
                                emp_org_code = rsQual("emp_org_code")
                                emp_org_name = rsQual("org_name")

								qual_pay_id = ""

								If rsQual("qual_pay_id") = "Y" Then
									qual_pay_id = "����"
							    Else
									qual_pay_id = "����"
								End If
						%>
							<tr>
								<td><%=rsQual("qual_empno")%>&nbsp;</td>
								<td><%=emp_name%>&nbsp;</td>
								<td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
								<td><%=rsQual("qual_type")%>&nbsp;</td>
								<td><%=rsQual("qual_grade")%>&nbsp;</td>
								<td><%=rsQual("qual_pass_date")%>&nbsp;</td>
								<td><%=rsQual("qual_org")%>&nbsp;</td>
								<td class="left"><%=rsQual("qual_no")%>&nbsp;</td>
								<td class="left"><%=rsQual("qual_passport")%>&nbsp;</td>
								<td><%=qual_pay_id%>&nbsp;</td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_qual_add.asp?qual_empno=<%=rsQual("qual_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>','�ڰ������� ���','scrollbars=yes,width=750,height=300')">���</a>
								</td>
								<td>
									<a href="#" onClick="pop_Window('/insa/insa_qual_add.asp?qual_empno=<%=rsQual("qual_empno")%>&qual_seq=<%=rsQual("qual_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=U','�ڰ������� ����','scrollbars=yes,width=750,height=300')">����</a>
								</td>
								<%If insa_grade = "0" Then %>
								<td>
									<a href="#" onClick="qual_del('<%=rsQual("qual_empno")%>', '<%=rsQual("qual_seq")%>', '<%=emp_name%>', '<%=owner_view%>');return false;">����</a>
								</td>
								<%End If %>
							</tr>
						<%
								rsQual.MoveNext()
							Loop
						End If
						rsQual.close() : Set rsQual = Nothing
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
					<tr>
				    <td>
					<div class="btnRight">
					<%
					If owner_view = "T" And f_toString(view_condi, "") <> "" And qual_yn = "N" Then
						objBuilder.Append "SELECT emp_name FROM emp_master WHERE emp_no = '"&view_condi&"';"

						Set rsEmp = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsEmp.EOF Then
				    %>
						<a href="#" onClick="pop_Window('/insa/insa_qual_add.asp?qual_empno=<%=view_condi%>&emp_name=<%=rsEmp("emp_name")%>','insa_qual_add2_pop','scrollbars=yes,width=750,height=300')" class="btnType04">�ڰ� ���</a>
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

