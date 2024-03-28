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
Dim rs_emp, rsSch, arrSch
Dim title_line

in_empno = f_Request("in_empno")

If f_toString(in_empno, "") <> "" Then
	objBuilder.Append "SELECT emp_name FROM emp_master "
	objBuilder.Append "WHERE emp_no = '"&in_empno&"';"

	Set rs_emp = DbConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	in_name = rs_emp("emp_name")
	rs_emp.Close() : Set rs_emp = Nothing
Else
	in_empno = emp_no
	in_name = user_name
End If

objBuilder.Append "SELECT sch_start_date, sch_end_date, sch_school_name, sch_dept, sch_major, sch_sub_major, "
objBuilder.Append "	sch_degree, sch_finish, sch_empno, sch_seq "
objBuilder.Append "FROM emp_school "
objBuilder.Append "WHERE sch_empno = '"&in_empno&"' "
objBuilder.Append "ORDER BY sch_empno, sch_seq ASC "

Set rsSch = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsSch.EOF Then
	arrSch = rsSch.getRows()
End If
rsSch.Close() : Set rsSch = Nothing
DBConn.Close() : Set DBConn = Nothing

title_line = "�з� ����"
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
				return "0 1";
			}

			function goAction(){
			   window.close();
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
		</script>
		<style type="text/css">
			.no-input{
				color:gray;
				background-color:#E0E0E0;
				border:1px solid #999999;
			}
		</style>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="/person/insa_individual_school.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
                        <dd>
                            <p>
							<strong>��� : </strong>
							<label>
								<input type="text" name="in_empno" id="in_empno" value="<%=in_empno%>" style="width:80px;" class="no-input" readonly/>
							</label>
                            <strong>���� : </strong>
							<label>
								<input type="text" name="in_name" id="in_name" value="<%=in_name%>" style="width:80px;" class="no-input" readonly/>
							</label>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="5%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th colspan="3">�Ⱓ</th>
                                <th colspan="2">�б���</th>
                                <th colspan="2">�а�</th>
                                <th colspan="2">����</th>
                                <th >������</th>
                                <th >����</th>
                                <th>����</th>
                                <th>����</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						Dim i, sch_start_date, sch_end_date, sch_school_name, sch_dept, sch_major, sch_sub_major
						Dim sch_degree, sch_finish, sch_empno, sch_seq

						If IsArray(arrSch) Then
							For i = LBound(arrSch) To UBound(arrSch, 2)
								sch_start_date = arrSch(0, i)
								sch_end_date = arrSch(1, i)
								sch_school_name = arrSch(2, i)
								sch_dept = arrSch(3, i)
								sch_major = arrSch(4, i)
								sch_sub_major = arrSch(5, i)
								sch_degree = arrSch(6, i)
								sch_finish = arrSch(7, i)
								sch_empno = arrSch(8, i)
								sch_seq = arrSch(9, i)
						%>
							<tr>
                              <td colspan="3" ><%=sch_start_date%>��<%=sch_end_date%>&nbsp;</td>
                              <td colspan="2" ><%=sch_school_name%>&nbsp;</td>
                              <td colspan="2" ><%=sch_dept%>&nbsp;</td>
                              <td colspan="2" ><%=sch_major%>&nbsp;</td>
                              <td ><%=sch_sub_major%>&nbsp;</td>
                              <td ><%=sch_degree%>&nbsp;</td>
                              <td ><%=sch_finish%>&nbsp;</td>
							  <td><a href="#" onClick="pop_Window('/person/insa_school_add.asp?sch_empno=<%=sch_empno%>&sch_seq=<%=sch_seq%>&emp_name=<%=in_name%>&u_type=U','�з� ����','scrollbars=yes,width=750,height=300')">����</a></td>
							</tr>
						<%
							Next
						Else
							Response.Write "<tr><td colspan='13' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('/person/insa_school_add.asp?sch_empno=<%=in_empno%>&emp_name=<%=in_name%>','�з� ����','scrollbars=yes,width=750,height=300')" class="btnType04">�зµ��</a>
					</div>
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="sch_empno" value="<%=in_empno%>"/>
			</form>
		</div>
	</div>
	</body>
</html>