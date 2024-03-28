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
Dim u_type, sch_empno, sch_seq, emp_name, sch_start_date, sch_end_date
Dim sch_school_name, sch_dept, sch_major, sch_sub_major, sch_degree
Dim sch_finish, sch_comment, curr_date, title_line, view_condi, rs_etc
Dim rsSch

u_type = Request.QueryString("u_type")
sch_empno = Request.QueryString("sch_empno")
sch_seq = Request.QueryString("sch_seq")
emp_name = Request.QueryString("emp_name")

sch_start_date = ""
sch_end_date = ""
sch_school_name = ""
sch_dept = ""
sch_major = ""
sch_sub_major = ""
sch_degree = ""
sch_finish = ""
sch_comment = ""
curr_date = Mid(CStr(Now()), 1, 10)
title_line = "�з»��� ���"

If u_type = "U" Then
	objBuilder.Append "SELECT sch_empno, sch_seq, sch_start_date, sch_end_date, sch_dept, sch_major, "
	objBuilder.Append "	sch_sub_major, sch_degree, sch_finish, sch_comment, sch_school_name "
	objBuilder.Append "FROM emp_school "
	objBuilder.Append "WHERE sch_empno = '"&sch_empno&"' AND sch_seq = '"&sch_seq&"';"

	Set rsSch = DBConn.Execute(objBuilder.ToSTring())
	objBuilder.Clear()

    sch_empno = rsSch("sch_empno")
    sch_seq = rsSch("sch_seq")
	sch_start_date = rsSch("sch_start_date")
    sch_end_date = rsSch("sch_end_date")
    sch_dept = rsSch("sch_dept")
    sch_major = rsSch("sch_major")
    sch_sub_major = rsSch("sch_sub_major")
    sch_degree = rsSch("sch_degree")
	sch_finish = rsSch("sch_finish")

    view_condi = rsSch("sch_comment")
	sch_school_name = rsSch("sch_school_name")

	rsSch.Close() : Set rsSch = Nothing

	title_line = "�з»��� ����"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ�������</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			//�Ⱓ ������
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=sch_start_date%>" );
			});

			//�Ⱓ ������
			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%=sch_end_date%>" );
			});

			function goAction(){
			   window.close();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.sch_end_date.value == ""){
					alert('�������ڸ� �Է����ּ���.');
					frm.sch_end_date.focus();
					return false;
				}

				if(document.frm.view_condi.value =="1"){
					if(document.frm.sch_high_name.value ==""){
						alert('�б����� �Է����ּ���.');
						frm.sch_high_name.focus();
						return false;
					}
				}

				if(document.frm.view_condi.value =="2"){
					if(document.frm.sch_school_name.value ==""){
						alert('�б����� �������ּ���.');
						frm.sch_school_name.focus();
						return false;
					}
				}

				if(document.frm.sch_dept.value ==""){
					alert('�а��� �Է����ּ���.');
					frm.sch_dept.focus();
					return false;
				}

				if(document.frm.sch_major.value ==""){
					alert('������ �Է����ּ���.');
					frm.sch_major.focus();
					return false;
				}

				if(document.frm.sch_finish.value ==""){
					alert('�������θ� �������ּ���.');
					frm.sch_finish.focus();
					return false;
				}

				var result='��� �Ͻðڽ��ϱ�?';

				if(result){
					return true;
				}
				return false;
			}

			function condi_view(){
				var k=0;

				for(j=0; j<2; j++){
					if(eval("document.frm.view_condi["+j+"].checked")){
						k=j+1;
					}
				}

				if(k==1){
					document.frm.sch_high_name.style.display='';
					document.frm.sch_school_name.style.display='none';
				}

				if(k == 2){
					document.frm.sch_high_name.style.display='none';
					document.frm.sch_school_name.style.display='';
				}
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
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/person/insa_school_add_save.asp" method="post" name="frm">
			<div class="gView">
			  <table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="11%" >
					<col width="22%" >
					<col width="11%" >
					<col width="22%" >
					<col width="11%" >
					<col width="*" >
				</colgroup>
				<tbody>
				<tr>
					<th style="background:#FFFFE6">���</th>
					<td class="left" bgcolor="#FFFFE6">
						<input type="text" name="sch_empno" id="sch_empno" size="14" value="<%=sch_empno%>" class="no-input" readonly/>
						<input type="hidden" name="sch_seq" value="<%=sch_seq%>"/>
					</td>
					<th style="background:#FFFFE6">����</th>
					<td colspan="3" class="left" bgcolor="#FFFFE6">
						<input type="text" name="emp_name" id="emp_name" size="14" value="<%=emp_name%>" class="no-input" readonly/>
					</td>
				</tr>
				<tr>
					<th>�Ⱓ<span style="color:red;">*</span></th>
					<td colspan="5" class="left">
						<input type="text" name="sch_start_date" value="<%=sch_start_date%>" style="width:80px;text-align:center" id="datepicker"/>
						&nbsp;-&nbsp;
						<input type="text" name="sch_end_date" value="<%=sch_end_date%>" style="width:80px;text-align:center" id="datepicker1"/>
					</td>
				</tr>
				<tr>
					<th>�б���<span style="color:red;">*</span></th>
					<td colspan="5" class="left">
						<input type="radio" name="view_condi" value="1" <%If view_condi = "1" Then %>checked<%End If %> title="����б�" style="width:30px" onClick="condi_view()"/>����б�
					<%If view_condi = "1" Then %>
						<input type="text" name="sch_high_name" id="sch_high_name" value="<%=sch_school_name%>" style="width:150px"/>
					<%Else  %>
						<input type="text" name="sch_high_name" id="sch_high_name" style="display:none; width:150px"/>
					<%End If %>
						<input type="radio" name="view_condi" value="2" <%If view_condi = "2" Then %>checked<%End If%> title="����" style="width:30px" onClick="condi_view()"/>����
					<%
					objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
					objBuilder.Append "WHERE emp_etc_type = '20' ORDER BY emp_etc_name ASC;"

					Set rs_etc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()

					If view_condi = "2" Then
					%>
						<select name="sch_school_name" id="sch_school_name" style="width:150px">
					<%
					Else
					%>
						<select name="sch_school_name" id="sch_school_name" style="display:none; width:150px">
					<%
					End If
					%>
							<option value="" <%If sch_school_name = "" Then %>selected<%End If%>>����</option>
					<%
						Do Until rs_etc.EOF
					%>
							<option value='<%=rs_etc("emp_etc_name")%>' <%If sch_school_name = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
					<%
							rs_etc.MoveNext()
						Loop
						rs_etc.Close() : Set rs_etc = Nothing
						DBConn.Close() : Set DBConn = Nothing
					%>
						</select>
					</td>
				</tr>
				<tr>
					<th>�а�<span style="color:red;">*</span></th>
					<td class="left">
						<input type="text" name="sch_dept" id="sch_dept" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=sch_dept%>"/>&nbsp;
					</td>
					<th>����<span style="color:red;">*</span></th>
					<td class="left">
						<input type="text" name="sch_major" id="sch_major" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=sch_major%>"/>&nbsp;
					</td>
					<th>������</th>
					<td class="left">
						<input type="text" name="sch_sub_major" id="sch_sub_major" style="width:130px; ime-mode:active" onKeyUp="checklength(this,30);" value="<%=sch_sub_major%>"/>&nbsp;
					</td>
				</tr>
				<tr>
					<th>����<span style="color:red;">*</span></th>
					<td class="left">
						<select name="sch_finish" id="sch_finish" value="<%=sch_finish%>" style="width:100px;">
							<option value="" <%If sch_finish = "" Then %>selected<%End If %>>����</option>
							<option value='����' <%If sch_finish = "����" Then %>selected<%End If %>>����</option>
							<option value='����' <%If sch_finish = "����" Then %>selected<%End If %>>����</option>
							<option value='����' <%If sch_finish = "����" Then %>selected<%End If %>>����</option>
						</select>
					</td>
					<th>����</th>
					<td colspan="3" class="left">
						<select name="sch_degree" id="sch_degree" value="<%=sch_degree%>" style="width:100px;">
							<option value="" <%If sch_degree = "" Then %>selected<%End If %>>����</option>
							<option value='�����л�' <%If sch_degree = "�����л�" Then %>selected<%End If %>>�����л�</option>
							<option value='�л�' <%If sch_degree = "�л�" Then %>selected<%End If %>>�л�</option>
							<option value='����' <%If sch_degree = "����" Then %>selected<%End If %>>����</option>
							<option value='�ڻ�' <%If sch_degree = "�ڻ�" Then %>selected<%End If %>>�ڻ�</option>
						</select>
					</td>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="<%If u_type = "U" Then%>����<%Else%>���<%End If%>" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
			</div>
			<input type="hidden" name="u_type" value="<%=u_type%>"/>
			</form>
		</div>
	</body>
</html>