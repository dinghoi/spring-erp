<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon_db.asp" -->
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
Dim lang_seq, lang_id, lang_id_type
Dim lang_point, lang_grade, pang_get_date, curr_date, title_line, lang_get_date
Dim rsLng, rs_etc, rsEtc

curr_date = Mid(CStr(Now()), 1, 10)

title_line = "���дɷ� ���"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>ȸ�� ����</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "" );
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
				if(document.frm.lang_id.value == ""){
					alert('���б����� �����ϼ���');
					frm.lang_id.focus();
					return false;
				}

				if(document.frm.lang_id_type == ""){
					alert('���������� �����ϼ���');
					frm.lang_id_type.focus();
					return false;
				}

				if(document.frm.lang_grade.value == ""){
					alert('�޼��� �Է��ϼ���');
					frm.lang_grade.focus();
					return false;
				}

				if(document.frm.lang_point.value == ""){
					alert('������ �Է��ϼ���');
					frm.lang_point.focus();
					return false;
				}

				if(document.frm.lang_get_date.value == ""){
					alert('������� �Է��ϼ���');
					frm.lang_get_date.focus();
					return false;
				}

				if(document.frm.lang_get_date.value > document.frm.curr_date.value){
					alert('������� �����Ϻ��� �����ϴ�');
					frm.lang_get_date.focus();
					return false;
				}

				var result = confirm('��� �Ͻðڽ��ϱ�?');

				if(result){
					return true;
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
	<body>
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
			<form action="/member/member_language_proc.asp" method="post" name="frm">
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
					<th style="background:#FFFFE6">����</th>
					<td colspan="3" class="left" bgcolor="#FFFFE6">
						<input type="text" name="m_name" id="m_name" value="<%=m_name%>" size="14" class="no-input" readonly/>
					</td>
				</tr>
				<tr>
					<th>���б���<span style="color:red;">*</span></th>
					<td colspan="2" class="left">
					<%
					objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
					objBuilder.Append "WHERE emp_etc_type = '08' ORDER BY emp_etc_code ASC;"

					Set rs_etc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()
					%>
						<select name="lang_id" id="lang_id" style="width:90px">
							<option value="">����</option>
					<%
					Do until rs_etc.EOF
					%>
							<option value='<%=rs_etc("emp_etc_name")%>'><%=rs_etc("emp_etc_name")%></option>
					<%
						rs_etc.MoveNext()
					Loop
					rs_etc.Close() : Set rs_etc = Nothing
					%>
				  </select>
				  </td>
				</tr>
				<tr>
					<th>��������<span style="color:red;">*</span></th>
					<td class="left">
					<%
					objBuilder.Append "SELECT emp_etc_name  FROM emp_etc_code "
					objBuilder.Append "WHERE emp_etc_type = '09' ORDER BY emp_etc_code ASC;"

					Set rsEtc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()
					%>
						<select name="lang_id_type" id="lang_id_type" style="width:90px">
							<option value="">����</option>
					<%
					Do Until rsEtc.EOF
					%>
								<option value='<%=rsEtc("emp_etc_name")%>'><%=rsEtc("emp_etc_name")%></option>
					<%
						rsEtc.MoveNext()
					Loop
					rsEtc.Close() : Set rsEtc = Nothing
					DBConn.Close : Set DBConn = Nothing
					%>
						</select>
					</td>
					<th>�޼�<span style="color:red;">*</span></th>
					<td class="left">
						<select name="lang_grade" id="lang_grade" style="width:100px">
							<option value="">����</option>
							<option value='�޼�����'>�޼�����</option>
							<option value='3��'>3��</option>
							<option value='2��'>2��</option>
							<option value='1��'>1��</option>
						</select>
					</td>
				  <th>����<span style="color:red;">*</span></th>
				  <td class="left">
					<input type="text" name="lang_point" id="lang_point" style="width:80px; ime-mode:active" onKeyUp="checklength(this,4);"/>
				  </td>
				</tr>
				<tr>
					<th>�����<span style="color:red;">*</span></th>
					<td colspan="5" class="left">
						<input type="text" name="lang_get_date" style="width:80px;text-align:center" id="datepicker"/>&nbsp;
					</td>
				</tr>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
			</div>
			<input type="hidden" name="curr_date" value="<%=curr_date%>"/>
			</form>
		</div>
	</body>
</html>