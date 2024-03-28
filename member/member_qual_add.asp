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
Dim qual_seq, emp_name, qual_type, qual_grade
Dim qual_pass_date, qual_org, qual_no, qual_passport, qual_pay_id
Dim curr_date, title_line, rsQual, rs_etc

title_line = "�ڰݻ��� ���"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
			$(function(){
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
				if(document.frm.qual_type.value == ""){
					alert('�ڰ������� �Է��ϼ���');
					frm.qual_type.focus();
					return false;
				}

				if(document.frm.qual_org == ""){
					alert('�߱ޱ���� �����ϼ���');
					frm.qual_org.focus();
					return false;
				}

				if(document.frm.qual_no.value == ""){
					alert('�ڰݵ�Ϲ�ȣ�� �Է��ϼ���');
					frm.qual_no.focus();
					return false;
				}

				if(document.frm.qual_pass_date.value == ""){
					alert('�հݳ���ϸ� �Է��ϼ���');
					frm.qual_pass_date.focus();
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
			<form action="/member/member_qual_proc.asp" method="post" name="frm">
			<div class="gView">
			  <table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="11%" >
					<col width="22%" >
					<col width="11%" >
					<col width="*" >
					<col width="11%" >
					<col width="22%" >
				</colgroup>
				<tbody>
				<tr>
					<th style="background:#FFFFE6">����</th>
					<td colspan="5" class="left" bgcolor="#FFFFE6">
						<input type="text" name="m_name" id="m_name" size="14" value="<%=m_name%>" class="no-input" readonly/>
					</td>
				</tr>
				<tr>
					<th>�ڰ�����<span style="color:red;">*</span></th>
					<td class="left">
					<%
					objBuilder.Append "SELECT emp_etc_name FROM emp_etc_code "
					objBuilder.Append "WHERE emp_etc_type = '30' ORDER BY emp_etc_name ASC;"

					Set rs_etc = DBConn.Execute(objBuilder.ToString())
					objBuilder.Clear()
					%>
						<select name="qual_type" id="qual_type" style="width:140px">
							<option value="">����</option>
							<%
							Do Until rs_etc.EOF
							%>
								<option value='<%=rs_etc("emp_etc_name")%>'><%=rs_etc("emp_etc_name")%></option>
							<%
								rs_etc.MoveNext()
							Loop
							rs_etc.Close() : Set rs_etc = Nothing
							DBConn.Close() : Set DBConn = Nothing
							%>
						</select>
					</td>
					<th>���</th>
					<td colspan="3" class="left">
						<select name="qual_grade" id="qual_grade" style="width:90px">
							<option value="">����</option>
							<option value='1��'>1��</option>
							<option value='2��'>2��</option>
							<option value='3��'>3��</option>
							<option value='�ʱ�'>�ʱ�</option>
							<option value='�߱�'>�߱�</option>
							<option value='���'>���</option>
							<option value='Ư��'>Ư��</option>
						</select>
					</td>
				</tr>
				<tr>
					<th>�߱ޱ��<span style="color:red;">*</span></th>
					<td class="left">
						<input type="text" name="qual_org" id="qual_org" style="width:140px; ime-mode:active" onKeyUp="checklength(this,30);"/>
					</td>
					<th>�ڰ�<br>��Ϲ�ȣ<span style="color:red;">*</span></th>
					<td colspan="3" class="left">
						<input type="text" name="qual_no" id="qual_no" style="width:150px; ime-mode:active" onKeyUp="checklength(this,30);"/>&nbsp;
					</td>
				</tr>
				<tr>
					<th>�հݳ����<span style="color:red;">*</span></th>
					<td colspan="5" class="left">
						<input type="text" name="qual_pass_date" style="width:80px;text-align:center" id="datepicker"/>&nbsp;
					</td>
				</tr>
				<tr>
					<th>��¼�øNo</th>
					<td colspan="5" class="left">
						<input type="text" name="qual_passport" id="qual_passport" style="width:140px; ime-mode:active" onKeyUp="checklength(this,20);"/>
					</td>
				</tr>
				</tbody>
			  </table>
			</div>
			<br>
			<div align="center">
				<span class="btnType01"><input type="button" value="���" onclick="javascript:frmcheck();"/></span>
				<span class="btnType01"><input type="button" value="���" onclick="javascript:goAction();"/></span>
			</div>
			</form>
		</div>
	</body>
</html>