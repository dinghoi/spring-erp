<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim title_line, slip_month, from_date, end_date, to_date
Dim rsAs, arrAs

slip_month = f_Request("slip_month")

If slip_month = "" Then
	slip_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)
End If

from_date = Mid(slip_month, 1, 4) & "-" & Mid(slip_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

objBuilder.Append "SELECT as_company, as_set, set_time, as_error, as_testing, as_collect, "
objBuilder.Append "	as_give_cowork, as_get_cowork, as_total, total_time "
objBuilder.Append "FROM as_acpt_status "
objBuilder.Append "WHERE as_month = '"&slip_month&"' "
objBuilder.Append "ORDER BY as_seq ASC "

Set rsAs = DBConn.Execute(objBuilder.ToString())

If Not rsAs.EOF Then
	arrAs = rsAs.getRows()
End If
rsAs.Close() : Set rsAs = Nothing

title_line = "���� A/S ��Ȳ"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
			/*
			$(function(){
				$("#datepicker").datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%'=request_date%>" );
			});

			$(function(){
				$( "#datepicker1" ).datepicker();
				$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker1" ).datepicker("setDate", "<%'=end_date%>" );
			});
			*/
			function frmcheck(){
				if(chkfrm()){
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if(document.frm.slip_month.value == "") {
					alert ("��ϳ���� �Է��ϼ���");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/service/as_acpt_statics_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
									<strong>��ϳ�� : </strong>
                                	<input name="slip_month" type="text" value="<%=slip_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
            					<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="*" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">��ȣ</th>
								<th scope="col">����</th>
								<th scope="col">��ġ/���� �Ǽ�</th>
								<th scope="col">��ġ/����(�ð�)</th>
								<th scope="col">��� �Ǽ�</th>
								<th scope="col">���� �Ǽ�</th>
								<th scope="col">ȸ�� �Ǽ�</th>
								<th scope="col">�������� �Ǽ�</th>
								<th scope="col">�������� �Ǽ�</th>
								<th scope="col">�� �Ǽ�</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim as_company, as_set, set_time, as_error, as_collect, as_testing
						Dim tot_setting, tot_set_time, tot_error, tot_collect, tot_testing
						Dim i, as_total, total_time, tot_total

						Dim as_give_cowork, as_get_cowork, tot_give_cowork, tot_get_cowork

						tot_setting = 0
						tot_set_time = 0
						tot_error = 0
						tot_testing = 0
						tot_collect = 0

						tot_give_cowork = 0
						tot_get_cowork = 0

						If IsArray(arrAs) Then
							For i = LBound(arrAs) To UBound(arrAs, 2)
								as_company = arrAs(0, i)	'�ŷ�ó
								as_set = f_toString(arrAs(1, i), 0)	'��ġ/����
								set_time = f_toString(arrAs(2, i), 0)	'��ġ/���� �ð�
								as_error = f_toString(arrAs(3, i), 0)	'���
								as_testing = f_toString(arrAs(4, i), 0)	'����
								as_collect = f_toString(arrAs(5, i), 0)	'ȸ��

								as_give_cowork = f_toString(arrAs(6, i), 0)	'�ѰǼ�
								as_get_cowork = f_toString(arrAs(7, i), 0)	'�ѰǼ�


								as_total = f_toString(arrAs(8, i), 0)	'�ѰǼ�
								total_time = f_toString(arrAs(9, i), 0)	'�ѽð�

								tot_setting = tot_setting + as_set
								tot_set_time = tot_set_time + set_time
								tot_error = tot_error + as_error
								tot_testing = tot_testing + as_testing

								tot_give_cowork = tot_give_cowork + as_give_cowork
								tot_get_cowork = tot_get_cowork + as_get_cowork

								tot_collect = tot_collect + as_collect
								tot_total = tot_total + as_total
								%>
								<tr>
									<td class="first"><%=i+1%></td>
									<td><%=as_company%></td>
									<td><%=FormatNumber(as_set, 0)%></td>
									<td><%=FormatNumber(set_time, 0)%></td>
									<td><%=FormatNumber(as_error, 0)%></td>
									<td><%=FormatNumber(as_testing, 0)%></td>
									<td><%=FormatNumber(as_collect, 0)%></td>
									<td><%=FormatNumber(as_give_cowork, 0)%></td>
									<td><%=FormatNumber(as_get_cowork, 0)%></td>
									<td><%=FormatNumber(as_total, 0)%></td>
								</tr>
						<%
							Next	'Loop End

							DBConn.Close() : Set DBConn = Nothing
						End If
						%>
							<tr>
								<th class="first" colspan="2">��</th>
								<th><%=FormatNumber(tot_setting, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_set_time, 0)%>&nbsp;�ð�</th>
								<th><%=FormatNumber(tot_error, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_testing, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_collect, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_give_cowork, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_get_cowork, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_total, 0)%>&nbsp;��</th>
							</tr>
						</tbody>
					</table>
				</div>
				</form>
		</div>
	</div>
	</body>
</html>
