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
Dim abc,filenm
Dim tot_cnt, tot_err, tot_dept, tot_cust, tot_ddd
Dim tot_tel, tot_sido, tot_gugun, tot_dong, tot_addr
Dim tot_ce
Dim card_gubun, slip_month
Dim from_date, end_date, to_date, file_type
Dim ck_sw

Dim cn, rs

Dim objFile, rowcount
Dim title_line

Set abc = Server.CreateObject("ABCUpload4.XForm")
abc.AbsolutePath = True
abc.Overwrite = True
abc.MaxUploadSize = 1024*1024*50

tot_cnt = 0
tot_err = 0
tot_dept = 0
tot_cust = 0
tot_ddd = 0
tot_tel = 0
tot_sido = 0
tot_gugun = 0
tot_dong = 0
tot_addr = 0
tot_ce = 0

slip_month = abc("slip_month")
att_file = abc("att_file")

If slip_month = "" Then
	slip_month = Mid(Now(), 1, 4) + Mid(Now(), 6, 2)
End If

from_date = Mid(slip_month, 1, 4) & "-" & Mid(slip_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Dim path, filename, fileType, file_name, save_path
Dim company, as_type, paper_no
Dim xgr, fldcount, att_file

If att_file = "" Then
	ck_sw = "y"
Else
	ck_sw = "n"
End If

If ck_sw = "n" Then
	Set filenm = abc("att_file")(1)

	path = Server.MapPath ("/large_file")
	filename = filenm.safeFileName
	fileType = Mid(filename, InStrRev(filename, ".") + 1)
	file_name = company & "_" & as_type & "_" & paper_no

	save_path = path & "\" & file_name&"."&fileType

	If fileType = "xls" Or fileType = "xlk" Then
		file_type = "Y"
		filenm.save save_path

		objFile = save_path
'		objFile = Request.form("att_file")
'		objFile = SERVER.MapPath("att_file")
'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
'		response.write(objFile)

		cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
		rs.Open "select * from [1:10000]", cn, "0"

		rowcount = -1
		xgr = rs.getrows
		rowcount = UBound(xgr, 2)
		fldcount = rs.fields.count
		tot_cnt = rowcount + 1
	Else
		objFile = "none"
		rowcount = -1
		file_type = "N"
	End If
Else
	objFile = "none"
	rowcount = -1
End If

title_line = "�ŷ�ó�� ��Ȳ ���ε�"
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
				if(document.frm.slip_month.value == ""){
					alert ("��ϳ���� �Է��ϼ���");
					return false;
				}


				if(document.frm.att_file.value == ""){
					alert ("���ε� ���� ������ �����ϼ���");
					return false;
				}

				return true;
			}

			function frm1check(){
				if(chkfrm1()){
					document.frm1.submit();
				}
			}

			function chkfrm1(){
				if(confirm('DB�� ���ε� �Ͻðڽ��ϱ�?') == true){
					return true;
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/service/as_acpt_statics_up.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ε峻��</dt>
                        <dd>
                            <p>
								<label>
									<strong>��ϳ�� : </strong>
                                	<input name="slip_month" type="text" value="<%=slip_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
									<strong>���ε����� : </strong>
									<input name="att_file" type="file" id="att_file" size="60" value="<%=att_file%>" style="text-align:left">
								</label>

            					<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">
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
							<!--<col width="10%" >
							<col width="10%" >-->
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">No.</th>
								<th scope="col">����</th>
								<th scope="col">��ġ/���� �Ǽ�</th>
								<th scope="col">��ġ/����(�ð�)</th>
								<th scope="col">��� �Ǽ�</th>
								<th scope="col">���� �Ǽ�</th>
								<th scope="col">ȸ�� �Ǽ�</th>
								<!--<th scope="col">�������� �Ǽ�</th>
								<th scope="col">�������� �Ǽ�</th>-->
								<th scope="col">�հ�</th>
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

						If rowcount > -1 Then
							For i = 0 To rowcount
								If f_toString(xgr(0, i), "") = "" Or xgr(0, i) = "NaN" Then
									Exit For
								End If

								as_company = f_toString(xgr(0, i), "")	'�ŷ�ó
								as_set = f_toString(xgr(1, i), 0)	'��ġ/����
								set_time = f_toString(xgr(2, i), 0)	'��ġ/���� �ð�
								as_error = f_toString(xgr(3, i), 0)	'���
								as_testing = f_toString(xgr(4, i), 0)	'����
								as_collect = f_toString(xgr(5, i), 0)	'ȸ��

								'as_give_cowork = f_toString(xgr(6, i), 0)	'��������
								'as_get_cowork = f_toString(xgr(7, i), 0)	'��������

								'as_total = f_toString(xgr(8, i), 0)	'�ѰǼ�
								'total_time = f_toString(xgr(9, i), 0)	'�ѽð�

								as_total = f_toString(xgr(6, i), 0)'�հ�(�ŷ�ó �� �Ǽ� �հ�)

								tot_setting = tot_setting + as_set
								tot_set_time = tot_set_time + set_time
								tot_error = tot_error + as_error
								tot_testing = tot_testing + as_testing
								tot_collect = tot_collect + as_collect
								tot_total = tot_total + as_total

								'tot_give_cowork = tot_give_cowork + as_give_cowork
								'tot_get_cowork = tot_get_cowork + as_get_cowork
								%>
								<tr>
									<td class="first"><%=i+1%></td>
									<td><%=as_company%></td>
									<td><%=FormatNumber(as_set, 0)%></td>
									<td><%=FormatNumber(set_time, 0)%></td>
									<td><%=FormatNumber(as_error, 0)%></td>
									<td><%=FormatNumber(as_testing, 0)%></td>
									<td><%=FormatNumber(as_collect, 0)%></td>
									<!--<td><%'=FormatNumber(as_give_cowork, 0)%></td>
									<td><%'=FormatNumber(as_get_cowork, 0)%></td>-->
									<td><%=FormatNumber(as_total, 0)%></td>
								</tr>
						<%
							Next

							DBConn.Close() : Set DBConn = Nothing
						Else
							Response.Write "<tr><td colspan='8' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If
						%>
							<tr>
								<th class="first" colspan="2">��</th>
								<th><%=FormatNumber(tot_setting, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_set_time, 0)%>&nbsp;�ð�</th>
								<th><%=FormatNumber(tot_error, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_testing, 0)%>&nbsp;��</th>
								<th><%=FormatNumber(tot_collect, 0)%>&nbsp;��</th>
								<!--<th><%'=FormatNumber(tot_give_cowork, 0)%>&nbsp;��</th>
								<th><%'=FormatNumber(tot_get_cowork, 0)%>&nbsp;��</th>-->
								<th><%=FormatNumber(tot_total, 0)%>&nbsp;��</th>
							</tr>
						</tbody>
					</table>
				</div>
				</form>
			<% If tot_cnt <> 0 And tot_err = 0 Then %>
				<form action="/service/as_acpt_statics_proc.asp" method="post" name="frm1">
					<br>
                    <div align="center">
                        <span class="btnType01"><input type="button" value="DB ���ε�" onclick="javascript:frm1check();"></span>
                    </div>
					<input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
                    <input name="slip_month" type="hidden" id="slip_month" value="<%=slip_month%>">
					<br>
				</form>
			<% End If %>
		</div>
	</div>
	</body>
</html>
