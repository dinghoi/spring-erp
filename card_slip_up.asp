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

card_gubun = abc("card_gubun")
slip_month = abc("slip_month")
file_type = abc("file_type")

If slip_month = "" Then
	slip_month = Mid(Now(), 1, 4) + Mid(Now(), 6, 2)
End If

from_date = Mid(slip_month, 1, 4) & "-" & Mid(slip_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

If card_gubun = "" Then
	ck_sw = "y"
Else
	ck_sw = "n"
End If

Set cn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")

Dim path, filename, fileType, file_name, save_path
Dim company, as_type, paper_no
Dim xgr, fldcount

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

title_line = "ī�� ���� ���ε�"

Dim att_file
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ȸ�� �ý���</title>
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
				if(document.frm.card_gubun.value == "") {
					alert ("ī�������� �����ϼ���");
					return false;
				}

				if(document.frm.slip_month.value == "") {
					alert ("����� �����ϼ���");
					return false;
				}

				if(document.frm.att_file.value == "") {
					alert ("���ε� ���� ������ �����ϼ���");
					return false;
				}

				return true;
			}

			function frm1check(){
				if(chkfrm1()){
					document.frm1.submit ();
				}
			}

			function chkfrm1(){
				if(confirm('DB�� ���ε� �Ͻðڽ��ϱ�?') == true) {
					return true;
				}

				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/card_slip_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="card_slip_up.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ε峻��</dt>
                        <dd>
                            <p>
								<label>
								<strong>ī������ : </strong>
                                    <select name="card_gubun" id="card_gubun" style="width:80px">
                                        <option value="">����</option>
                                        <option value="BCī��" <%If card_gubun = "BCī��" Then %>selected<%End If %>>BCī��</option>
                                        <option value="kb����ī��" <%If card_gubun = "kb����ī��" Then %>selected<%End If %>>kb����ī��</option>
                                        <option value="����ī��" <%If card_gubun = "����ī��" Then %>selected<%End If %>>����ī��</option>
                                        <option value="��Ƽī��" <%If card_gubun = "��Ƽī��" Then %>selected<%End If %>>��Ƽī��</option>
                                        <option value="�Ե�ī��" <%If card_gubun = "�Ե�ī��" Then %>selected<%End If %>>�Ե�ī��</option>
                                    </select>
								</label>
								<label>
								<strong>��ǥ��� : </strong>
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
							<col width="4%" >
							<col width="6%" >
							<col width="7%" >
							<col width="11%" >
							<col width="6%" >
							<col width="*" >
							<col width="10%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�Ǽ�</th>
								<th scope="col">���</th>
								<th scope="col">�����</th>
								<th scope="col">ī������</th>
								<th scope="col">ī���ȣ</th>
								<th scope="col">�����</th>
								<th scope="col">�ŷ�ó</th>
								<th scope="col">����</th>
								<th scope="col">��������</th>
								<th scope="col">����</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim rs_emp
						Dim card_num
						Dim tot_price, tot_cost, tot_cost_vat, tot_upjong
						Dim tot_account, date_err, i
						Dim approve_no, slip_date, card_no, customer, customer_no
						Dim upjong, price, cost_vat, cancel_yn
						Dim sql, rs_card, rs_upjong
						Dim reg_sw, date_sw, card_month
						Dim cost, upjong_sw, account_sw, account, account_item
						Dim owner_sw, card_type, emp_name, car_vat_sw
						Dim imsi_no

						tot_price = 0
						tot_cost = 0
						tot_cost_vat = 0
						tot_err = 0
						tot_upjong = 0
						tot_account = 0
						date_err = 0

						If rowcount > -1 Then
							For i = 0 To rowcount
								If xgr(1, i) = "" Or IsNull(xgr(1, i)) Then
									Exit For
								End If

								'BCī���� ���
								If card_gubun = "BCī��" Then
									If Trim(xgr(0, i)) = "�ű�" Then
										cancel_yn = "N"
									Else
										cancel_yn = "Y"
									End If

									slip_date = xgr(8, i)
									card_no = xgr(1, i)
									customer = xgr(22, i)
									customer_no = xgr(21, i)
									upjong = Replace(xgr(23, i), " ", "")
									price = xgr(11, i)
									cost_vat = xgr(15, i)
									approve_no = xgr(7, i)

									If price = "" Or IsNull(price) Then
										price = "'"&xgr(11, i)
										response.write(price)
									End If
 								End If

								'//2017-06-08 add. kb����ī��
								If card_gubun = "kb����ī��" Then
									If Trim(xgr(15, i)) = "����" Then
										cancel_yn = "N"
									Else
										cancel_yn = "Y"
									End If

									slip_date = xgr(0,i)

									If Trim(slip_date & "") <> "" Then
										slip_date = Replace(slip_date, ".", "-")
									End If

									card_num = xgr(4, i)
									card_no = xgr(4, i)
									card_no = Right(card_no, 7)
									'Response.write card_no

									customer = xgr(6,i)
									customer_no = xgr(18, i)
									upjong = xgr(7, i)
	'Response.write("<br>[["&VarType(xgr(10,i))&"]] ") ' 8 : String 1 : null ���Ű��� asp!! ������ ����� ���д´�. ������� ������ ���̰� ���� ������ �ؾ�..
									price = xgr(10, i)
									cost_vat = xgr(11, i)
									approve_no = xgr(14, i)

									If price = "" Or IsNull(price) Then
										price = "'"&xgr(11,i)
										response.write(price)
									End If
								End If

								'��Ƽī��
								If card_gubun = "��Ƽī��" Then
									If Trim(xgr(14, i)) = "����" Then
										cancel_yn = "N"
									Else
										cancel_yn = "Y"
									End If

									slip_date = xgr(4, i)
									imsi_no = xgr(1, i)
									card_no = Mid(imsi_no, 1, 4) & "-" & Mid(imsi_no, 5, 4) & "-" & Mid(imsi_no, 9, 4) & "-" & Right(imsi_no, 4)
									customer = xgr(8, i)
									customer_no = xgr(9, i)
									upjong = Replace(Trim(xgr(17, i)), " ", "")
									price = xgr(10, i)
									cost_vat = xgr(21, i)
									approve_no = xgr(20, i)
								End If

								' ����ī�� 	L(9410-6440-9)
								If card_gubun = "����ī��" Then

									'�ű� �ۼ�[����ȣ_20201215]	=======================
									slip_date = replace(xgr(6,i),".","-")

									'imsi_no = xgr(0,i)
									'card_no = mid(imsi_no,2,3) & "-" &right(imsi_no,4)
									card_no = xgr(0, i)	'�̿� ī��(ī�� ��ȣ)

									customer = xgr(19,i)
									imsi_no = xgr(18,i)
									customer_no = mid(imsi_no,1,3) & "-" & mid(imsi_no,4,2) & "-" &right(imsi_no,5)
									upjong = replace(xgr(20,i)," ","")
									price = xgr(9,i)
									cost_vat = xgr(12,i)
									approve_no = xgr(5,i)

									'�������ڴ� �̿��Ͻ÷� ����[����ȣ_20201217]
									'slip_date = Replace(Left(xgr(0, i), 10), ".", "-")
									'approve_no = xgr(2, i)	'���ι�ȣ
									'card_no = xgr(3, i)	'�̿� ī��(ī�� ��ȣ)
									'customer = xgr(5, i)	'������ ��
									'customer_no = ""	'������ ��ȣ
									'upjong = xgr(6, i)	'����
									'price = xgr(7, i)	'�̿�ݾ�

									'�ΰ��� ���
									'If xgr(10,i) = "����" Then
									'	cost_vat = price - Int(price/1.1)
									'Else
									'	cost_vat = 0
									'End If

									'��� ����
									If price < 0 Then
										cancel_yn = "Y"
									Else
										cancel_yn = "N"
									End If
								End If

			                    ' �Ե�ī�� 	LOCAL -> ù 4�ڸ� 9409, AMEX -> ù 4�ڸ� 3762 , VISA -> ù 4�ڸ� 4670
								If card_gubun = "�Ե�ī��" Then
									slip_date = Replace(xgr(5, i), ".", "-")
									imsi_no = xgr(2, i)
									imsi_card_no = Right(imsi_no, 3)

									sql = "select * from card_owner where card_type like '%�Ե�%' and right(card_no,3) = '"&imsi_card_no&"'"
									Set rs_card = DBConn.Execute(sql)

									If rs_card.EOF Or rs_card.BOF Then
										card_no = imsi_no
									Else
										card_no = rs_card("card_no")
									End If

	'								if xgr(1,i) = "LOCAL" then
	'									card_no = "9409" + mid(imsi_no,5)
	'								  elseif xgr(1,i) = "VISA" then
	'									card_no = "4670" + mid(imsi_no,5)
	'								  else
	'									card_no = "3762" + mid(imsi_no,5)
	'								end if

									customer = xgr(7, i)
									customer_no = xgr(25, i)
									upjong = replace(xgr(26, i), " ", "")
									price = xgr(8, i)

									If xgr(1, i) = "LOCAL" Then
										cost_vat = price - Int(price/1.1)
									Else
										cost_vat = 0
									End If

									approve_no = xgr(15, i)

									If Trim(xgr(12, i)) = "���Կ���" Then
										cancel_yn = "N"
									Else
										cancel_yn = "Y"
									End If
								End If

								If approve_no = "" Or IsNull(approve_no) Or approve_no = " " Then
									approve_no = CStr(Mid(slip_date, 1, 4)) + CStr(Mid(slip_date, 6, 2)) + CStr(Mid(slip_date, 9, 2))
								End If

								'If slip_date => from_date And slip_date <= to_date Then
								If slip_date >= from_date And slip_date <= to_date Then

									'ī�� ��� ���� ��ȸ
									'sql = "select * from card_slip where approve_no = '"&approve_no&"' and cancel_yn = '"&cancel_yn&"'"
									'Set rs_card = dbconn.execute(sql)

									'If rs_card.EOF Or rs_card.BOF Then
									'	reg_sw = "N"
									'Else
									'	reg_sw = "Y"
									'End If
									objBuilder.Append "SELECT COUNT(*) AS card_cnt "
									objBuilder.Append "FROM card_slip "
									objBuilder.Append "WHERE approve_no = '"&approve_no&"' AND cancel_yn = '"&cancel_yn&"' "

									Set rs_card = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_card("card_cnt") = "0" Then
										reg_sw = "N"
									Else
										reg_sw = "Y"
									End If

									rs_card.Close()

									date_sw = "Y"
									card_month = Mid(slip_date, 1, 4) & Mid(slip_date, 6, 2)

									If card_month <> slip_month Then
										date_err = date_err + 1
										date_sw = "N"
									End If

									cost = Int(price) - Int(cost_vat)
									tot_price = tot_price + Int(price)
									tot_cost = tot_cost + Int(cost)
									tot_cost_vat = tot_cost_vat + Int(cost_vat)

									upjong_sw = "Y"
									account_sw = "Y"

									objBuilder.Append "SELECT account, account_item, tax_yn "
									objBuilder.Append "FROM card_upjong "
									objBuilder.Append "WHERE card_upjong = '" & upjong &"' "

									Set rs_upjong = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_upjong.EOF Or rs_upjong.BOF Then
										upjong_sw = "N"
										tot_upjong = tot_upjong + 1
										account = "�����"
										account_item = "�����"
									Else
										account = rs_upjong("account")
										account_item = rs_upjong("account_item")

										If account = "" Or account_item = "" Or IsNull(account) Or IsNull(account_item) Then
											account_sw = "N"
											tot_account = tot_account + 1
										End If

										If rs_upjong("tax_yn") = "Y" Then
											If cost_vat = 0 Then
												cost_vat = CLng((price/1.1)/10)
												cost = price - cost_vat
											End If
										ElseIf rs_upjong("tax_yn") = "N" Then
											If cost_vat <> 0 Then
												cost_vat = 0
												cost = price
											End If
										End If
									End If

									rs_upjong.Close()

									owner_sw = "Y"

									objBuilder.Append "SELECT cdot.card_no, cdot.car_vat_sw, cdot.card_type, cdot.emp_no, "
									objBuilder.Append "	emtt.emp_pay_id "
									objBuilder.Append "FROM card_owner AS cdot "
									objBuilder.Append "INNER JOIN emp_master AS emtt ON cdot.emp_no = emtt.emp_no "

									If card_gubun = "����ī��" Then
										'ī�� ��ȣ ��ȸ_�� 8�ڸ� �� ��ȸ���� �� 4�ڸ�, �� 4�ڸ��� ����[����ȣ_20201217]
										objBuilder.Append "WHERE RIGHT(cdot.card_no, 4) = '" & Right(card_no, 4) &"' "
										objBuilder.Append "AND LEFT(cdot.card_no, 4) = '" & Left(card_no, 4) &"'"
									ElseIf card_gubun = "kb����ī��" Then
										' 20180727 ���� - ī���ȣ �Է��� �߸� �� ���� �ִ�..
										'objBuilder.Append "WHERE LEFT(cdot.card_no,7) = '" & Left(card_num,7) & "' "
										'objBuilder.Append "AND RIGHT(cdot.card_no,4) = '" & Right(card_num, 4) & "' "
										objBuilder.Append "WHERE cdot.card_no = '"&card_num&"' "
									Else
										objBuilder.Append "WHERE cdot.card_no = '" & card_no &"' "
									End If

									Set rs_card = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_card.EOF Or rs_card.BOF Then
										owner_sw = "N"
										tot_err = tot_err + 1
										card_type = "�̵��"
										emp_name = "������"
										emp_no = ""
										car_vat_sw = "C"
									Else
										card_no = rs_card("card_no")
										car_vat_sw = rs_card("car_vat_sw")
										card_type = rs_card("card_type")
										emp_no = rs_card("emp_no")

										'NKP ����ڸ� ��ȸ
										objBuilder.Append "SELECT user_name "
										objBuilder.Append "FROM memb "
										objBuilder.Append "WHERE user_id = '"&emp_no&"' "

										Set rs_emp = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										If rs_emp.EOF Or rs_emp.BOF Then
											emp_name = "�������"
										Else
											emp_name = rs_emp("user_name")
										End If

										'����� ���̵� üũ �߰�[����ȣ_20210716]
										If rs_card("emp_pay_id") = "2" Then
											owner_sw = "N"
											tot_err = tot_err + 1
											emp_name = "�������(���))"
										End If

										rs_emp.Close
									End If

									' ����ī�尡 �ٲ�� ���� �ؾ���
									If card_type = "�Ե�����ī��" Then
										account = "����������"
										account_item = "������"
										' 2014�� 12������ ����
										' car_vat_sw = "Y"
									End If
									' ����ī�� ���� ��

									If account = "����������" Then
										If car_vat_sw = "N" Then
											cost_vat = 0
											cost = price
										End If

										If car_vat_sw = "Y" Then
											If cost_vat = 0 Then
												cost_vat = CLng((price/1.1)/10)
												cost = price - cost_vat
											End If
										End If
									End If
								%>
								<tr>
									<td class="first"><%=i+1%></td>
									<td <%If reg_sw = "Y" Then%>bgcolor="#FFCCFF"<%End If%>>
									<!--���-->
									<%'���� ī�� ��ȣ ��� ����
									If reg_sw = "N" Then
										Response.Write "�̵��"
									Else
										Response.Write "���"
									End If
									%>
									</td>
									<!--�����-->
									<td <%If date_sw = "N" Then%>bgcolor="#FFCCFF" <%End If%>><%=slip_date%></td>
									<!--ī������-->
									<td><%=card_type%></td>
									<!--ī���ȣ-->
									<td><%=card_no%></td>
									<!--�����-->
									<td <%If owner_sw = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=emp_name%></td>
									<!--�ŷ�ó-->
									<td><%=customer%></td>
									<!--����-->
									<td <%If upjong_sw = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=upjong%></td>
									<!--��������-->
									<td <%If account_sw = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=account%></td>
									<!--����-->
									<td><%=account_item%>&nbsp;</td>
									<!--�հ�-->
									<td class="right"><%=FormatNumber(price, 0)%></td>
									<!--���ް���-->
									<td class="right"><%=FormatNumber(cost, 0)%></td>
									<!--�ΰ���-->
									<td class="right"><%=FormatNumber(cost_vat, 0)%></td>
								</tr>
						<%
								End If

								rs_card.Close()
							Next	'Loop End

							Set rs_emp = Nothing
							Set rs_upjong = Nothing
							Set rs_card = Nothing

							DBConn.Close()
							Set DBConn = Nothing

						End If
						%>
							<tr>
								<th class="first">��(���)</th>
								<th>&nbsp;</th>
								<th><%=FormatNumber(date_err, 0)%></th>
								<th><%=FormatNumber(tot_err, 0)%></th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th><%=FormatNumber(tot_upjong, 0)%></th>
								<th><%=FormatNumber(tot_account, 0)%></th>
								<th>��(�ݾ�)</th>
								<th class="right"><%=FormatNumber(tot_price, 0)%></th>
								<th class="right"><%=FormatNumber(tot_cost, 0)%></th>
								<th class="right"><%=FormatNumber(tot_cost_vat, 0)%></th>
							</tr>
						</tbody>
					</table>
				</div>
				</form>
			<% If tot_cnt <> 0 And tot_err = 0 Then %>
				<form action="card_slip_up_ok.asp" method="post" name="frm1">
					<br>
                    <div align="center">
                        <span class="btnType01"><input type="button" value="DB����" onclick="javascript:frm1check();"NAME="Button1"></span>
                    </div>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
                    <input name="card_gubun" type="hidden" id="card_gubun" value="<%=card_gubun%>">
                    <input name="slip_month" type="hidden" id="slip_month" value="<%=slip_month%>">
					<br>
				</form>
			<% End If %>
		</div>
	</div>
	</body>
</html>
