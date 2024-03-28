<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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
Dim uploadForm, filenm
Dim sales_month, file_type
Dim from_date, to_date, end_date, cost_year, ck_sw
Dim path, filename, fileType, file_name, save_path, objFile
Dim rowcount, title_line, att_file, tot_price, tot_cost, tot_cost_vat
Dim reg_cnt, cn, rs, xgr, fldcount, tot_cnt

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")

uploadForm.AbsolutePath = True
uploadForm.Overwrite = true
uploadForm.MaxUploadSize = 1024*1024*50

sales_month = uploadForm("sales_month")
file_type = uploadForm("file_type")

If sales_month = "" Then
	sales_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)
End If

from_date = Mid(sales_month, 1, 4) & "-" & Mid(sales_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
cost_year = Mid(sales_month, 1, 4)

If sales_month = "" Then
	ck_sw = "y"
Else
	ck_sw = "n"
End If

If ck_sw = "n" Then
	Set filenm = uploadForm("att_file")(1)

	path = Server.MapPath ("/large_file")
	filename = filenm.safeFileName
	fileType = Mid(filename, InStrRev(filename, ".") + 1)
	file_name = "����κ�����"

	save_path = path & "\" & file_name & "." & fileType

	If fileType = "xls" Or fileType = "xlk" Then
		file_type = "Y"
		filenm.save save_path

		objFile = save_path
'		objFile = Request.form("att_file")
'		objFile = SERVER.MapPath("att_file")
'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
'		response.write(objFile)

		Set cn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")

		cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
		rs.Open "select * from [6:10000]", cn, "0"

		rowcount = -1
		xgr = rs.getRows
		rowcount = UBound(xgr, 2)
		fldcount = rs.fields.count
		tot_cnt = rowcount + 1

		'�ʵ� ���� üũ
		If fldcount <> 35 Then
			fld_cnt_err = "Y"
		End If
	Else
		objFile = "none"
		rowcount = -1
		file_type = "N"
	End If
Else
	objFile = "none"
	rowcount = -1
End If

title_line = "���� ���ݰ�꼭 ���ε�"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ȸ�� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
        <script src="/java/jquery-1.9.1.js"></script>
        <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>

		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}

			$(document).ready(function(){
				var rowcnt = '<%=rowcount%>';
				var fldcnt = '<%=fldcount%>';

				//���ε� �׸� ���� Ȯ��
				//console.log(rowcnt);
				if(parseInt(rowcnt) > -1 && parseInt(fldcnt) !== 35){
					alert('���ε� �׸� ������ ��ġ���� �ʽ��ϴ�.(�ʼ� �׸� ����:35��)');
					location.href = '/finance/sales_bill_upload.asp';
					return;
				}
			});

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				/*if (document.frm.bill_id.value == "") {
					alert ("��꼭 ������ �����ϼ���");
					return false;
				}*/

				if(document.frm.sales_month.value == ""){
					alert ("����� �����ϼ���");
					return false;
				}

				if(document.frm.att_file.value == ""){
					alert ("���ε� ���� ������ �����ϼ���");
					return false;
				}

				return true;
			}

			//DB ���ε�
			function upload_ok(){
				var result = confirm('DB�� ���ε� �Ͻðڽ��ϱ�?');

				if(result == true){
					document.frm.action = "/finance/sales_bill_upload_ok.asp";
					document.frm.submit();
				}
				return false;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/account_cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/finance/sales_bill_upload.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ε峻��</dt>
                        <dd>
                            <p>
								<label>
								<strong>������ : </strong>
                                	<input name="sales_month" type="text" value="<%=sales_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
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
							<col width="3%" >
							<col width="4%" >
							<col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="*" >
							<col width="4%" >
							<col width="7%" >

							<!--<col width="7%" >
							<col width="6%" >-->
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�Ǽ�</th>
								<th scope="col">���</th>
								<th scope="col">������</th>
								<th scope="col">����ȸ��</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">��ȣ</th>
								<th scope="col">��ǥ�ڸ�</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�����</th>
								<th scope="col">���������</th>

								<!--<th scope="col">��ǥ��ȣ</th>
								<th scope="col">�Աݿ�����</th>-->
							</tr>
						</thead>
						<tbody>
						<%
						Dim trade_no_err_cnt, error_cnt, i, rec_cnt
						Dim sales_date, approve_no, sales_company, trade_no, trade_owner, trade_company
						Dim price, cost, cost_vat, sales_memo, emp_name
						Dim reg_sw, sales_com_err, emp_name_err, trade_name, cost_sum_err, sum_cost, saupbu_err
						Dim rs_etc, rs_trade, rs_emp, rsTradeName
						Dim company_err_cnt, emp_err_cnt, cost_err_cnt, saupbu_err_cnt, date_err_cnt
						Dim fld_cnt_err, sales_date_err, collect_err_cnt
						Dim rsApprove, approve_err_cnt, approve_no_err

						'Dim slip_no, collect_due_date, slip_no_err, collect_due_date_err

						'�� �հ� ���� �ʱ�ȭ
						tot_price = 0
						tot_cost = 0
						tot_cost_vat = 0

						'�������� ���� ���� �ʱ�ȭ
						sales_date_err = "N"
						date_err_cnt = 0

						'���(���ι�ȣ) �ߺ� �Ǽ�
						reg_cnt = 0

						company_err_cnt = 0
						emp_err_cnt = 0
						trade_no_err_cnt = 0
						cost_err_cnt = 0
						saupbu_err_cnt = 0

						fld_cnt_err = "N"

						approve_err_cnt = 0

						'�� ���� ����
						error_cnt = 0

						'���ε� ������ ����
						If rowcount > -1 Then
							'���� ���ε带 ���� ���� �� ����(0->5)[����ȣ_20220223]
							For i = 0 To rowcount
								'���� ��ȣ üũ(���� ���� ������ ���� ��� rowcount ���ԵǹǷ� ���ι�ȣ�� üũ��)
								If f_toString(xgr(1, i), "") = "" Then
									Exit For
								End If

								sales_date = xgr(0, i)	'�ۼ�����
								approve_no = xgr(1, i)	'���ι�ȣ
								sales_company = f_SalesCompany(xgr(6, i))	'��ȣ(������)
								trade_no = xgr(9, i)	'���޹޴��ڻ���ڵ�Ϲ�ȣ
								trade_company = xgr(11, i)	'��ȣ(�ŷ�ó)
								trade_owner = xgr(12, i)	'��ǥ�ڸ�
								price = f_toString(xgr(14, i), 0)	'�հ�ݾ�
								cost = f_toString(xgr(15, i), 0)	'���ް���
								cost_vat = f_toString(xgr(16, i), 0)	'����
								sales_memo = xgr(26, i)	'ǰ���
								emp_name = xgr(33, i)	'�����
								saupbu = xgr(34, i)	'�μ�

								'slip_no = xgr(35, i)	'��ǥ��ȣ
								'collect_due_date = xgr(36, i)	'�Աݿ�����

								'�� �հ�
								tot_price = tot_price + CDbl(price)	'�հ� Total
								tot_cost = tot_cost + CDbl(cost)	'���� Total
								tot_cost_vat = tot_cost_vat + CDbl(cost_vat)	'���� Total

								'�������� ���� üũ
								If (sales_date < from_date Or sales_date > to_date) Or f_toString(sales_date, "") = "" Then
									date_err_cnt = date_err_cnt + 1
									sales_date_err = "Y"
								Else
									sales_date_err = "N"
								End If

								'�˻���� ���ι�ȣ �ߺ� �Ǽ� üũ
								objBuilder.Append "SELECT approve_no FROM saupbu_sales "
								objBuilder.Append "WHERE approve_no='"&approve_no&"' "
								objBuilder.Append "	AND REPLACE(SUBSTRING(sales_date,1,7),'-','')='"&sales_month&"';"

								Set rs_etc = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rs_etc.EOF Or rs_etc.BOF Then
									reg_sw = "N"
								Else
									reg_cnt = reg_cnt + 1
									reg_sw = "Y"
								End If
								rs_etc.Close()

								'�˻���� ���� ���ι�ȣ �ߺ� �Ǽ� üũ
								objBuilder.Append "SELECT approve_no FROM saupbu_sales "
								objBuilder.Append "WHERE approve_no='"&approve_no&"' "
								objBuilder.Append "	AND REPLACE(SUBSTRING(sales_date,1,7),'-','')<'"&sales_month&"';"

								Set rsApprove = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsApprove.EOF Or rsApprove.BOF Then
									approve_no_err="N"
								Else
									approve_no_err="Y"
									approve_err_cnt=approve_err_cnt+1
								End If

								'����ȸ��(��ȣ��) ���� üũ
								sales_com_err = "N"

								objBuilder.Append "SELECT trade_id FROM trade WHERE trade_name = '"&sales_company&"' "

								Set rs_trade = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rs_trade.EOF Or rs_trade.BOF Then
									company_err_cnt = company_err_cnt + 1
									sales_com_err = "Y"
								Else
									If rs_trade("trade_id") <> "�迭��" Then
										company_err_cnt = company_err_cnt + 1
										sales_com_err = "Y"
									End If
								End If
								rs_trade.Close()

								'����� ���� üũ
								emp_name_err = "N"

								If saupbu = "��Ÿ�����" Or saupbu = "ȸ�簣�ŷ�" Then
									'100359 : ������ �̻�
									objBuilder.Append "SELECT emp_no FROM emp_master "
									objBuilder.Append "WHERE emp_no = '100359' AND emp_name = '"&emp_name&"' "

									Set rs_emp = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_emp.EOF Or rs_emp.BOF Then
										emp_name_err = "Y"
										emp_err_cnt = emp_err_cnt + 1
										emp_no = "error"
									Else
										emp_no = rs_emp("emp_no")
									End If

								Else
									objBuilder.Append "SELECT emp_no FROM emp_master AS emmt "
									objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON eomt.org_code = emmt.emp_org_code "
									objBuilder.Append "WHERE emmt.emp_name = '"&emp_name&"' AND eomt.org_bonbu = '"&saupbu&"' "

									Set rs_emp = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_emp.EOF Or rs_emp.BOF Then
										emp_name_err = "Y"
										emp_err_cnt = emp_err_cnt + 1
										emp_no = "error"
									Else
										emp_no = rs_emp("emp_no")
									End If
								End If
								rs_emp.Close()

								'���޹޴��ڻ���ڹ�ȣ ����
								If f_toString(trade_no, "") <> "" Then
									trade_no = Replace(trade_no, "-", "")
								Else
									trade_no = ""
								End If

								objBuilder.Append "SELECT trade_name FROM trade WHERE trade_no = '"&trade_no&"' "

								Set rsTradeName = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsTradeName.EOF Or rsTradeName.BOF Then
									'trade_name = sales_company
									trade_name = trade_company
								Else
									trade_name = rsTradeName("trade_name")
								End If
								rsTradeName.Close()

								'�հ�ݾ� ���� üũ(�հ�ݾ�=���ް���+����)
								cost_sum_err = "N"
								sum_cost = CDbl(cost) + CDbl(cost_vat)

								If sum_cost <> CDbl(price) Then
									cost_err_cnt = cost_err_cnt + 1
									cost_sum_err = "Y"
								End If

								'��������� ���� üũ
								saupbu_err = "N"

								If saupbu = "��Ÿ�����" Or saupbu = "ȸ�簣�ŷ�" Then
									saupbu_err = "N"
								Else
									objBuilder.Append "SELECT saupbu FROM sales_org "
									objBuilder.Append "WHERE saupbu = '"&saupbu&"' AND sales_year='"&cost_year&"' "
									objBuilder.Append "ORDER BY sort_seq "

									Set rs_etc = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_etc.EOF or rs_etc.BOF Then
										saupbu_err_cnt = saupbu_err_cnt + 1
										saupbu_err = "Y"
									End If
									rs_etc.Close()
								End If

								'��ǥ��ȣ, �Աݿ����� �׸� ���� ó��[����ȣ_20220413]
								'��ǥ��ȣ ���� üũ(���� üũ �ڵ� ����)
								'slip_no_err = "N"

								'�Աݿ����� ���� üũ
								'collect_due_date_err = "N"

								'If collect_due_date = "" Or IsNull(collect_due_date) Then
								'	collect_due_date = ""
								'Else
								'	collect_due_date = "20" & Replace(collect_due_date, " . ", " ")
								'End If

								'If collect_due_date <> "" Then
								'	If IsDate(collect_due_date) Then
								'		collect_due_date_err = "N"
								'	Else
								'		collect_err_cnt = collect_err_cnt + 1
								'		collect_due_date_err = "Y"
								'	End If
								'End If
								%>
								<tr <%If reg_sw = "Y" Then%>style="background-color:burlywood;"<%End If%>>
									<td class="first"><%=i+1%></td>
									<td <%If approve_no_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%If reg_sw = "Y" Then%>���<%Else%>�̵��<%End If%></td>
									<td <%If sales_date_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=sales_date%></td>
									<td <%If sales_com_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=sales_company%></td>
									<td><%=trade_no%></td>
									<td><%=trade_name%></td>
									<td><%=trade_owner%></td>
									<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(price, 0)%></td>
									<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(cost, 0)%></td>
									<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(cost_vat, 0)%></td>
									<td class="left"><%=sales_memo%></td>
									<td <%If emp_name_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %>><%=emp_name%></td>
									<td <%If saupbu_err = "Y" Then %>bgcolor="#FFCCFF"<%End If%>><%=saupbu%>&nbsp;</td>

									<!--<td <%'If slip_no_err = "Y" Then %>bgcolor="#FFCCFF"<%'End If %>><%'=slip_no%>&nbsp;</td>
									<td <%'If collect_due_date_err = "Y" Then %>bgcolor="#FFCCFF"<%'End If %>><%'=collect_due_date%>&nbsp;</td>-->
								</tr>
						<%
							Next
							Set rs_etc = Nothing
							Set rs_trade =Nothing
							Set rsTradeName = Nothing
							Set rs_emp = Nothing

							rs.Close() : Set rs = Nothing
							cn.Close() :  Set cn = Nothing

							'�� ���� ����
							error_cnt=date_err_cnt+company_err_cnt+emp_err_cnt+cost_err_cnt
							error_cnt=error_cnt+saupbu_err_cnt+collect_err_cnt+approve_err_cnt
						Else
							Response.Write "<tr><td colspan='13' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If

						DBConn.Close() : Set DBConn = Nothing

						'����Ʈ �� ����
						rec_cnt = i
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>��</strong></td>
								<td class="right"><%=FormatNumber(reg_cnt, 0)%> ��</td><!--���-->
								<td>&nbsp;</td>
								<td class="right">&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=FormatNumber(tot_price, 0)%></td>
								<td class="right"><%=FormatNumber(tot_cost, 0)%></td>
								<td class="right"><%=FormatNumber(tot_cost_vat, 0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>

								<!--<td>&nbsp;</td>
								<td>&nbsp;</td>-->
							</tr>
						<%
						'���� �Ǽ�
						If error_cnt > 0 Then
						%>
							<tr bgcolor="#FFCCFF">
								<td class="first"><strong>Error</strong></td>
								<td class="right"><%=FormatNumber(approve_err_cnt, 0)%> ��</td><!--���ι�ȣ �ߺ�(�˻��⵵ ����)-->
								<td class="right"><%=FormatNumber(date_err_cnt, 0)%> ��</td><!--������-->
								<td class="right"><%=FormatNumber(company_err_cnt, 0)%> ��</td><!--����ȸ��-->
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right" colspan="3"><%=FormatNumber(cost_err_cnt, 0)%> ��</td><!--�հ�-->
								<td>&nbsp;</td>
								<td class="right"><%=FormatNumber(emp_err_cnt, 0)%> ��</td><!--�����-->
								<td class="right"><%=FormatNumber(saupbu_err_cnt, 0)%> ��</td><!--���������-->

								<!--<td>&nbsp;</td>
								<td class="right"><%'=FormatNumber(collect_err_cnt, 0)%> ��</td><!--�Աݿ�����-->
							</tr>
						<%End If%>
						</tbody>
					</table>
				</div>
				<%
				'DB Upload ���� ����
				'If reg_cnt <> rec_cnt  And rowcount > -1 And error_cnt = 0 Then
				If rowcount > -1 And error_cnt = 0 Then
				%>
					<br>
                    <div align="center">
                        <span class="btnType01">
							<input type="button" value="DB ���ε�" onclick="javascript:upload_ok();" />
						</span>
                    </div>
				<%
				End If
				%>
					<br>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>" />
				</form>
			</div>
		</div>
	</body>
</html>