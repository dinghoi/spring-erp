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
Dim uploadForm, bill_month, file_type, from_date, end_date, to_date
Dim ck_sw, filenm, cn, rs, title_line, objFile, rowcount, att_file
Dim path, filename, fileType, file_name, save_path, xgr, tot_cnt
Dim fld_cnt_err, fldcount

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")
uploadForm.AbsolutePath = True
uploadForm.Overwrite = true
uploadForm.MaxUploadSize = 1024*1024*50

bill_month = uploadForm("bill_month")
file_type = uploadForm("file_type")

If bill_month = "" Then
	bill_month = Mid(Now(),1,4)&Mid(Now(),6,2)
End If

from_date = Mid(bill_month,1,4)&"-"&Mid(bill_month,5,2)&"-01"
end_date = DateValue(from_date)
end_date = DateAdd("m",1,from_date)
to_date = CStr(DateAdd("d",-1,end_date))

If bill_month = "" Then
	ck_sw = "y"
Else
	ck_sw = "n"
End If

If ck_sw = "n" Then
	Set filenm = uploadForm("att_file")(1)

	path = Server.MapPath("/large_file")
	filename = filenm.safeFileName
	fileType = Mid(filename,InStrRev(filename,".")+1)
	file_name = "e�����ϰ�ó��"

'		save_path = path & "\" & filename
	save_path = path & "\" & file_name&"."&fileType

	If fileType = "xls" or fileType = "xlk" Then
		file_type = "Y"
		filenm.save save_path
		objFile = save_path

'		objFile = Request.form("att_file")
'		objFile = SERVER.MapPath("att_file")
'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
'		response.write(objFile)

		Set cn = Server.CreateObject("ADODB.Connection")
		Set rs = Server.CreateObject("ADODB.Recordset")

		cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ="&objFile&";"
		rs.Open "select * from [3:10000]",cn,"0"

		rowcount = -1
		xgr = rs.getrows
		rowcount = UBound(xgr,2)
		fldcount = rs.fields.count
		tot_cnt = rowcount + 1

		'�ʵ� ���� üũ
		If fldcount <> 14 Then
			fld_cnt_err = "Y"
		End If
	Else
		objFile = "none"
		rowcount=-1
		file_type = "N"
	End If
Else
	objFile = "none"
	rowcount = -1
End If

title_line = "E���� ����ϰ����ε�"

' 2019.02.15 �ڼ��� ��û 19�� ���� X,Y �÷��� ���� '��Ź����ڵ�Ϲ�ȣ','��ȣ' �� �߰��Ǿ���
' ��Ģ������ ���α׷��� �����ؾ� �ϳ� �ڼ��κ����� �� �� �÷��� �����ϰ� ���ε��ϰڴٰ� ��..
' ������ ����� �ٸ���(�ٸ�����)���� 	�ߵȴٰ� ��.. (������ �� ���� ��°��� �ǽ�..)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
		<script src="/java/jquery-1.9.1.js"></script>
		<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<!--<script type="text/javascript" src="/java/js_window.js"></script>-->
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}

			$(document).ready(function(){
				var rowcnt = '<%=rowcount%>';
				var fldcnt = '<%=fldcount%>';

				//���ε� �׸� ���� Ȯ��
				console.log(rowcnt);
				console.log(fldcnt);
				if(parseInt(rowcnt) > -1 && parseInt(fldcnt) !== 14){
					alert('���ε� �׸� ������ ��ġ���� �ʽ��ϴ�.(�ʼ� �׸� ����:14��)');
					location.href = '/cost/tax_esero_upload.asp';
					return;
				}
			});

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.bill_month.value == ""){
					alert ("����� �����ϼ���");
					return false;
				}

				if(document.frm.att_file.value == ""){
					alert ("���ε� ���� ������ �����ϼ���");
					return false;
				}
				return true;
			}

			function upload_ok(){
				var result = confirm('DB�� ���ε� �Ͻðڽ��ϱ�?');

				if(result == true){
					document.frm.action = "/cost/tax_esero_upload_proc.asp";
					document.frm.submit();
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3><br/>
				<form action="/cost/tax_esero_upload.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ε峻��</dt>
						<dd>
							<p>
							<label>
								<strong>��꼭 ������ : </strong>
								<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);"/>
							</label>
							<label>
								<strong>���ε����� : </strong>
								<input name="att_file" type="file" id="att_file" size="60" value="<%=att_file%>" style="text-align:left"/>
							</label>
							<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>"/>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"/></a>
							</p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%">
							<col width="6%">
							<col width="10%">
							<col width="11%">
							<col width="6%">
							<col width="7%">
							<col width="7%">
							<col width="6%">
							<col width="12%">
							<col width="7%">
							<col width="7%">
							<col width="7%">
							<!--<col width="7%">-->
							<col width="*">
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�Ǽ�</th>
								<th scope="col">��������</th>
								<th scope="col">��꼭����ȸ��</th>
								<th scope="col">��ȣ��</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">�ŷ�����</th>
								<th scope="col">�����</th>
								<th scope="col">��������ڵ�</th>
								<th scope="col">����</th>
								<th scope="col">�������</th>
								<!--<th scope="col">�������</th>-->
								<th scope="col">��������</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim error_cnt, i, reg_cnt, owner_cnt, trade_no_err_cnt, tot_price, tot_cost, tot_cost_vat
						Dim t_bill_date, t_approve, t_owner_company, t_trade_name, t_emp_name, t_emp_no, t_cost
						Dim t_cost_vat, t_tax_bill_memo, t_org_code, t_company, t_mg_saupbu
						Dim t_account, arr_str, arr_account, t_account_item, bill_date_err, date_err_cnt
						Dim rs_trade, owner_err, owner_trade_no, owner_company, t_price, slip_sw, t_account_str
						Dim org_code_err, org_code_cnt, saupbu_err, saupbu_cnt, cost_sum_err, slip_err, slip_cnt
						Dim sum_cost, price, cost_err_cnt, tot_err, slip_gubun, j, slip_account
						Dim emp_name_cnt, emp_name_err, rsEmp, k

						date_err_cnt = 0
						org_code_cnt = 0'��������ڵ� ���� ����
						owner_cnt = 0'�ŷ�ó ���� ����
						saupbu_cnt = 0'����� ���� ����
						slip_cnt = 0
						cost_err_cnt = 0

						emp_name_cnt = 0

						error_cnt = 0'�� ���� ����

						tot_price = 0
						tot_cost = 0
						tot_cost_vat = 0

						'���ε� ����Ÿ ����
						If rowcount > -1 Then
							For i=0 To rowcount
								'����� �� üũ(��������ڵ�, ����,�������, �������, �������� �� üũ)
								'If f_toString(xgr(1,i), "") = "" Then
								'	Exit For
								'End If

								t_bill_date = f_toString(xgr(0,i), "")'��������
								't_approve = f_toString(xgr(1,i), "")'���ι�ȣ
								t_owner_company = f_toString(xgr(2,i), "")'��꼭����ȸ��
								t_trade_name = f_toString(xgr(3,i), "")'��ȣ��
								t_price = f_toString(xgr(4,i), 0)'�հ�
								t_cost = f_toString(xgr(5,i), 0)'���ް���
								t_cost_vat = f_toString(xgr(6,i), 0)'�ΰ���
								t_tax_bill_memo = f_toString(xgr(7,i), "")'�ŷ�����

								't_emp_no = f_toString(xgr(8,i), "")'������
								t_emp_name = f_toString(xgr(9,i), "")'�����

								t_org_code = f_toString(xgr(10,i), "")'��������ڵ�
								t_company = f_toString(xgr(11,i), "")'����
								t_mg_saupbu = f_toString(xgr(12,i), "")'�������
								't_slip_gubun = f_toString(xgr(13,i), "")'�������
								t_account_str = f_toString(xgr(13,i), "")'��������

								If t_bill_date <> "" Then
									t_bill_date = CStr(t_bill_date)
								End If

								tot_err = "N"'��ü ����
								bill_date_err = "N"

								'�������� ���� üũ
								If (t_bill_date < from_date Or t_bill_date > to_date) Or f_toString(t_bill_date, "") = "" Then
									date_err_cnt = date_err_cnt + 1
									bill_date_err = "Y"

									tot_err = "Y"
								End If

								org_code_err = "N"'��������ڵ� üũ �ڵ�

								If t_org_code = "" And (t_company <> "" Or t_mg_saupbu <> "" Or t_account_str <> "") Then
									org_code_err = "Y"
									org_code_cnt = org_code_cnt + 1

									tot_err = "Y"
								End If

								'��� ��� ����� ��ȸ
								emp_name_err = "N"

								If t_company <> "" And t_mg_saupbu <> "" And t_account_str <> "" And t_org_code <> "" Then
									If t_emp_name = "" Then
										emp_name_err = "Y"
										emp_name_cnt = emp_name_cnt + 1

										tot_err = "Y"
									Else
										objBuilder.Append "SELECT emp_no FROM emp_master "
										objBuilder.Append "WHERE (emp_end_date IS NULL OR emp_end_date <> '' OR emp_end_date = '1900-01-01') "
										objBuilder.Append "	AND emp_org_code = '"&t_org_code&"' AND emp_name = '"&t_emp_name&"';"

										Set rsEmp = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										If rsEmp.EOF Or rsEmp.BOF Then
											emp_name_err = "Y"
											emp_name_cnt = emp_name_cnt + 1

											tot_err = "Y"
										End If
										rsEmp.Close()
									End If
								End If

								owner_err = "N"'���� üũ �ڵ�

								If t_company = "" And (t_org_code <> "" Or t_mg_saupbu <> "" Or t_account_str <> "") Then
									owner_err = "Y"
									owner_cnt = owner_cnt + 1

									tot_err = "Y"
								End If

								objBuilder.Append "SELECT trade_name FROM trade "
								objBuilder.Append "WHERE trade_name='"&t_company&"';"

								Set rs_trade = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rs_trade.EOF Or rs_trade.BOF Then
									owner_err = "Y"
									owner_cnt = owner_cnt + 1
									owner_company = owner_trade_no&"_Error"

									tot_err = "Y"
								Else
									owner_company = rs_trade("trade_name")
								End If
								rs_trade.Close()

								saupbu_err = "N"'������� üũ �ڵ�

								If t_mg_saupbu = "" And (t_org_code <> "" Or t_company <> "" Or t_account_str <> "") Then
									saupbu_err = "Y"
									saupbu_cnt = saupbu_cnt + 1

									tot_err = "Y"
								End If

								'������� üũ
								slip_err = "N"'


								'������� ����
								If t_account_str = "" And (t_org_code <> "" Or t_company <> "" Or t_mg_saupbu <> "") Then
									t_account = ""
									t_account_item = ""

									slip_err = "Y"
									slip_cnt = slip_cnt + 1

									tot_err = "Y"
								Else
									arr_str = Split(t_account_str, ")")'��������

									For j = 0 To UBound(arr_str)
										If j = 0 Then
											slip_gubun = Replace(arr_str(j), "(", "")
										Else
											slip_account = arr_str(j)
										End If
									Next

									If slip_gubun = "���" Then
										arr_account = Split(slip_account, "-")

										For k = 0 To UBound(arr_account)
											If k = 0 Then
												t_account = arr_account(k)
											Else
												t_account_item = arr_account(k)
											End If
										Next
									Else
										t_account = slip_account
										t_account_item = slip_account
									End If
								End If

								'�հ�ݾ� ���� üũ(�հ�ݾ�=���ް���+����)
								cost_sum_err = "N"
								sum_cost = CDbl(t_cost) + CDbl(t_cost_vat)

								If sum_cost <> CDbl(t_price) Then
									cost_err_cnt = cost_err_cnt + 1
									cost_sum_err = "Y"

									tot_err = "Y"
								End If
						%>
							<tr <%If tot_err = "Y" Then%>style="background-color:burlywood;"<%End If%>>
								<td class="first"><%=i+1%></td>
								<td <%If bill_date_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=t_bill_date%></td>
								<td><%=t_owner_company%></td>
								<td><%=t_trade_name%></td>
								<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(t_price,0)%></td>
								<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(t_cost,0)%></td>
								<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(t_cost_vat,0)%></td>
								<td><%=t_tax_bill_memo%></td>
								<td <%If emp_name_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=t_emp_name%></td>
								<td <%If org_code_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=t_org_code%></td>
								<td <%If owner_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=t_company%></td>
								<td	<%If saupbu_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=t_mg_saupbu%></td>
								<!--<td	<%'If slip_err = "Y" Then%>bgcolor="#FFCCFF"<%'End If%>><%'=t_slip_gubun%></td>-->
								<td	<%If slip_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=t_account_str%></td>
						<%
								tot_price = tot_price + t_price
								tot_cost = tot_cost + t_cost
								tot_cost_vat = tot_cost_vat + t_cost_vat
							Next
							Set rs_trade = Nothing
							Set rsEmp = Nothing

							rs.Close() : Set rs = Nothing
							cn.Close() :  Set cn = Nothing

							'�� ���� ����
							error_cnt = date_err_cnt + org_code_cnt + owner_cnt + saupbu_cnt + slip_cnt + cost_err_cnt + emp_name_cnt

							DBConn.Close() : Set DBConn = Nothing
						Else
							Response.Write "<tr><td colspan='13' style='height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If

						'����Ʈ �� ����
						'rec_cnt = i
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>��</strong></td>
								<td class="right"><%=FormatNumber(i,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=FormatNumber(tot_price,0)%></td>
								<td class="right"><%=FormatNumber(tot_cost,0)%></td>
								<td class="right"><%=FormatNumber(tot_cost_vat,0)%></td>
								<!--<td>&nbsp;</td>-->
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
						'���� �Ǽ�
						If error_cnt > 0 Then
						%>
							<tr bgcolor="#FFCCFF">
								<td class="first"><strong>Error</strong></td>
								<td class="right"><%=FormatNumber(date_err_cnt, 0)%> ��</td><!--��������-->
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right" colspan="3"><%=FormatNumber(cost_err_cnt, 0)%> ��</td><!--�հ�-->
								<!--<td>&nbsp;</td>-->
								<td class="right"><%=FormatNumber(emp_name_cnt, 0)%> ��</td>
								<td class="right"><%=FormatNumber(org_code_cnt, 0)%> ��</td>
								<td class="right"><%=FormatNumber(owner_cnt, 0)%> ��</td>
								<td class="right"><%=FormatNumber(saupbu_cnt, 0)%> ��</td>
								<td class="right"><%=FormatNumber(slip_cnt, 0)%> ��</td>
							</tr>
						<%End If%>
						</tbody>
					</table>
				</div>
				<%
				'DB Upload ���� ����
				'If reg_cnt <> rec_cnt And owner_cnt = 0 And trade_no_err_cnt = 0 And rowcount > -1 Then
				If rowcount > -1 And error_cnt = 0 Then
				%>
					<br>
                    <div align="center">
                        <span class="btnType01"><input type="button" value="DB�� ���ε�" onclick="javascript:upload_ok();"/></span>
                    </div>
				<%End If %>
					<br>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>"/>
				</form>
		</div>
	</div>
	</body>
</html>