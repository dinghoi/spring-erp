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
Dim uploadForm, bill_id, bill_month, file_type, from_date, end_date, to_date
Dim ck_sw, filenm, cn, rs, title_line, objFile, rowcount, att_file
Dim path, filename, fileType, file_name, save_path, xgr, tot_cnt
Dim fld_cnt_err, fldcount

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")
uploadForm.AbsolutePath = True
uploadForm.Overwrite = true
uploadForm.MaxUploadSize = 1024*1024*50

bill_id = uploadForm("bill_id")
bill_month = uploadForm("bill_month")
file_type = uploadForm("file_type")

If bill_month = "" Then
	bill_month = Mid(Now(),1,4)&Mid(Now(),6,2)
End If

from_date = Mid(bill_month,1,4)&"-"&Mid(bill_month,5,2)&"-01"
end_date = DateValue(from_date)
end_date = DateAdd("m",1,from_date)
to_date = CStr(DateAdd("d",-1,end_date))

If bill_id = "" Then
	ck_sw = "y"
Else
	ck_sw = "n"
End If

If ck_sw = "n" Then
	Set filenm = uploadForm("att_file")(1)

	path = Server.MapPath("/large_file")
	filename = filenm.safeFileName
	fileType = Mid(filename,InStrRev(filename,".")+1)
	file_name = "e����"

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
		rs.Open "select * from [6:10000]",cn,"0"

		rowcount = -1
		xgr = rs.getrows
		rowcount = UBound(xgr,2)
		fldcount = rs.fields.count
		tot_cnt = rowcount + 1

		'�ʵ� ���� üũ
		If fldcount <> 33 Then
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

title_line = "�̼��θ��� ���ݰ�꼭 ���ε�"

' ������ǥ �Ϸ������� ����
bill_id = "1"

' 2019.02.15 �ڼ��� ��û 19�� ���� X,Y �÷��� ���� '��Ź����ڵ�Ϲ�ȣ','��ȣ' �� �߰��Ǿ���
' ��Ģ������ ���α׷��� �����ؾ� �ϳ� �ڼ��κ����� �� �� �÷��� �����ϰ� ���ε��ϰڴٰ� ��..
' ������ ����� �ٸ���(�ٸ�����)���� 	�ߵȴٰ� ��.. (������ �� ���� ��°��� �ǽ�..)
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
		<!--<script type="text/javascript" src="/java/js_window.js"></script>-->
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}

			$(document).ready(function(){
				var rowcnt = '<%=rowcount%>';
				var fldcnt = '<%=fldcount%>';

				//���ε� �׸� ���� Ȯ��
				//console.log(rowcnt);
				if(parseInt(rowcnt) > -1 && parseInt(fldcnt) !== 33){
					alert('���ε� �׸� ������ ��ġ���� �ʽ��ϴ�.(�ʼ� �׸� ����:33��)');
					location.href = '/finance/tax_bill_upload.asp';
					return;
				}
			});

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.bill_id.value == ""){
					alert ("��꼭 ������ �����ϼ���");
					return false;
				}

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
				a=confirm('DB�� ���ε� �Ͻðڽ��ϱ�?');

				if(a==true){
					document.frm.action = "/finance/tax_bill_upload_ok.asp";
					document.frm.submit();
				}
				return false;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/tax_bill_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/finance/tax_bill_upload.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ε峻��</dt>
						<dd>
							<p>
							<label>
								<strong>��꼭 ���� : </strong>
								<input type="radio" name="bill_id" value="1" <%if bill_id = "1" Then %>checked<%End If %> style="width:25px"/>����
								<input type="radio" name="bill_id" value="2" <%if bill_id = "2" Then %>checked<%End If %> style="width:25px"/>����
							</label>
							<label>
								<strong>��꼭 ������ : </strong>
								<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);"/>
							</label>
							<label>
								<strong>���ε����� : </strong>
								<input name="att_file" type="file" id="att_file" size="60" value="<%=att_file%>" style="text-align:left"/>
							</label>
							<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>"/>
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
							<col width="10%" >
							<col width="7%" >
							<col width="11%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="6%" >
							<col width="3%" >
							<col width="12%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�Ǽ�</th>
								<th scope="col">���</th>
								<th scope="col">������</th>
								<th scope="col">��꼭����ȸ��</th>
								<th scope="col">����ڹ�ȣ</th>
								<th scope="col">��ȣ</th>
								<th scope="col">��ǥ�ڸ�</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">û��</th>
								<th scope="col">��꼭�̸���</th>
								<th scope="col">ǰ���</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim tot_price, tot_cost, tot_cost_vat, reg_cnt, trade_no_err_cnt
						Dim i, rec_cnt, owner_cnt, bill_date, approve_no, trade_no, trade_name
						Dim trade_owner, owner_trade_no, price, cost, cost_vat, bill_collect
						Dim send_email, receive_email, tax_bill_memo, email_view
						Dim rs_etc, reg_sw, rs_trade, owner_sw, owner_company, trade_no_err
						Dim date_err_cnt, bill_date_err, cost_sum_err, sum_cost
						Dim cost_err_cnt, error_cnt, rsApprove, approve_no_err, approve_err_cnt

						tot_price = 0
						tot_cost = 0
						tot_cost_vat = 0

						reg_cnt = 0
						owner_cnt = 0
						'trade_no_err_cnt = 0

						'�� ���� ����
						error_cnt = 0

						'���ε� ������ ����
						If rowcount > -1 Then
							For i=0 To rowcount
								'���� ��ȣ üũ(���� ���� ������ ���� ��� rowcount ���ԵǹǷ� ���ι�ȣ�� üũ��)
								If f_toString(xgr(1,i), "") = "" Then
									Exit For
								End If

								'If xgr(0,i) => from_date And xgr(0,i) <= to_date Then
								bill_date = xgr(0,i)'�ۼ�����
								approve_no = xgr(1,i)'���ι�ȣ

								price = xgr(14,i)'�հ�ݾ�
								cost = xgr(15,i)'���ް���
								cost_vat = xgr(16,i)'����

								bill_collect = xgr(21,i)'����/û�� ����
								send_email = xgr(22,i)'�������̸���
								receive_email = xgr(23,i)'���޹޴����̸���1

								tax_bill_memo = xgr(26,i)'ǰ���

								'��꼭 ���� ����(����/����)
								If bill_id = "1" Then
									trade_no = xgr(4,i)'�����ڻ���ڵ�Ϲ�ȣ
									trade_name = xgr(6,i)'��ȣ(������)
									trade_owner = xgr(7,i)'��ǥ�ڸ�(������)
									owner_trade_no = xgr(9,i)'���޹޴��ڻ���ڵ�Ϲ�ȣ
									email_view = send_email
								Else
									owner_trade_no = xgr(4,i)'�����ڻ���ڵ�Ϲ�ȣ
									trade_no = xgr(9,i)'���޹޴��ڻ���ڵ�Ϲ�ȣ
									trade_name = xgr(11,i)'��ȣ(���޹޴���)
									trade_owner = xgr(12,i)'��ǥ�ڸ�(���޹޴���)
									email_view = receive_email
								End If

								tot_price = tot_price + price
								tot_cost = tot_cost + cost
								tot_cost_vat = tot_cost_vat + cost_vat

								'if bill_id = "1" then
								'	email_view = send_email
								'else
								'  	email_view = receive_email
								'end if

								'�ۼ����� ���� üũ
								If (bill_date < from_date Or bill_date > to_date) Or f_toString(bill_date, "") = "" Then
									date_err_cnt = date_err_cnt + 1
									bill_date_err = "Y"
								Else
									bill_date_err = "N"
								End If

								'�˻���� ���ι�ȣ �ߺ� �Ǽ� üũ
								objBuilder.Append "SELECT approve_no FROM tax_bill "
								objBuilder.Append "WHERE approve_no='"&approve_no&"' "
								objBuilder.Append "	AND REPLACE(SUBSTRING(bill_date,1,7),'-','')='"&bill_month&"';"

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
								objBuilder.Append "SELECT approve_no FROM tax_bill "
								objBuilder.Append "WHERE approve_no='"&approve_no&"' "
								objBuilder.Append "	AND REPLACE(SUBSTRING(bill_date,1,7),'-','')<'"&bill_month&"';"

								Set rsApprove = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsApprove.EOF Or rsApprove.BOF Then
									approve_no_err="N"
								Else
									approve_no_err="Y"
									approve_err_cnt=approve_err_cnt+1
								End If

								owner_trade_no = Replace(owner_trade_no,"-","")

								objBuilder.Append "SELECT trade_name FROM trade "
								objBuilder.Append "WHERE trade_no='"&owner_trade_no&"';"

								Set rs_trade = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rs_trade.EOF Or rs_trade.BOF Then
									owner_sw = "N"
									owner_cnt = owner_cnt + 1
									owner_company = owner_trade_no&"_Error"
								Else
									owner_sw = "Y"
									owner_company = rs_trade("trade_name")
								End If
								rs_trade.Close()

								'trade_no_err = "N"

								'�հ�ݾ� ���� üũ(�հ�ݾ�=���ް���+����)
								cost_sum_err = "N"
								sum_cost = CDbl(cost) + CDbl(cost_vat)

								If sum_cost <> CDbl(price) Then
									cost_err_cnt = cost_err_cnt + 1
									cost_sum_err = "Y"
								End If
						%>
							<tr <%If reg_sw = "Y" Then%>style="background-color:burlywood;"<%End If%>>
								<td class="first"><%=i+1%></td>
								<td <%If approve_no_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%If reg_sw = "Y" Then%>���<%Else%>�̵��<%End If%></td>
								<td <%If bill_date_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=bill_date%></td>
								<td <%If owner_sw = "N" Then%>bgcolor="#FFCCFF"<%End If%>><%=owner_company%></td>
								<td><%=trade_no%></td>
								<td><%=trade_name%></td>
								<td><%=trade_owner%></td>
								<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(price,0)%></td>
								<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(cost,0)%></td>
								<td <%If cost_sum_err = "Y" Then %>bgcolor="#FFCCFF"<%End If %> class="right"><%=FormatNumber(cost_vat,0)%></td>
								<td><%=bill_collect%></td>
								<td><%=email_view%>&nbsp;</td>
								<td class="left"><%=tax_bill_memo%></td>
						<%
							Next
							Set rs_etc = Nothing
							Set rs_trade = Nothing

							rs.Close() : Set rs = Nothing
							cn.Close() :  Set cn = Nothing

							'�� ���� ����
							error_cnt=date_err_cnt+approve_err_cnt+owner_cnt+cost_err_cnt

							DBConn.Close() : Set DBConn = Nothing
						Else
							Response.Write "<tr><td colspan='13' style='font-weight:bold;height:30px;'>��ȸ�� ������ �����ϴ�.</td></tr>"
						End If

						'����Ʈ �� ����
						rec_cnt = i
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>��</strong></td>
								<td class="right"><%=FormatNumber(reg_cnt,0)%></td>
								<td>&nbsp;</td>
								<td class="right"><%=FormatNumber(owner_cnt,0)%></td>
								<td class="right"><%=FormatNumber(trade_no_err_cnt,0)%></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"><%=FormatNumber(tot_price,0)%></td>
								<td class="right"><%=FormatNumber(tot_cost,0)%></td>
								<td class="right"><%=FormatNumber(tot_cost_vat,0)%></td>
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
								<td class="right"><%=FormatNumber(approve_err_cnt, 0)%> ��</td><!--���ι�ȣ �ߺ�(�˻��⵵ ����)-->
								<td class="right"><%=FormatNumber(date_err_cnt, 0)%> ��</td><!--�ۼ�����-->
								<td class="right"><%=FormatNumber(owner_cnt, 0)%> ��</td><!--����ȸ��-->
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right" colspan="3"><%=FormatNumber(cost_err_cnt, 0)%> ��</td><!--�հ�-->
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
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