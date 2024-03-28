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
Dim uploadForm, filenm, slip_month, from_date, end_date, to_date, file_type
Dim cost_year, ck_sw, Path, fileName, fileType, file_name, save_path
Dim objFile, rowCount, title_line, att_file
Dim cn, rs, xgr, fldCount, tot_cnt
Dim end_saupbu, rs_end, new_date

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")
uploadForm.AbsolutePath = True
uploadForm.Overwrite = True
uploadForm.MaxUploadSize = 1024*1024*50

slip_month = uploadForm("slip_month")
file_type = uploadForm("file_type")

If slip_month = "" Then
	slip_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)
End If

if saupbu = "" then
	end_saupbu = "사업부외나머지"
else
  	end_saupbu = saupbu
end if

'마감 일자
objBuilder.Append "SELECT MAX(end_month) AS max_month "
objBuilder.Append "FROM cost_end "
objBuilder.Append "WHERE saupbu = '"&end_saupbu&"' AND end_yn = 'Y' "

set rs_end = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rs_end("max_month")) Then
	end_date = "2014-08-31"
Else
	new_date = DateAdd("m", 1, DateValue(Mid(rs_end("max_month"), 1, 4) & "-" & Mid(rs_end("max_month"), 5, 2) & "-01"))
	end_date = DateAdd("d", -1, new_date)
End If
rs_end.close() : Set rs_end = Nothing

cost_year = Mid(slip_month, 1, 4)

If slip_month = "" then
	ck_sw = "y"
else
	ck_sw = "n"
end if

If ck_sw = "n" Then
	Set filenm = uploadForm("att_file")(1)

	path = Server.MapPath ("/large_file")
	filename = filenm.safeFileName
	fileType = mid(filename,inStrRev(filename,".")+1)
	file_name = "일반경비"

'		save_path = path & "\" & filename
	save_path = path & "\" & file_name&"."&fileType

	If fileType = "xls" Or fileType = "xlk" Then
		file_type = "Y"
		filenm.save save_path

		objFile = save_path
'		objFile = Request.form("att_file")
'		objFile = SERVER.MapPath("att_file")
'		objFile = SERVER.MapPath(".") & "\kwon_upload\excel_data.xls"
'		response.write(objFile)

		set cn = Server.CreateObject("ADODB.Connection")
		set rs = Server.CreateObject("ADODB.Recordset")

		cn.open "Driver={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & objFile & ";"
		rs.Open "select * from [1:10000]", cn, "0"

		rowcount = -1
		xgr = rs.getrows
		rowcount = UBound(xgr, 2)
		fldCount = rs.fields.count
		tot_cnt = rowcount + 1
	Else
		objFile = "none"
		rowcount = -1
		file_type = "N"
	End If
Else
	objFile = "none"
	rowcount=-1
End If

title_line = "영업 및 관리부서 경비 업로드"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
        <script src="/java/jquery-1.9.1.js"></script>
        <script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 2";
			}

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
//				if (document.frm.bill_id.value == "") {
//					alert ("계산서 유형을 선택하세요");
//					return false;
//				}
				if(document.frm.slip_month.value == ""){
					alert ("년월을 선택하세요");
					return false;
				}
				if(document.frm.att_file.value == ""){
					alert ("업로드 엑셀 파일을 선택하세요");
					return false;
				}
				return true;
			}

			function upload_ok(){
				a=confirm('DB에 업로드 하시겠습니까?');

				if(a === true){
					//document.frm.action = "/sales_bill_upload_ok.asp";
					document.frm.action = "/cost/general_cost_excel_upload_ok.asp";
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
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/general_cost_excel_upload.asp" method="post" name="frm" enctype="multipart/form-data">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>업로드내용</dt>
                        <dd>
                            <p>
								<label>
								<strong>발생년월 : </strong>
                                	<input name="slip_month" type="text" value="<%=slip_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);">
								</label>
                                <label>
								<strong>업로드파일 : </strong>
								<input name="att_file" type="file" id="att_file" size="60" value="<%=att_file%>" style="text-align:left">
								</label>
            					<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>">
            					<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
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
							<col width="*" >
							<col width="6%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">건수</th>
								<th scope="col">발생일자</th>
								<th scope="col">사용조직</th>
								<th scope="col">사용자</th>
								<th scope="col">회사</th>
								<th scope="col">비용항목</th>
								<th scope="col">비용상센</th>
								<th scope="col">사용구분/금액</th>
								<th scope="col">고객사</th>
								<th scope="col">상호명(가게이름)</th>
								<th scope="col">정산여부</th>
								<th scope="col">사용내역</th>
								<th scope="col">손익포함</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim reg_cnt, trade_no_err_cnt, error_cnt
						Dim i, rec_cnt
						Dim slip_date, account, price, company
						Dim customer, pay_yn, slip_memo, pl_yn
						Dim emp_name, account_name, account_item

						Dim date_err, date_err_cnt, org_name_err, org_name_errCnt
						Dim emp_name_err, emp_name_errCnt, emp_company_err, emp_company_errCnt
						Dim account_name_err, account_name_errCnt, account_item_err, account_item_errCnt
						Dim price_err, price_errCnt, customer_err, customer_errCnt
						Dim pay_yn_err, pay_yn_errCnt, pl_yn_err, pl_yn_cnt
						Dim company_err, company_err_cnt

						Dim rsOrg

						error_cnt = 0

						date_err_cnt = 0
						org_name_errCnt = 0
						emp_name_errCnt = 0
						emp_company_errCnt = 0
						account_name_errCnt = 0
						account_item_errCnt = 0
						price_errCnt = 0
						customer_errCnt = 0
						pay_yn_errCnt = 0
						pl_yn_cnt = 0

						if rowCount > -1 Then
							for i=0 to rowcount
								if xgr(0,i) = "" or isnull(xgr(0,i)) then
									exit for
								end If

								slip_date = xgr(0,i)
								org_name = xgr(1,i)
								emp_name = xgr(2,i)
								emp_company = xgr(3,i)
								account_name = xgr(4,i)
								account_item = xgr(5,i)
								price = toString(xgr(6,i), 0)
								company = xgr(7,i)
								customer = xgr(8,i)
								pay_yn = xgr(9,i)
								slip_memo = xgr(10,i)
								pl_yn = xgr(11,i)

								'발생일자
								If slip_date = "" Or IsNull(slip_date) Then
									date_err = "Y"
									date_errCnt = date_errCnt + 1
								Else
									date_err = "N"
								End If

								'사용조직
								If org_name = "" Or IsNull(org_name) Then
									org_name_err = "Y"
									org_name_errCnt = org_name_errCnt + 1
								Else
									objBuilder.Append "SELECT org_name FROM emp_org_mst "
									objBuilder.Append "WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
									objBuilder.Append "	AND org_name = '"&org_name&"' AND org_company = '"&emp_company&"' "
									Set rsOrg = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rsOrg.BOF Or rsOrg.EOF Then
										org_name_err = "Y"
										error_cnt = error_cnt + 1
									Else
										org_name_err = "N"
									End If
								End If

								'사용자
								If emp_name = "" Or IsNull(emp_name) Then
									emp_name_err = "Y"
									emp_name_errCnt = emp_name_errCnt + 1
								Else
									objBuilder.Append "SELECT emp_no FROM emp_master "
									objBuilder.Append "WHERE emp_name = '"&emp_name&"' "
									Set rsOrg = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rsOrg.BOF Or rsOrg.EOF Then
										emp_name_err = "Y"
										error_cnt = error_cnt + 1
									Else
										emp_name_err = "N"
									End If
								End If

								'회사
								If emp_company = "" Or IsNull(emp_company) Then
									emp_company_err = "Y"
									emp_company_errCnt = emp_company_errCnt + 1
								Else
									emp_company_err = "N"
								End If

								'비용항목
								If account_name = "" Or IsNull(account_name) Then
									account_name_err = "Y"
									account_name_errCnt = account_name_errCnt + 1
								Else
									account_name_err = "N"
								End If

								'비용상세
								If account_item = "" Or IsNull(account_item) Then
									account_item_err = "Y"
									account_item_errCnt = account_item_errCnt + 1
								Else
									account_item_err = "N"
								End If

								'사용구분/금액
								If price = "" Or IsNull(price) Then
									price_err = "Y"
									price_errCnt = price_errCnt + 1
								Else
									price_err = "N"
								End If

								'상호명(가게이름)
								If company = "" Or IsNull(company) Then
									company_err = "Y"
									company_err_cnt = company_err_cnt + 1
								Else
									company_err = "N"
								End If

								'고객사
								If customer = "" Or IsNull(customer) Then
									customer_err = "Y"
									customer_errCnt = customer_errCnt + 1
								Else
									customer_err = "N"
								End If

								'정산여부
								If pay_yn = "" Or IsNull(pay_yn) Then
									pay_yn_err = "Y"
									pay_yn_errCnt = pay_yn_errCnt + 1
								Else
									pay_yn_err = "N"
								End If

								'손익포함
								If pl_yn = "" Or IsNull(pl_yn) Then
									pl_yn_err = "Y"
									pl_yn_errCnt = pl_yn_errCnt + 1
								Else
									pl_yn_err = "N"
								End If

								reg_cnt = reg_cnt + 1

								%>
								<tr>
									<td class="first"><%=i+1%></td>
									<td <%If date_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=slip_date%></td>
									<td <%If org_name_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=org_name%></td>
									<td <%If emp_name_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=emp_name%></td>
									<td <%If emp_company_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=emp_company%></td>
									<td <%If account_name_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=account_name%></td>
									<td <%If account_item_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=account_item%></td>
									<td <%If price_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=FormatNumber(price, 0)%></td>
									<td <%If company_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=company%></td>
									<td><%=customer%></td>
									<td <%If pay_yn_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=pay_yn%></td>
									<td><%=slip_memo%></td>
									<td <%If pl_yn_err = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=pl_yn%></td>
								</tr>
								<%

							next
						end if
						rec_cnt = i
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>계</strong></td>
								<td class="right"><%=formatnumber(reg_cnt,0)%></td>
								<td><strong>Error</strong></td>
								<td class="right"><%=formatnumber(error_cnt,0)%>건</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right"></td>
								<td class="right"></td>
								<td class="right"></td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>

							</tr>
						</tbody>
					</table>
				</div>
				<%
				if error_cnt = 0 and rowcount > -1 then %>
					<br>
                    <div align="center">
                        <span class="btnType01"><input type="button" value="DB에 업로드" onclick="javascript:upload_ok();"NAME="upload_btn"></span>
                    </div>
				<% end if %>
					<br>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>">
				</form>
		</div>
	</div>
	</body>
</html>

