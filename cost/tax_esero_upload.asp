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
	file_name = "e세로일괄처리"

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

		'필드 개수 체크
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

title_line = "E세로 비용일괄업로드"

' 2019.02.15 박성민 요청 19년 부터 X,Y 컬럼에 각각 '수탁사업자등록번호','상호' 가 추가되었음
' 원칙적으로 프로그램을 수정해야 하나 박성민본인이 이 두 컬럼을 삭제하고 업로드하겠다고 함..
' 이유을 물어보니 다른곳(다른엑셀)에선 	잘된다고 함.. (엑셀이 다 따로 노는건지 의심..)
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
		<!--<script type="text/javascript" src="/java/js_window.js"></script>-->
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}

			$(document).ready(function(){
				var rowcnt = '<%=rowcount%>';
				var fldcnt = '<%=fldcount%>';

				//업로드 항목 개수 확인
				console.log(rowcnt);
				console.log(fldcnt);
				if(parseInt(rowcnt) > -1 && parseInt(fldcnt) !== 14){
					alert('업로드 항목 개수가 일치하지 않습니다.(필수 항목 개수:14개)');
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
				var result = confirm('DB에 업로드 하시겠습니까?');

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
					<legend>조회영역</legend>
					<dl>
						<dt>업로드내용</dt>
						<dd>
							<p>
							<label>
								<strong>계산서 발행년월 : </strong>
								<input name="bill_month" type="text" value="<%=bill_month%>" maxlength="6" size="6" onKeyUp="checkNum(this);"/>
							</label>
							<label>
								<strong>업로드파일 : </strong>
								<input name="att_file" type="file" id="att_file" size="60" value="<%=att_file%>" style="text-align:left"/>
							</label>
							<input name="file_type" type="hidden" id="file_type" value="<%=file_type%>"/>
							<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"/></a>
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
								<th class="first" scope="col">건수</th>
								<th scope="col">발행일자</th>
								<th scope="col">계산서소유회사</th>
								<th scope="col">상호명</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">거래내역</th>
								<th scope="col">담당자</th>
								<th scope="col">사용조직코드</th>
								<th scope="col">고객사</th>
								<th scope="col">담당사업부</th>
								<!--<th scope="col">비용유형</th>-->
								<th scope="col">세부유형</th>
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
						org_code_cnt = 0'사용조직코드 오류 개수
						owner_cnt = 0'거래처 오류 개수
						saupbu_cnt = 0'사업부 오류 개수
						slip_cnt = 0
						cost_err_cnt = 0

						emp_name_cnt = 0

						error_cnt = 0'총 에러 개수

						tot_price = 0
						tot_cost = 0
						tot_cost_vat = 0

						'업로드 데이타 개수
						If rowcount > -1 Then
							For i=0 To rowcount
								'비용등록 값 체크(사용조직코드, 고객사,담당사업부, 비용유형, 세부유형 값 체크)
								'If f_toString(xgr(1,i), "") = "" Then
								'	Exit For
								'End If

								t_bill_date = f_toString(xgr(0,i), "")'발행일자
								't_approve = f_toString(xgr(1,i), "")'승인번호
								t_owner_company = f_toString(xgr(2,i), "")'계산서소유회사
								t_trade_name = f_toString(xgr(3,i), "")'상호명
								t_price = f_toString(xgr(4,i), 0)'합계
								t_cost = f_toString(xgr(5,i), 0)'공급가액
								t_cost_vat = f_toString(xgr(6,i), 0)'부가세
								t_tax_bill_memo = f_toString(xgr(7,i), "")'거래내역

								't_emp_no = f_toString(xgr(8,i), "")'담당사사번
								t_emp_name = f_toString(xgr(9,i), "")'담당자

								t_org_code = f_toString(xgr(10,i), "")'사용조직코드
								t_company = f_toString(xgr(11,i), "")'고객사
								t_mg_saupbu = f_toString(xgr(12,i), "")'담당사업부
								't_slip_gubun = f_toString(xgr(13,i), "")'비용유형
								t_account_str = f_toString(xgr(13,i), "")'세부유형

								If t_bill_date <> "" Then
									t_bill_date = CStr(t_bill_date)
								End If

								tot_err = "N"'전체 에러
								bill_date_err = "N"

								'빌헹일자 에러 체크
								If (t_bill_date < from_date Or t_bill_date > to_date) Or f_toString(t_bill_date, "") = "" Then
									date_err_cnt = date_err_cnt + 1
									bill_date_err = "Y"

									tot_err = "Y"
								End If

								org_code_err = "N"'사용조직코드 체크 코드

								If t_org_code = "" And (t_company <> "" Or t_mg_saupbu <> "" Or t_account_str <> "") Then
									org_code_err = "Y"
									org_code_cnt = org_code_cnt + 1

									tot_err = "Y"
								End If

								'비용 등록 사용자 조회
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

								owner_err = "N"'고객사 체크 코드

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

								saupbu_err = "N"'담당사업부 체크 코드

								If t_mg_saupbu = "" And (t_org_code <> "" Or t_company <> "" Or t_account_str <> "") Then
									saupbu_err = "Y"
									saupbu_cnt = saupbu_cnt + 1

									tot_err = "Y"
								End If

								'비용유형 체크
								slip_err = "N"'


								'비용유형 설정
								If t_account_str = "" And (t_org_code <> "" Or t_company <> "" Or t_mg_saupbu <> "") Then
									t_account = ""
									t_account_item = ""

									slip_err = "Y"
									slip_cnt = slip_cnt + 1

									tot_err = "Y"
								Else
									arr_str = Split(t_account_str, ")")'세부유형

									For j = 0 To UBound(arr_str)
										If j = 0 Then
											slip_gubun = Replace(arr_str(j), "(", "")
										Else
											slip_account = arr_str(j)
										End If
									Next

									If slip_gubun = "비용" Then
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

								'합계금액 에러 체크(합계금액=공급가액+세액)
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

							'총 에러 개수
							error_cnt = date_err_cnt + org_code_cnt + owner_cnt + saupbu_cnt + slip_cnt + cost_err_cnt + emp_name_cnt

							DBConn.Close() : Set DBConn = Nothing
						Else
							Response.Write "<tr><td colspan='13' style='height:30px;'>조회된 내역이 없습니다.</td></tr>"
						End If

						'리스트 총 개수
						'rec_cnt = i
						%>
							<tr bgcolor="#FFE8E8">
								<td class="first"><strong>계</strong></td>
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
						'에러 건수
						If error_cnt > 0 Then
						%>
							<tr bgcolor="#FFCCFF">
								<td class="first"><strong>Error</strong></td>
								<td class="right"><%=FormatNumber(date_err_cnt, 0)%> 건</td><!--발행일자-->
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td>&nbsp;</td>
								<td class="right" colspan="3"><%=FormatNumber(cost_err_cnt, 0)%> 건</td><!--합계-->
								<!--<td>&nbsp;</td>-->
								<td class="right"><%=FormatNumber(emp_name_cnt, 0)%> 건</td>
								<td class="right"><%=FormatNumber(org_code_cnt, 0)%> 건</td>
								<td class="right"><%=FormatNumber(owner_cnt, 0)%> 건</td>
								<td class="right"><%=FormatNumber(saupbu_cnt, 0)%> 건</td>
								<td class="right"><%=FormatNumber(slip_cnt, 0)%> 건</td>
							</tr>
						<%End If%>
						</tbody>
					</table>
				</div>
				<%
				'DB Upload 노출 조건
				'If reg_cnt <> rec_cnt And owner_cnt = 0 And trade_no_err_cnt = 0 And rowcount > -1 Then
				If rowcount > -1 And error_cnt = 0 Then
				%>
					<br>
                    <div align="center">
                        <span class="btnType01"><input type="button" value="DB에 업로드" onclick="javascript:upload_ok();"/></span>
                    </div>
				<%End If %>
					<br>
                    <input name="objFile" type="hidden" id="objFile" value="<%=objFile%>"/>
				</form>
		</div>
	</div>
	</body>
</html>