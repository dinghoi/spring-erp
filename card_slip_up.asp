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

title_line = "카드 내역 업로드"

Dim att_file
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리 회계 시스템</title>
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
					alert ("카드유형을 선택하세요");
					return false;
				}

				if(document.frm.slip_month.value == "") {
					alert ("년월을 선택하세요");
					return false;
				}

				if(document.frm.att_file.value == "") {
					alert ("업로드 엑셀 파일을 선택하세요");
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
				if(confirm('DB에 업로드 하시겠습니까?') == true) {
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
					<legend>조회영역</legend>
					<dl>
						<dt>업로드내용</dt>
                        <dd>
                            <p>
								<label>
								<strong>카드유형 : </strong>
                                    <select name="card_gubun" id="card_gubun" style="width:80px">
                                        <option value="">선택</option>
                                        <option value="BC카드" <%If card_gubun = "BC카드" Then %>selected<%End If %>>BC카드</option>
                                        <option value="kb국민카드" <%If card_gubun = "kb국민카드" Then %>selected<%End If %>>kb국민카드</option>
                                        <option value="신한카드" <%If card_gubun = "신한카드" Then %>selected<%End If %>>신한카드</option>
                                        <option value="씨티카드" <%If card_gubun = "씨티카드" Then %>selected<%End If %>>씨티카드</option>
                                        <option value="롯데카드" <%If card_gubun = "롯데카드" Then %>selected<%End If %>>롯데카드</option>
                                    </select>
								</label>
								<label>
								<strong>전표년월 : </strong>
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
								<th class="first" scope="col">건수</th>
								<th scope="col">등록</th>
								<th scope="col">사용일</th>
								<th scope="col">카드유형</th>
								<th scope="col">카드번호</th>
								<th scope="col">사용인</th>
								<th scope="col">거래처</th>
								<th scope="col">업종</th>
								<th scope="col">계정과목</th>
								<th scope="col">적요</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
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

								'BC카드일 경우
								If card_gubun = "BC카드" Then
									If Trim(xgr(0, i)) = "신규" Then
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

								'//2017-06-08 add. kb국민카드
								If card_gubun = "kb국민카드" Then
									If Trim(xgr(15, i)) = "정상" Then
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
	'Response.write("<br>[["&VarType(xgr(10,i))&"]] ") ' 8 : String 1 : null 병신같은 asp!! 정수를 제대로 못읽는다. 어깨점을 일일이 붙이고 난뒤 저장을 해야..
									price = xgr(10, i)
									cost_vat = xgr(11, i)
									approve_no = xgr(14, i)

									If price = "" Or IsNull(price) Then
										price = "'"&xgr(11,i)
										response.write(price)
									End If
								End If

								'씨티카드
								If card_gubun = "씨티카드" Then
									If Trim(xgr(14, i)) = "정상" Then
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

								' 신한카드 	L(9410-6440-9)
								If card_gubun = "신한카드" Then

									'신규 작성[허정호_20201215]	=======================
									slip_date = replace(xgr(6,i),".","-")

									'imsi_no = xgr(0,i)
									'card_no = mid(imsi_no,2,3) & "-" &right(imsi_no,4)
									card_no = xgr(0, i)	'이용 카드(카드 번호)

									customer = xgr(19,i)
									imsi_no = xgr(18,i)
									customer_no = mid(imsi_no,1,3) & "-" & mid(imsi_no,4,2) & "-" &right(imsi_no,5)
									upjong = replace(xgr(20,i)," ","")
									price = xgr(9,i)
									cost_vat = xgr(12,i)
									approve_no = xgr(5,i)

									'승인일자는 이용일시로 적용[허정호_20201217]
									'slip_date = Replace(Left(xgr(0, i), 10), ".", "-")
									'approve_no = xgr(2, i)	'승인번호
									'card_no = xgr(3, i)	'이용 카드(카드 번호)
									'customer = xgr(5, i)	'가맹점 명
									'customer_no = ""	'가맹점 번호
									'upjong = xgr(6, i)	'업종
									'price = xgr(7, i)	'이용금액

									'부가세 계산
									'If xgr(10,i) = "국내" Then
									'	cost_vat = price - Int(price/1.1)
									'Else
									'	cost_vat = 0
									'End If

									'취소 여부
									If price < 0 Then
										cancel_yn = "Y"
									Else
										cancel_yn = "N"
									End If
								End If

			                    ' 롯데카드 	LOCAL -> 첫 4자리 9409, AMEX -> 첫 4자리 3762 , VISA -> 첫 4자리 4670
								If card_gubun = "롯데카드" Then
									slip_date = Replace(xgr(5, i), ".", "-")
									imsi_no = xgr(2, i)
									imsi_card_no = Right(imsi_no, 3)

									sql = "select * from card_owner where card_type like '%롯데%' and right(card_no,3) = '"&imsi_card_no&"'"
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

									If Trim(xgr(12, i)) = "매입여부" Then
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

									'카드 사용 내역 조회
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
										account = "접대비"
										account_item = "접대비"
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

									If card_gubun = "신한카드" Then
										'카드 번호 조회_앞 8자리 비교 조회에서 앞 4자리, 뒤 4자리로 변경[허정호_20201217]
										objBuilder.Append "WHERE RIGHT(cdot.card_no, 4) = '" & Right(card_no, 4) &"' "
										objBuilder.Append "AND LEFT(cdot.card_no, 4) = '" & Left(card_no, 4) &"'"
									ElseIf card_gubun = "kb국민카드" Then
										' 20180727 수정 - 카드번호 입력이 잘못 될 수도 있다..
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
										card_type = "미등록"
										emp_name = "미지정"
										emp_no = ""
										car_vat_sw = "C"
									Else
										card_no = rs_card("card_no")
										car_vat_sw = rs_card("car_vat_sw")
										card_type = rs_card("card_type")
										emp_no = rs_card("emp_no")

										'NKP 사용자명 조회
										objBuilder.Append "SELECT user_name "
										objBuilder.Append "FROM memb "
										objBuilder.Append "WHERE user_id = '"&emp_no&"' "

										Set rs_emp = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()

										If rs_emp.EOF Or rs_emp.BOF Then
											emp_name = "사번누락"
										Else
											emp_name = rs_emp("user_name")
										End If

										'퇴사자 아이디 체크 추가[허정호_20210716]
										If rs_card("emp_pay_id") = "2" Then
											owner_sw = "N"
											tot_err = tot_err + 1
											emp_name = "사번오류(퇴사))"
										End If

										rs_emp.Close
									End If

									' 주유카드가 바뀌면 수정 해야함
									If card_type = "롯데주유카드" Then
										account = "차량유지비"
										account_item = "유류대"
										' 2014년 12월부터 변경
										' car_vat_sw = "Y"
									End If
									' 주유카드 변경 끝

									If account = "차량유지비" Then
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
									<!--등록-->
									<%'기존 카드 번호 등록 여부
									If reg_sw = "N" Then
										Response.Write "미등록"
									Else
										Response.Write "등록"
									End If
									%>
									</td>
									<!--사용일-->
									<td <%If date_sw = "N" Then%>bgcolor="#FFCCFF" <%End If%>><%=slip_date%></td>
									<!--카드유형-->
									<td><%=card_type%></td>
									<!--카드번호-->
									<td><%=card_no%></td>
									<!--사용인-->
									<td <%If owner_sw = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=emp_name%></td>
									<!--거래처-->
									<td><%=customer%></td>
									<!--업종-->
									<td <%If upjong_sw = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=upjong%></td>
									<!--계정과목-->
									<td <%If account_sw = "Y" Then%>bgcolor="#FFCCFF"<%End If%>><%=account%></td>
									<!--적요-->
									<td><%=account_item%>&nbsp;</td>
									<!--합계-->
									<td class="right"><%=FormatNumber(price, 0)%></td>
									<!--공급가액-->
									<td class="right"><%=FormatNumber(cost, 0)%></td>
									<!--부가세-->
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
								<th class="first">계(장애)</th>
								<th>&nbsp;</th>
								<th><%=FormatNumber(date_err, 0)%></th>
								<th><%=FormatNumber(tot_err, 0)%></th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th>&nbsp;</th>
								<th><%=FormatNumber(tot_upjong, 0)%></th>
								<th><%=FormatNumber(tot_account, 0)%></th>
								<th>계(금액)</th>
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
                        <span class="btnType01"><input type="button" value="DB저장" onclick="javascript:frm1check();"NAME="Button1"></span>
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
