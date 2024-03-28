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
Dim month_tab(24, 2)
Dim page, view_condi, pmg_yymm, pmg_yymm_to, to_date, be_pg
Dim curr_dd, from_date, datYear, datMonth, datLastDay, exec_LastDay
Dim st_in_date, rever_year, cal_month, view_month, i, j, cal_year
Dim pgsize, start_page, stpage, long_hap, total_record, total_page
Dim rsInsEmp, epi_emp, epi_com, rsInsHap, rsCount, rsPay, title_line
Dim rsOrg, pg_url

Dim arrPay

Dim emp_name, emp_first_date, emp_in_date, emp_type, emp_grade, emp_position
Dim emp_bonbu, emp_saupbu, emp_team, emp_org_code, emp_org_name, emp_reside_place, emp_reside_company
Dim pmg_emp_no, pmg_base_pay, pmg_give_total, de_emp_no, de_deduct_total, pmg_curr_pay
Dim incom_base_pay, incom_meals_pay, incom_overtime_pay, incom_month_amount
Dim incom_nps_amount, incom_nhis_amount, incom_nps, incom_nhis, incom_go_yn
Dim incom_long_yn, incom_wife_yn, incom_age20, incom_age60, incom_old
Dim pmg_tax_yes, pmg_tax_no, incom_family_cnt, inc_st_amt, inc_incom
Dim rs_sod, de_income_tax, de_nps_amt, de_nhis_amt, long_amt, de_longcare_amt
Dim we_tax, de_wetax, pmg_meals_pay, pmg_overtime_pay, epi_amt, de_epi_amt

page = f_Request("page")
view_condi = f_Request("view_condi")
pmg_yymm = f_Request("pmg_yymm")
pmg_yymm_to = f_Request("pmg_yymm_to")
to_date = f_Request("to_date")

be_pg = "/pay/insa_pay_month_batch.asp"
title_line = " 급여기초이월 처리 "

If view_condi = "" Then
	view_condi = "전체"
	curr_dd = CStr(DatePart("d", Now()))
	from_date = Mid(CStr(Now() - curr_dd + 1), 1, 10)
	pmg_yymm_to = Mid(CStr(from_date), 1, 4)&Mid(CStr(from_date), 6, 2)
	pmg_yymm = CStr(Mid(DateAdd("m", -1, Now()), 1, 4))&CStr(Mid(DateAdd("m", -1, Now()), 6, 2))

	'매월 말일 구하기
	datYear = Mid(CStr(pmg_yymm_to), 1, 4)
	datMonth = Mid(CStr(pmg_yymm_to), 5, 2)

	If datMonth = 4 Or datMonth = 6 Or datMonth=9 Or datMonth=11 Then  '4월 6월 9월 11월이면 월말값은 30일
		datLastDay = 30
	ElseIf datMonth = 2 And Not (datYear Mod 4) = 0 Then  '2월이고  년도를 4로 나눈 값이 0이 아니면 28일
		datLastDay = 28
	ElseIf datMonth = 2 And (datYear Mod 4) = 0 Then '윤달 계산
		If (datYear Mod 100) = 0 Then
			If (datYear Mod 400) = 0 Then
				datLastDay=29
			Else
				datLastDay=28
			End If
		Else
			datLastDay=29
		End If
	Else
		datLastDay=31
	End If

	exec_LastDay = datLastDay
	to_date = Mid(CStr(pmg_yymm_to), 1, 4)&"-"&Mid(CStr(pmg_yymm_to), 5, 2)&"-"&CStr(exec_LastDay)
	'to_date = ""
End If

'당월 입사/퇴사일이 15일 이전이면 당월 급여대상임
'st_es_date = Mid(CStr(pmg_yymm_to), 1, 4)&"-"&Mid(CStr(pmg_yymm_to), 5, 2)&"-"&"01"
st_in_date = Mid(CStr(pmg_yymm_to), 1, 4)&"-"&Mid(CStr(pmg_yymm_to), 5, 2)&"-"&"16"
rever_year = Mid(CStr(pmg_yymm_to), 1, 4) '귀속년도

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = Mid(CStr(Now()), 1, 4)&Mid(CStr(Now()), 6, 2)
month_tab(24, 1) = cal_month
view_month = Mid(cal_month, 1, 4)&"년 "&mid(cal_month, 5, 2)&"월"
month_tab(24, 2) = view_month

For i = 1 To 23
	cal_month = CStr(Int(cal_month) - 1)

	If Mid(cal_month, 5) = "00" Then
		cal_year = CStr(Int(Mid(cal_month, 1, 4)) - 1)
		cal_month = cal_year&"12"
	End If

	view_month = Mid(cal_month, 1, 4)&"년 "&Mid(cal_month, 5, 2)&"월"
	j = 24 - i
	month_tab(j, 1) = cal_month
	month_tab(j, 2) = view_month
Next

pgsize = 10 ' 화면 한 페이지

If page = "" Then
	page = 1
	start_page = 1
End If

stpage = Int((page - 1) * pgsize)
pg_url = "&view_condi="&view_condi&"&pmg_yymm="&pmg_yymm&"&pmg_yymm_to="&pmg_yymm_to&"&to_date="&to_date

'고용보험(실업) 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5503' and insu_class = '01'"
objBuilder.Append "SELECT emp_rate, com_rate FROM pay_insurance "
objBUilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5503' AND insu_class = '01';"

Set rsInsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsInsEmp.EOF Then
	epi_emp = FormatNumber(rsInsEmp("emp_rate"), 3)
	epi_com = FormatNumber(rsInsEmp("com_rate"), 3)
Else
	epi_emp = 0
	epi_com = 0
End If
rsInsEmp.Close() : Set rsInsEmp = Nothing

'장기요양보험 요율
'Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5504' and insu_class = '01'"
objBuilder.Append "SELECT hap_rate FROM pay_insurance "
objBuilder.Append "WHERE insu_yyyy = '"&rever_year&"' AND insu_id = '5504' AND insu_class = '01';"

Set rsInsHap = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsInsHap.EOF Then
	long_hap = FormatNumber(rsInsHap("hap_rate"), 3)
Else
	long_hap = 0
End If
rsInsHap.Close() : Set rsInsHap = Nothing

objBuilder.Append "SELECT COUNT(*) FROM emp_master "
objBuilder.Append "WHERE (ISNULL(emp_end_date) OR emp_end_date = '1900-01-01' OR emp_end_date >= '"&st_in_date&"') "
objBuilder.Append "	AND emp_in_date < '"&st_in_date&"' "
objBuilder.Append "	AND emp_pay_id <> '5' AND emp_no < '900000' "

If view_condi <> "전체" Then
	objBuilder.Append "	AND emp_company = '"&view_condi&"' "
End If

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount

rsCount.Close() : Set rsCount = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
			/*
		    $(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%'=from_date%>" );
			});
			*/
			$(function(){
				$( "#to_date" ).datepicker();
				$( "#to_date" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#to_date" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck(){
				if(formcheck(document.frm)){
					document.frm.submit();
				}
			}

			function delcheck () {
				if(form_chk(document.frm_del)){
					document.frm_del.submit();
				}
			}

			function form_chk(){
				var result = confirm('삭제하시겠습니까?');

				if(result == true){
					return true;
				}
				return false;
			}

			function pay_month_transe(val, val2, val3, val4){
				var tVal = document.getElementById(val).value;
				var tVal2 = document.getElementById(val2).value;
				var tVal3 = document.getElementById(val3).value;
				var tVal4 = document.getElementById(val4).value;

				if(tVal == null || tVal == ""){
					alert("이월대상년월을 선택해 주세요.");
					return;
				}

				if(tVal2 == null || tVal2 == ""){
					alert("회사를 선택해 주세요.");
					return;
				}

				if(tVal3==null || tVal3==""){
					alert("귀속년월을 선택해 주세요.");
					return;
				}

				if(tVal4 == null || tVal4 == ""){
					alert("지급일을 선택해 주세요.");
					return;
				}

				if(!confirm("전월 급여를 이월처리 하시겠습니까?")) return;

				var frm = document.frm;

				document.frm.pmg_yymm1.value = tVal;
				document.frm.view_condi1.value = tVal2;
				document.frm.pmg_yymm_to1.value = tVal3;
				document.frm.to_date1.value = tVal4;

				document.frm.action = "/pay/insa_pay_month_transe_save.asp";
				document.frm.submit();
            }
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3><br/>
				<form action="<%=be_pg%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								'Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
								objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
								objBuilder.Append "	AND org_level = '회사' AND org_code <> '6272' "
								objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

	                            Set rsOrg = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px;">
                                    <option value="전체" <%If view_condi = "전체" Then %>selected<%End If %>>전체</option>
                			  <%
								Do Until rsOrg.EOF
			  				  %>
                					<option value='<%=rsOrg("org_name")%>' <%If view_condi = rsOrg("org_name") Then %>selected<% End If %>><%=rsOrg("org_name")%></option>
                			  <%
									rsOrg.MoveNext()
								Loop
								rsOrg.Close() : Set rsOrg = Nothing
							  %>
            					</select>
                                </label>
                                <label>
								<strong>이월대상년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px;">
                                    <%For i = 24 To 1 Step -1	%>
										<option value="<%=month_tab(i, 1)%>" <%If pmg_yymm = month_tab(i, 1) Then %>selected<%End If %>><%=month_tab(i, 2)%></option>
                                    <%Next	%>
                                 </select>
								</label>
                                <label>
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm_to" id="pmg_yymm_to" type="text" value="<%=pmg_yymm_to%>" style="width:90px;">
                                    <%For i = 24 To 1 Step -1	%>
										<option value="<%=month_tab(i, 1)%>" <%If pmg_yymm_to = month_tab(i, 1) Then %>selected<%End If %>><%=month_tab(i, 2)%></option>
                                    <%Next	%>
                                 </select>
								</label>
                                <label>
								<strong>귀속지급일 : </strong>
                                	<input name="to_date" id="to_date" type="text" value="<%=to_date%>" style="width:70px;">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
                                    '당월 입사/퇴사일이 15일 이전이면 당월 급여대상임
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="9%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
								<th scope="col">최초입사일</th>
                                <th scope="col">입사일</th>
                                <th scope="col">소속</th>
								<th scope="col">기본급</th>
                                <th scope="col">지급액계</th>
                                <th scope="col">공제액계</th>
                                <th scope="col">차인지급액</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
							</tr>
						</thead>
						<tbody>
						<%
						'급여 정보 조회
						objBuilder.Append "SELECT emtt.emp_no, emtt.emp_name, emtt.emp_first_date, emtt.emp_in_date, "
						objBuilder.Append "	emtt.emp_type, emtt.emp_grade, emtt.emp_position, emtt.emp_company, emtt.emp_bonbu, emtt.emp_saupbu, "
						objBuilder.Append "	emtt.emp_team, emtt.emp_org_code, emtt.emp_org_name, emtt.emp_reside_place, emtt.emp_reside_company, "

						objBuilder.Append "	pmgt.pmg_emp_no, pmgt.pmg_base_pay, pmgt.pmg_give_total, "

						objBuilder.Append "	pmdt.de_deduct_total, "

						objBuilder.Append "	pyit.incom_base_pay, pyit.incom_meals_pay, pyit.incom_overtime_pay, pyit.incom_month_amount, "
						objBuilder.Append "	pyit.incom_family_cnt, pyit.incom_nps_amount, pyit.incom_nhis_amount, pyit.incom_nps, "
						objBuilder.Append "	pyit.incom_nhis, pyit.incom_wife_yn, pyit.incom_age20, pyit.incom_age60, pyit.incom_old, "

						objBuilder.Append "	pyit.incom_go_yn, pyit.incom_long_yn "

						objBUilder.Append "FROM emp_master AS emtt "
						objBuilder.Append "LEFT OUTER JOIN pay_month_give AS pmgt ON emtt.emp_no = pmgt.pmg_emp_no "
						objBuilder.Append "	AND emtt.emp_company = pmgt.pmg_company "
						objBuilder.Append "	AND pmgt.pmg_yymm = '"&pmg_yymm&"' AND pmgt.pmg_id = '1' "
						objBuilder.Append "LEFT OUTER JOIN pay_month_deduct AS pmdt ON emtt.emp_no = pmdt.de_emp_no "
						objBuilder.Append "	AND emtt.emp_company = pmdt.de_company "
						objBuilder.Append "	AND pmdt.de_yymm = '"&pmg_yymm&"' AND pmdt.de_id = '1' "
						objBuilder.Append "LEFT OUTER JOIN pay_year_income AS pyit ON emtt.emp_no = pyit.incom_emp_no "
						objBuilder.Append "	AND pyit.incom_year = '"&rever_year&"' "
						objBuilder.Append "WHERE (ISNULL(emtt.emp_end_date) OR emtt.emp_end_date = '1900-01-01' OR emtt.emp_end_date >= '"&st_in_date&"') "
						objBuilder.Append "	AND emtt.emp_in_date < '"&st_in_date&"' "
						objBuilder.Append "	AND emtt.emp_pay_id <> '5' AND emtt.emp_no < '900000' "

						If view_condi <> "전체" Then
							objBuilder.Append "	AND emtt.emp_company = '"&view_condi&"' "
						End If

						objBuilder.Append "ORDER BY emtt.emp_in_date, emtt.emp_no "
						objBuilder.Append "LIMIT "&stpage&","&pgsize&";"

						Set rsPay = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsPay.EOF Then
							arrPay = rsPay.getRows()
						End If
						rsPay.Close() : Set rsPay = Nothing

						If IsArray(arrPay) Then
							For i = LBound(arrPay) To UBound(arrPay, 2)
								emp_no = arrPay(0, i)
								emp_name = arrPay(1, i)
								emp_first_date = arrPay(2, i)
								emp_in_date = arrPay(3, i)
								emp_type = arrPay(4, i)
								emp_grade = arrPay(5, i)
								emp_position = arrPay(6, i)
								emp_company = arrPay(7, i)
								emp_bonbu = arrPay(8, i)
								emp_saupbu = arrPay(9, i)
								emp_team = arrPay(10, i)
								emp_org_code = arrPay(11, i)
								emp_org_name = arrPay(12, i)
								emp_reside_place = arrPay(13, i)
								emp_reside_company = arrPay(14, i)

								pmg_emp_no = arrPay(15, i)
								pmg_base_pay = CLng(f_toString(arrPay(16, i), 0))
								pmg_give_total = CLng(f_toString(arrPay(17, i), 0))

								de_deduct_total = CLng(f_toString(arrPay(18, i), 0))

								incom_base_pay = CLng(f_toString(arrPay(19, i), 0))
								incom_meals_pay = CLng(f_toString(arrPay(20, i), 0))
								incom_overtime_pay = CLng(f_toString(arrPay(21, i), 0))
								incom_month_amount = CLng(f_toString(arrPay(22, i), 0))
								incom_family_cnt = CLng(f_toString(arrPay(23, i), 0))
								incom_nps_amount = CLng(f_toString(arrPay(24, i), 0))
								incom_nhis_amount = CLng(f_toString(arrPay(25, i), 0))
								incom_nps = CLng(f_toString(arrPay(26, i), 0))
								incom_nhis = CLng(f_toString(arrPay(27, i), 0))
								incom_wife_yn = CLng(f_toString(arrPay(28, i), 0))
								incom_age20 = CLng(f_toString(arrPay(29, i), 0))
								incom_age60 = CLng(f_toString(arrPay(30, i), 0))
								incom_old = CLng(f_toString(arrPay(31, i), 0))

								incom_go_yn = f_toString(arrPay(32, i), "여")
								incom_long_yn = f_toString(arrPay(33, i), "여")

								'귀속월 급여 지급 여부
								If f_toString(pmg_emp_no, "") <> "" Then
									pmg_curr_pay = pmg_give_total - de_deduct_total
								Else
									pmg_curr_pay = 0

									If incom_base_pay <> 0 Then
										pmg_base_pay = incom_base_pay
										pmg_meals_pay = incom_meals_pay
										pmg_overtime_pay = incom_overtime_pay

										If incom_month_amount = 0 then
											incom_month_amount = incom_base_pay + incom_overtime_pay
										End If
									End If

									pmg_tax_yes = pmg_base_pay + pmg_overtime_pay
									pmg_tax_no = pmg_meals_pay
									pmg_give_total = pmg_tax_yes + pmg_tax_no

									'if incom_family_cnt = 0 then
									incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + incom_old + 1 '부양가족은 본인포함으로
									'end if

									'근로소득 간이세액 산출
									inc_st_amt = 0
									inc_incom = 0

									objBuilder.Append "SELECT inc_st_amt, inc_incom1, inc_incom2, inc_incom3, inc_incom4, inc_incom5 "
									objBuilder.Append "	inc_incom6, inc_incom7, inc_incom8, inc_incom9, inc_incom10, inc_incom11 "
									objBUilder.Append "FROM pay_income_amount "
									objBuilder.Append "WHERE ('"&incom_month_amount&"' BETWEEN inc_from_amt AND inc_to_amt) "
									objBuilder.Append "	AND inc_yyyy = '"&rever_year&"';"

									Set rs_sod = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If Not rs_sod.EOF Then
										inc_st_amt = CInt(f_toString(rs_sod("inc_st_amt"), 0))

										If incom_family_cnt = 1 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom1"), 0))
										End If

										If incom_family_cnt = 2 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom2"), 0))
										End If

										If incom_family_cnt = 3 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom3"), 0))
										End If

										If incom_family_cnt = 4 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom4"), 0))
										End If

										If incom_family_cnt = 5 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom5"), 0))
										End If

										If incom_family_cnt = 6 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom6"), 0))
										End If

										If incom_family_cnt = 7 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom7"), 0))
										End If

										If incom_family_cnt = 8 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom8"), 0))
										End If

										If incom_family_cnt = 9 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom9"), 0))
										End If

										If incom_family_cnt = 10 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom10"), 0))
										End If

										If incom_family_cnt = 11 Then
											inc_incom = CInt(f_toString(rs_sod("inc_incom11"), 0))
										End If
									End If
									rs_sod.Close()

									'소득세
									de_income_tax = CLng(inc_incom)

									'국민연금 계산
									'nps_amt = incom_nps_amount * (nps_emp / 100)
									'nps_amt = int(nps_amt)
									'de_nps_amt = (int(nps_amt / 10)) * 10
									de_nps_amt = incom_nps

									'건강보험 계산
									'nhis_amt = incom_nhis_amount * (nhis_emp / 100)
									'nhis_amt = int(nhis_amt)
									'de_nhis_amt = (int(nhis_amt / 10)) * 10
									de_nhis_amt = incom_nhis

									'장기요양보험 계산
									If incom_long_yn = "여" Then
										long_amt = de_nhis_amt * (long_hap / 100)
										long_amt = CInt(long_amt)
										'long_amt = long_amt / 2
										de_longcare_amt = (CInt(long_amt / 10)) * 10
									Else
										de_longcare_amt = 0
									End If

									'고용보험 계산 : 비과세 포함한 금액으로 계산
									If incom_go_yn = "여" Then
										'epi_amt = inc_st_amt * (epi_emp / 100)

										'pmg_give_tot 지정된 값 없음->pmg_give_total로 변경[허정호_20220331]
										'epi_amt = pmg_give_tot * (epi_emp / 100)
										epi_amt = pmg_give_total * (epi_emp / 100)

										epi_amt = CInt(epi_amt)
										de_epi_amt = (CInt(epi_amt / 10)) * 10
									Else
										de_epi_amt = 0
									End If

									'지방소득세
									we_tax = inc_incom * (10 / 100)
									we_tax = CInt(we_tax)
									de_wetax = (CInt(we_tax / 10)) * 10

									de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax
									pmg_curr_pay = pmg_give_total - de_deduct_total
								End If
								'귀속월 급여 지급 여부_END
						%>
							<tr>
								<td class="first"><%=emp_no%>&nbsp;</td>
                                <td><%=emp_name%>&nbsp;</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=emp_first_date%>&nbsp;</td>
                                <td><%=emp_in_date%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(pmg_base_pay, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(pmg_give_total, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(de_deduct_total, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(pmg_curr_pay, 0)%>&nbsp;</td>
                                <td class="left">
								<%Call EmpOrgCodeSelect(emp_org_code)%>
								</td>
							</tr>
						<%
							Next
						End If
						Set rs_sod = Nothing
						%>
						</tbody>
					</table>
				</div>
	          	<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                   	<td width="25%">
					<div class="btnleft">
                    <a href="/pay/insa_excel_pay_month_batch.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_yymm_to=<%=pmg_yymm_to%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <%
					'Page Navi
					Call Page_Navi_Ver2(page, be_pg, pg_url, total_record, pgsize)
					DBConn.Close() : Set DBConn = Nothing
					%>
                    </td>

                    <td width="25%">
					<!--미사용 기능으로 숨김 처리[허정호_20220404]
					<div class="btnRight">
						<a href="#" onClick="pay_month_transe('pmg_yymm','view_condi','pmg_yymm_to','to_date');return false;" class="btnType04">급여이월자료 등록</a>
					</div>-->
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="pmg_yymm1" value="<%=pmg_yymm%>"/>
                  <input type="hidden" name="pmg_yymm_to1" value="<%=pmg_yymm_to%>"/>
                  <input type="hidden" name="view_condi1" value="<%=view_condi%>"/>
                  <input type="hidden" name="to_date1" value="<%=to_date%>"/>
			</form>
            </form>
		</div>
	</div>
	</body>
</html>

