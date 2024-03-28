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
Dim month_tab(24,2)

Dim uploadForm, pay_company, pay_month, give_date, file_type
Dim ck_sw, curr_dd, from_date, cal_month, view_month, i, j
Dim cal_year, objFile, rowcount, title_line, etc_code, rs_etc
Dim emp_payend_date, emp_payend_yn, be_pg, emp_payend, rs_org
Dim att_file, filenm, path, filename, fileType, file_name, save_path
Dim cn, rs, tot_dz, xgr, fldcount, rsDz, rs_bnk, rs_give, dz_sw, dz_id
Dim name_sw, bank_sw, pmg_base_pay, pmg_meals_pay, pmg_research_pay, pmg_postage_pay
Dim pmg_re_pay, pmg_overtime_pay, pmg_car_pay, pmg_position_pay, pmg_job_pay
Dim pmg_job_support, pmg_jisa_pay, pmg_long_pay, pmg_disabled_pay, pmg_family_pay
Dim pmg_school_pay, pmg_qual_pay, pmg_other_pay1, pmg_other_pay2, pmg_other_pay3
Dim pmg_tax_yes, pmg_tax_no, pmg_tax_reduced, pmg_custom_pay, pmg_give_total, de_nps_amt
Dim de_nhis_amt, de_epi_amt, de_longcare_amt, de_income_tax, de_wetax, de_year_incom_tax
Dim de_year_wetax, de_other_amt1, de_special_tax, de_saving_amt, de_sawo_amt, de_johab_amt
Dim de_school_amt, de_nhis_bla_amt, de_long_bla_amt, de_hyubjo_amt, de_year_incom_tax2
Dim de_year_wetax2, de_deduct_total, reg_sw, reg_flag, bgcolor0, bgcolor1, bgcolor2, bgcolor3
Dim emp_name, pmg_date, fld_cnt_err, field_cnt

Set uploadForm = Server.CreateObject("ABCUpload4.XForm")

uploadForm.AbsolutePath = True
uploadForm.Overwrite = True
uploadForm.MaxUploadSize = 1024*1024*50

pay_company = uploadForm("pay_company")
pay_month = uploadForm("pay_month")
give_date = uploadForm("give_date")
file_type = uploadForm("file_type")

pmg_date = f_toString(uploadForm("pmg_date"), Mid(CStr(Now()),1,10))

'if ck_sw = "y" then
'	pay_company = request("pay_company")
'	pay_month = request("pay_month")
'end if

be_pg = "/pay/insa_pay_month_up.asp"

If pay_company = "" Then
	ck_sw = "y"
Else
	ck_sw = "n"
End If

If pay_company = "" Then
    pay_company = "케이원"
    curr_dd = CStr(DatePart("d",Now()))
    give_date = Mid(CStr(Now()),1,10)
    from_date = Mid(CStr(Now()-curr_dd+1),1,10)
    pay_month = Mid(CStr(from_date),1,4)&Mid(CStr(from_date),6,2)
End If

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = Mid(CStr(Now()),1,4)&Mid(CStr(Now()),6,2)
month_tab(24,1) = cal_month
view_month = Mid(cal_month,1,4)&"년 "&Mid(cal_month,5,2)&"월"
month_tab(24,2) = view_month

For i = 1 To 23
	cal_month = CStr(CLng(cal_month) - 1)

	If Mid(cal_month,5) = "00" Then
		cal_year = CStr(CInt(Mid(cal_month,1,4)) - 1)
		cal_month = cal_year&"12"
	End If

	view_month = Mid(cal_month,1,4)&"년 "&Mid(cal_month,5,2)&"월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
Next

If ck_sw = "n" Then
	Set filenm = uploadForm("att_file")(1)

	path = Server.MapPath ("/pay_file")
	filename = filenm.safeFileName
	fileType = Mid(filename,InStrRev(filename,".")+1)
	file_name = pay_company&"_"&pay_month&"_급여"&give_date

'		save_path = path & "\" & filename
	save_path = path&"\"&file_name&"."&fileType

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
		rs.Open "select * from [2:10000]",cn,"0"

		rowcount = -1
		xgr = rs.getRows
		rowcount = UBound(xgr,2)
		fldcount = rs.fields.count
		tot_cnt = rowcount + 1

		'Response.write fldcount

		'회사별 항목 개수 체크
		Select Case pay_company
			Case "케이원"
				field_cnt = 40
			Case "케이네트웍스"
				field_cnt = 37
			Case "케이시스템"
				field_cnt = 31
		End Select

		If fldcount <> field_cnt Then
			fld_cnt_err = "Y"
		End If
	Else
		objFile = "none"
		rowcount = -1
		file_type = "N"
	End If
Else
	objFile = "none"
	rowcount=-1
End If

title_line = "급여 자료 업로드"

etc_code = "9999"

objBuilder.Append "SELECT emp_payend_date, emp_payend_yn FROM emp_etc_code WHERE emp_etc_code = '"&etc_code&"';"

Set rs_etc = DBConn.Execute(objBuilder.toString())
objBuilder.Clear()

emp_payend_date = rs_etc("emp_payend_date")
emp_payend_yn = rs_etc("emp_payend_yn")

rs_etc.Close() : Set rs_etc = Nothing

If pay_month > emp_payend_date Then
	emp_payend = "N"
Else
	emp_payend = "Y"
End If
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
		<!--<script type="text/javascript" src="/java/js_window.js"></script>-->
		<script type="text/javascript">
            // 검색 버튼 클릭!!
			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}
			//지급일자
			$(function(){
				$( "#datepicker" ).datepicker();
				$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
				$( "#datepicker" ).datepicker("setDate", "<%=pmg_date%>" );
			});

			function chkfrm(){
				if(document.frm.pay_company.value == ""){
					alert ("회사를 선택하세요");
					return false;
				}

				if(document.frm.pay_month.value == ""){
					alert ("귀속년월을 선택하세요");
					return false;
				}

				if(document.frm.pmg_date.value == ""){
					alert ("지급일자를 선택하세요");
					return false;
				}

				if(document.frm.att_file.value == ""){
					alert ("업로드 엑셀 파일을 선택하세요");
					return false;
				}
				return true;
            }

            // 급여 upload 버튼 클릭!!
			function frm1check(){
				if(chkfrm1()){
					document.frm1.submit();
				}
			}

            function chkfrm1(){
				if(confirm('DB에 업로드를 하시겠습니까?') == true){
					return true;
				}
				return false;
			}

			//급여 내역 삭제
            function pay_month_updel(val, val2){
				if(!confirm("급여 Upload자료를 삭제하시겠습니까?")) return;

                var frm = document.frm;

                document.frm.pay_month1.value = document.getElementById(val).value;
                document.frm.pay_company1.value = document.getElementById(val2).value;

                document.frm.action = "/pay/insa_pay_month_up_del.asp";
                document.frm.submit();
            }
		</script>
</head>
<body>
	<div id="wrap">
	<!--#include virtual = "/include/insa_pay_header.asp" -->
	<!--#include virtual = "/include/insa_pay_menu.asp" -->
		<div id="container">
			<h3 class="insa"><%=title_line%></h3><br/>
				<form action="<%=be_pg%>" method="post" name="frm" enctype="multipart/form-data">
					<fieldset class="srch">
						<legend>조회영역</legend>
						<dl>
							<dt>업로드내용</dt>
							<dd>
								<p>
									<label>
										<strong>회사: </strong>
										<%
                                        ' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
                                        'objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC;"
										objBuilder.Append "SELECT org_name FROM emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') "
										objBuilder.Append "	AND org_level = '회사' AND org_code <> '6272' "
										objBuilder.Append "ORDER BY FIELD(org_name, "&OrderByOrgName&") ASC;"

                                        Set rs_org = DBConn.Execute(objBuilder.ToString())
										objBuilder.Clear()
                                        %>
                                        <select name="pay_company" id="pay_company" type="text" style="width:110px;">
                                            <option value="">선택</option>
                                            <%
                                            Do Until rs_org.EOF
                                                %>
                                                <option value='<%=rs_org("org_name")%>' <%If pay_company = rs_org("org_name") Then %>selected<%End If %>><%=rs_org("org_name")%></option>
                                                <%
                                                rs_org.MoveNext()
                                            Loop
                                            rs_org.Close() : Set rs_org = Nothing
                                            %>
                                        </select>
                                    </label>
                                    <label>
                                        <strong>귀속년월: </strong>
                                        <select name="pay_month" id="pay_month" value="<%=pay_month%>" style="width:90px;">
                                            <%For i = 24 To 1 Step -1	%>
                                            <option value="<%=month_tab(i,1)%>" <%If pay_month = month_tab(i,1) Then %>selected<%End If %>><%=month_tab(i,2)%></option>
                                            <%Next	%>
                                        </select>
                                    </label>

									<label>
                                        <strong>지급일자: </strong>
                                        <input type="text" name="pmg_date" id="datepicker" value="<%=pmg_date%>" style="width:90px;"/>
                                    </label>

                                    <!--<br>-->
                                    <label>
                                        <strong>업로드파일: </strong>
                                        <input name="att_file" type="file" id="att_file" size="100" value="<%=att_file%>" style="text-align:left;"/>
                                    </label>

                                    <input name="file_type" type="hidden" id="file_type" value="<%=file_type%>"/>
                                    <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"/></a>
                                </p>
                            </dd>
						</dl>
					</fieldset>
					<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
							<colgroup>
								<col width="3%" >
								<col width="3%" >
								<col width="4%" >
								<col width="4%" >
								<col width="7%" >
								<col width="7%" >
								<col width="4%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="6%" >
								<col width="5%" >
								<col width="5%" >
								<col width="6%" >
                                <col width="6%" >
                                <col width="5%" >
                                <col width="*" >
                                <col width="6%" >
                                <col width="8%" >
							</colgroup>
							<thead>
								<tr>
									<th class="first" scope="col">건수</th>
									<th scope="col">등록</th>
									<th scope="col">사번</th>
									<th scope="col">급여ID</th>
									<th scope="col">성명</th>
									<th scope="col">기본급</th>
									<th scope="col">식대</th>
									<th scope="col">연구<br>수당</th>
									<th scope="col">통신비</th>
									<th scope="col">소급</th>
									<th scope="col">연장</th>
									<th scope="col">주차<br>지원</th>
									<th scope="col">직책</th>
									<th scope="col">고객<br>관리</th>
									<th scope="col">직무<br>보조</th>
                                    <th scope="col">업무<br>장려</th>
                                    <th scope="col">본지사<br>근무</th>
                                    <th scope="col">근속</th>
                                    <th scope="col">장애인</th>
                                    <th scope="col">지급<br>액계</th>
								</tr>
							</thead>
							<tbody>
							<%
							Dim tot_emp, tot_name, tot_bank, tot_err, tot_base_pay, tot_meals_pay, tot_research_pay
							Dim tot_postage_pay, tot_re_pay, tot_overtime_pay, tot_car_pay, tot_position_pay
							Dim tot_custom_pay, tot_job_pay, tot_job_support, tot_jisa_pay, tot_long_pay, tot_disabled_pay
							Dim tot_family_pay, tot_school_pay, tot_qual_pay, tot_other_pay1, tot_other_pay2, tot_other_pay3
							Dim tot_tax_yes, tot_tax_no, tot_tax_reduced, tot_give_total, error_line, bank_flag
							Dim emp_sw

							tot_emp = 0'직원미등록건수
							tot_name = 0'직원명미등록건수
							tot_bank = 0'은행미등록건수
							tot_err = 0'전체에러건수

							tot_base_pay = 0
							tot_meals_pay = 0
							tot_research_pay = 0
							tot_postage_pay = 0
							tot_re_pay = 0
							tot_overtime_pay = 0
							tot_car_pay = 0
							tot_position_pay = 0
							tot_custom_pay = 0
							tot_job_pay = 0
							tot_job_support = 0
							tot_jisa_pay = 0
							tot_long_pay = 0
							tot_disabled_pay = 0
							tot_family_pay = 0
							tot_school_pay = 0
							tot_qual_pay = 0
							tot_other_pay1 = 0
							tot_other_pay2 = 0
							tot_other_pay3 = 0
							tot_tax_yes = 0
							tot_tax_no = 0
							tot_tax_reduced = 0
							tot_give_total = 0

							tot_dz = 0'급여ID미등록건수

							If rowcount > -1 And fld_cnt_err <> "Y" Then
								For i=0 To rowcount
									If f_toString(xgr(0,i), "") = "" Then
										Exit For
									End If

									error_line = 0'에러표기라인 건수

									dz_sw = "Y"'급여ID 에러 구분
									emp_sw = "Y"'사번 에러 구분
									name_sw = "Y"'성명 에러 구분
									bank_flag = "Y"'은행 에러 구분

									'급여ID(사원코드) 체크

									dz_id = xgr(0, i)

									objBuilder.Append "SELECT dpit.dz_id, dpit.emp_company, dpit.emp_no, emtt.emp_name, "
									objBuilder.Append "	pbat.emp_no AS 'bank_emp_no' "
									objBuilder.Append "FROM dz_pay_info AS dpit "
									objBuilder.Append "LEFT OUTER JOIN emp_master AS emtt ON dpit.emp_no = emtt.emp_no "
									objBuilder.Append "	AND (ISNULL(emp_end_date) OR emp_end_date = '1900-01-01') "
									objBuilder.Append "LEFT OUTER JOIN pay_bank_account AS pbat ON dpit.emp_no = pbat.emp_no "
									objBuilder.Append "WHERE dpit.dz_id='"&dz_id&"' AND dpit.emp_company='"&pay_company&"';"

									Set rsDz = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rsDz.EOF Or rsDz.BOF Then
										tot_err = tot_err + 1
										error_line = error_line + 1

										tot_dz = tot_dz + 1
										tot_emp = tot_emp + 1
										tot_bank = tot_bank + 1

										dz_sw = "N"
										emp_sw = "N"
										name_sw = "N"
										bank_sw = "N"
										bank_flag = "N"

										emp_no = "Error"
										emp_name = "Error"
									Else
										emp_no = rsDz("emp_no")
										emp_name = rsDz("emp_name")
									End If
									rsDz.Close()

									'급여 등록 여부 확인
									objBuilder.Append "SELECT pmg_emp_no FROM pay_month_give "
									objBuilder.Append "WHERE pmg_yymm = '"&pay_month&"' AND pmg_id = '1' AND pmg_emp_no = '"&emp_no&"';"

									Set rs_give = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If rs_give.EOF Or rs_give.BOF Then
										reg_flag = "No"
									Else
										reg_flag = "Yes"
									End If
									rs_give.Close()

									' 지급항목
									pmg_base_pay = toString(xgr(4,i),0)	'기본급
									pmg_meals_pay = toString(xgr(5,i),0)	'식대
									pmg_postage_pay = toString(xgr(6,i),0)	'통신비(PL수당)
									pmg_re_pay = toString(xgr(7,i),0)	'소급급여
									pmg_overtime_pay = toString(xgr(8,i),0)	'연장근로수당

									'고객관리수당, 원본 급여 엑셀 파일 항목에 없음(202106_오지훈부장 데이타 1개 존재함)
									'pmg_custom_pay	  = toString(xgr(20,i),"0")	'고객관리수당
									pmg_custom_pay = 0	'고객관리수당

									Select Case pay_company
										Case "케이원"
											'지급항목
											pmg_car_pay = toString(xgr(9,i),0)	'주차지원금
											pmg_job_pay = toString(xgr(10,i),0)	'직무보조비(자격수당)
											pmg_job_support = toString(xgr(11,i) + xgr(15, i),0)	'업무장려비(업무장려비 + 시간외수당)
											pmg_jisa_pay = toString(xgr(12,i),0)	'본지사근무비
											pmg_disabled_pay = toString(xgr(13,i),0)	'장애인수당
											pmg_research_pay = toString(xgr(14,i),0)	'연구(연구수당)
											pmg_position_pay = toString(xgr(16,i),0)	'직책수당
											pmg_long_pay = toString(xgr(17,i),0)	'근속수당(PM수당)

											'공제항목
											de_nps_amt = toString(xgr(19,i),0)'국민연금
											de_nhis_amt = toString(xgr(20,i),0)'건강보험
											de_epi_amt = toString(xgr(21,i),0)'고용보험
											de_longcare_amt = toString(xgr(22,i),0)'장기요양보험료
											de_income_tax = toString(xgr(23,i),0)'소득세
											de_wetax = toString(xgr(24,i),0)'지방소득세
											de_year_incom_tax = toString(xgr(25,i),0)'연말정산소득세
											de_year_wetax = toString(xgr(26,i),0)'연말정산지방세
											de_other_amt1 = toString(xgr(30,i),0)'기타공제
											de_sawo_amt = toString(xgr(31,i),0)'사우회회비
											de_school_amt = toString(xgr(28,i),0)'학자금상환
											de_nhis_bla_amt = toString(xgr(33,i),0)'건강보험료정산
											de_long_bla_amt	= toString(xgr(34,i),0)'장기요양보험료정산
											de_hyubjo_amt = toString(xgr(32,i),0)'협조비

											de_year_incom_tax2 = toString(xgr(38,i),0)'연말재정산소득세
											de_year_wetax2 = toString(xgr(39,i),0)'연말재정산지방세
										Case "케이네트웍스"
											'지급항목
											pmg_car_pay = toString(xgr(9,i),0)	'주차지원금
											pmg_job_pay = toString(xgr(11,i),0)	'직무보조비(자격수당)
											pmg_job_support = toString(xgr(12,i) + xgr(14, i),0)	'업무장려비(업무장려비 + 시간외수당)
											pmg_jisa_pay = toString(xgr(13,i),0)	'본지사근무비
											pmg_disabled_pay = 0	'장애인수당
											pmg_research_pay = 0	'연구(연구수당)
											pmg_position_pay = toString(xgr(10,i),0)	'직책수당
											pmg_long_pay = toString(xgr(15,i),0)	'근속수당(PM수당)

											'공제항목
											de_nps_amt = toString(xgr(17,i),0)'국민연금
											de_nhis_amt = toString(xgr(18,i),0)'건강보험
											de_epi_amt = toString(xgr(19,i),0)'고용보험
											de_longcare_amt = toString(xgr(20,i),0)'장기요양보험료
											de_income_tax = toString(xgr(21,i),0)'소득세
											de_wetax = toString(xgr(22,i),0)'지방소득세
											de_year_incom_tax = toString(xgr(23,i),0)'연말정산소득세
											de_year_wetax = toString(xgr(24,i),0)'연말정산지방세

											de_other_amt1 = toString(xgr(27,i),0)'기타공제
											de_sawo_amt = toString(xgr(28,i),0)'사우회회비
											de_school_amt = toString(xgr(26,i),0)'학자금상환
											de_nhis_bla_amt = toString(xgr(29,i),0)'건강보험료정산
											de_long_bla_amt	= toString(xgr(30,i),0)'장기요양보험료정산
											de_hyubjo_amt = 0'협조비

											de_year_incom_tax2 = toString(xgr(35,i),0)'연말재정산소득세
											de_year_wetax2 = toString(xgr(36,i),0)'연말재정산지방세
										Case "케이시스템"
											'지급항목
											pmg_car_pay = 0
											pmg_job_pay = toString(xgr(11,i),0)	'직무보조비(자격수당)
											pmg_job_support = toString(xgr(9,i),0)	'업무장려비
											pmg_jisa_pay = 0	'본지사근무비
											pmg_disabled_pay = 0	'장애인수당
											pmg_research_pay = 0	'연구(연구수당)
											pmg_position_pay = 0	'직책수당
											pmg_long_pay = toString(xgr(10,i),0)	'근속수당(PM수당)

											'공제항목
											de_nps_amt = toString(xgr(13,i),0)'국민연금
											de_nhis_amt = toString(xgr(14,i),0)'건강보험
											de_epi_amt = toString(xgr(15,i),0)'고용보험
											de_longcare_amt = toString(xgr(16,i),0)'장기요양보험료
											de_income_tax = toString(xgr(17,i),0)'소득세
											de_wetax = toString(xgr(18,i),0)'지방소득세
											de_year_incom_tax = toString(xgr(19,i),0)'연말정산소득세
											de_year_wetax = toString(xgr(20,i),0)'연말정산지방세
											de_other_amt1 = toString(xgr(25,i),0)'기타공제
											de_sawo_amt = toString(xgr(26,i),0)'사우회회비
											de_school_amt = toString(xgr(22,i),0)'학자금상환
											de_nhis_bla_amt = toString(xgr(23,i),0)'건강보험료정산
											de_long_bla_amt	= toString(xgr(24,i),0)'장기요양보험료정산
											de_hyubjo_amt = 0'협조비

											de_year_incom_tax2 = toString(xgr(29,i),0)'연말재정산소득세
											de_year_wetax2 = toString(xgr(30,i),0)'연말재정산지방세
									End Select

									pmg_family_pay = 0
									pmg_school_pay = 0
									pmg_qual_pay = 0
									pmg_other_pay1 = 0
									pmg_other_pay2 = 0
									pmg_other_pay3 = 0
									pmg_tax_yes = 0
									pmg_tax_no = 0
									pmg_tax_reduced = 0

									de_special_tax = 0
									de_saving_amt = 0
									de_johab_amt = 0

									pmg_give_total = pmg_base_pay + pmg_meals_pay + pmg_research_pay + pmg_postage_pay + pmg_re_pay
									pmg_give_total = pmg_give_total + pmg_overtime_pay + pmg_car_pay + pmg_position_pay + pmg_custom_pay
									pmg_give_total = pmg_give_total + pmg_job_pay + pmg_job_support + pmg_jisa_pay + pmg_long_pay + pmg_disabled_pay
									'pmg_give_total = xgr(25,i)

									de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax
									de_deduct_total = de_deduct_total + de_wetax + de_year_incom_tax + de_year_wetax + de_year_incom_tax2
									de_deduct_total = de_deduct_total + de_year_wetax2 + de_other_amt1 + de_sawo_amt + de_school_amt
									de_deduct_total = de_deduct_total + de_nhis_bla_amt + de_long_bla_amt + de_hyubjo_amt
									'de_deduct_total = xgr(38,i)

									tot_base_pay = tot_base_pay + pmg_base_pay
									tot_meals_pay = tot_meals_pay + pmg_meals_pay
									tot_research_pay = tot_research_pay + pmg_research_pay
									tot_postage_pay = tot_postage_pay + pmg_postage_pay
									tot_re_pay = tot_re_pay + pmg_re_pay
									tot_overtime_pay = tot_overtime_pay + pmg_overtime_pay
									tot_car_pay = tot_car_pay + pmg_car_pay
									tot_position_pay = tot_position_pay + pmg_position_pay
									tot_custom_pay = tot_custom_pay + pmg_custom_pay
									tot_job_pay = tot_job_pay + pmg_job_pay
									tot_job_support = tot_job_support + pmg_job_support
									tot_jisa_pay = tot_jisa_pay + pmg_jisa_pay
									tot_long_pay = tot_long_pay + pmg_long_pay
									tot_disabled_pay = tot_disabled_pay + pmg_disabled_pay
									tot_family_pay = tot_family_pay + pmg_family_pay
									tot_school_pay = tot_school_pay + pmg_school_pay
									tot_qual_pay = tot_qual_pay + pmg_qual_pay
									tot_other_pay1 = tot_other_pay1 + pmg_other_pay1
									tot_other_pay2 = tot_other_pay2 + pmg_other_pay2
									tot_other_pay3 = tot_other_pay3 + pmg_other_pay3
									tot_tax_yes = tot_tax_yes + pmg_tax_yes
									tot_tax_no = tot_tax_no + pmg_tax_no
									tot_tax_reduced = tot_tax_reduced + pmg_tax_reduced
									tot_give_total = tot_give_total + pmg_give_total
							%>
								<tr <%If error_line > 0 Then%>style="background-color:#FFCCFF;"<%End If%>>
									<td class="first"><%=i+1%></td>
									<td><%=reg_flag%></td>
									<td <%If emp_sw = "N" Then%>style="color:red;"<%End If%>><%=emp_no%></td>
									<td><%=dz_id%></td>
									<td <%If name_sw = "N" Then%>style="color:red;"<%End If%>><%=emp_name%></td>
									<td><%=FormatNumber(pmg_base_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_meals_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_research_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_postage_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_re_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_overtime_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_car_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_position_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_custom_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_job_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_job_support,0)%></td>
									<td class="right"><%=FormatNumber(pmg_jisa_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_long_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_disabled_pay,0)%></td>
									<td class="right"><%=FormatNumber(pmg_give_total,0)%></td>
								</tr>
							<%
								Next

								Set rsDz = Nothing
								Set rs_give = Nothing
								rs.Close() : Set rs = Nothing
								cn.Close() :  Set cn = Nothing
							Else
								Response.Write "<tr><td colspan='20' style='height:30px;'>"

								If fld_cnt_err = "Y" Then
									Response.Write "업로드 파일의 항목 개수가 일치하지 않습니다.(필수 항목 : "&field_cnt&")"
								Else
									Response.Write "조회된 내역이 없습니다."
								End If

								Response.Write "</td></tr>"
							End If
							DBConn.Close() : Set DBConn = Nothing
							%>
								<tr>
									<th class="first">오류</th>
									<th title="급여계좌미등록건수"><%=FormatNumber(tot_bank,0)%></th>
									<!--<th title="직원미등록건수"><%'=FormatNumber(tot_emp,0)%></th>
									<th title="급여ID미등록건수"><%'=FormatNumber(tot_dz,0)%></th>-->
									<th colspan="2" title="미등록건수"><%=FormatNumber(tot_err,0)%></th>

									<!--<th><%'=FormatNumber(tot_name,0)%></th>-->
									<th>합계</th>

									<th class="right"><%=FormatNumber(tot_base_pay,0)%></th>
									<th class="right"><%=FormatNumber(tot_meals_pay,0)%></th>
									<th class="right"><%=FormatNumber(tot_research_pay,0)%></th>
									<th class="right"><%=FormatNumber(tot_postage_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_re_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_overtime_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_car_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_position_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_custom_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_job_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_job_support,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_jisa_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_long_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_disabled_pay,0)%></th>
                                    <th class="right"><%=FormatNumber(tot_give_total,0)%></th>
								</tr>
							</tbody>
						</table>
					</div>
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  	<tr>
                        <td width="15%"><div class="btnCenter"></div></td>
                        <td>
                            <div class="btnRight"><a href="#" onClick="pay_month_updel('pay_month','pay_company');return false;" class="btnType04">급여 Upload 삭제</a></div>
                        </td>
                    </tr>
                    </table>

                    <input type="hidden" name="pay_company1" value="<%=pay_company%>"/>
                    <input type="hidden" name="pay_month1" value="<%=pay_month%>"/>
				</form>
				<%
				Dim tot_cnt

                If emp_payend = "N" Then
                    If tot_cnt <> 0 And tot_err = 0 Then
                    %>
                        <form action="/pay/insa_pay_month_up_ok.asp" method="post" name="frm1">
                            <br>
                            <div align="center">
                                <span class="btnType01"><input type="button" value="급여자료 Upload" onclick="javascript:frm1check();"/></span>
                            </div>
                            <input type="hidden" name="objFile" id="objFile" value="<%=objFile%>"/>
                            <input type="hidden" name="pmg_yymm" id="pmg_yymm" value="<%=pay_month%>"/>
                            <input type="hidden" name="pmg_company" id="pmg_company" value="<%=pay_company%>"/>
							<!--<input type="hidden" name="pmg_date" id="pmg_date" value="<%'=give_date%>"/>-->
							<input type="hidden" name="pmg_date" id="pmg_date" value="<%=pmg_date%>"/>
                            <br/>
                        </form>
				    <%
					End If
			   	End If
			  %>
			</div>
		</div>
	</body>
</html>