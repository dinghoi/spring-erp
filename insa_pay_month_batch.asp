<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

be_pg = "insa_pay_month_batch.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	pmg_yymm=Request.form("pmg_yymm")
    pmg_yymm_to=Request.form("pmg_yymm_to")
	to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
	pmg_yymm_to=request("pmg_yymm_to")
	to_date=request("to_date")
end if

if view_condi = "" then
	view_condi = "전체"
	ck_sw = "n"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm_to = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))

'매월 말일 구하기
   datYear = mid(cstr(pmg_yymm_to),1,4)
   datMonth = mid(cstr(pmg_yymm_to),5,2)
   If datMonth=4 or datMonth=6 or datMonth=9 or datMonth=11 Then  '4월 6월 9월 11월이면 월말값은 30일
             datLastDay=30
      ElseIf datMonth=2 and not (datYear mod 4) = 0 Then  '2월이고  년도를 4로 나눈 값이 0이 아니면 28일
                    datLastDay=28
             ElseIf datMonth=2 and (datYear mod 4) = 0 Then '윤달 계산
                        if (datYear mod 100) = 0 Then
                              if (datYear mod 400) = 0 Then
                                      datLastDay=29
                                  else
                                      datLastDay=28
                              End If
                          else
                              datLastDay=29
                        End If
                    else
                        datLastDay=31
   End If
   exec_LastDay = datLastDay
'   to_date = mid(cstr(pmg_yymm_to),1,4) + "-" + mid(cstr(pmg_yymm_to),5,2) + "-" + cstr(exec_LastDay)

   to_date = ""
end if

'당월 입사/퇴사일이 15일 이전이면 당월 급여대상임
st_es_date = mid(cstr(pmg_yymm_to),1,4) + "-" + mid(cstr(pmg_yymm_to),5,2) + "-" + "01"
st_in_date = mid(cstr(pmg_yymm_to),1,4) + "-" + mid(cstr(pmg_yymm_to),5,2) + "-" + "16"
rever_year = mid(cstr(pmg_yymm_to),1,4) '귀속년도


' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

pgsize = 10 ' 화면 한 페이지
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_ins = Server.CreateObject("ADODB.Recordset")
Set Rs_sod = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'고용보험(실업) 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5503' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	epi_emp = formatnumber(rs_ins("emp_rate"),3)
		epi_com = formatnumber(rs_ins("com_rate"),3)
   else
		epi_emp = 0
		epi_com = 0
end if
rs_ins.close()

'장기요양보험 요율
Sql = "SELECT * FROM pay_insurance where insu_yyyy = '"&rever_year&"' and insu_id = '5504' and insu_class = '01'"
Set rs_ins = DbConn.Execute(SQL)
if not rs_ins.eof then
    	long_hap = formatnumber(rs_ins("hap_rate"),3)
   else
		long_hap = 0
end if
rs_ins.close()


if view_condi = "전체" then
		   Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC"
       else
           Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"') and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC"
end if

Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_condi = "전체" then
		   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize
       else
           Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"')  and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize
end if

Rs.Open Sql, Dbconn, 1
'Response.write Sql

title_line = " 급여기초이월 처리 "

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
		</script>
		<script type="text/javascript">
		    $(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});
			$(function() {    $( "#to_date" ).datepicker();
												$( "#to_date" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#to_date" ).datepicker("setDate", "<%=to_date%>" );
			});

			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}

			function form_chk(){
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->

			function pay_month_transe(val, val2, val3, val4) {

			var tVal = document.getElementById(val).value;
			var tVal2 = document.getElementById(val2).value;
			var tVal3 = document.getElementById(val3).value;
			var tVal4 = document.getElementById(val4).value;

			if( tVal==null || tVal=="" ){ alert("이월대상년월을 선택해 주세요."); return; }
			if( tVal2==null || tVal2=="" ){ alert("회사를 선택해 주세요."); return; }
			if( tVal3==null || tVal3=="" ){ alert("귀속년월을 선택해 주세요."); return; }
			if( tVal4==null || tVal4=="" ){ alert("지급일을 선택해 주세요."); return; }

            if (!confirm("전월 급여를 이월처리 하시겠습니까 ?")) return;

            var frm = document.frm;
			document.frm.pmg_yymm1.value = tVal;
			document.frm.view_condi1.value = tVal2;
			document.frm.pmg_yymm_to1.value = tVal3;
			document.frm.to_date1.value = tVal4;

            document.frm.action = "insa_pay_month_transe_save.asp";
            document.frm.submit();
            }
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_month_batch.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE (ISNULL(org_end_date) OR org_end_date = '0000-00-00') AND org_level = '회사'  ORDER BY org_company ASC"
	                            rs_org.Open Sql, Dbconn, 1
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
                                    <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                			  <%
								do until rs_org.eof
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()
								loop
								rs_org.Close()
							  %>
            					</select>
                                </label>
                                <label>
								<strong>이월대상년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <label>
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm_to" id="pmg_yymm_to" type="text" value="<%=pmg_yymm_to%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm_to = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <label>
								<strong>지급일 : </strong>
                                	<input name="to_date" id="to_date" type="text" value="<%=to_date%>" style="width:70px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
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
 	 do until rs.eof
  		  emp_no = rs("emp_no")
          emp_company = rs("emp_company")
          emp_name = rs("emp_name")
          emp_in_date = rs("emp_in_date")
          pmg_emp_type = rs("emp_type")
          pmg_grade = rs("emp_grade")
          pmg_position = rs("emp_position")
          pmg_company = rs("emp_company")
          pmg_bonbu = rs("emp_bonbu")
          pmg_saupbu = rs("emp_saupbu")
          pmg_team = rs("emp_team")
          pmg_org_code = rs("emp_org_code")
          pmg_org_name = rs("emp_org_name")
          pmg_reside_place = rs("emp_reside_place")
          pmg_reside_company = rs("emp_reside_company")

		  sql = "select * from pay_month_give where (pmg_yymm = '"&pmg_yymm&"' ) and (pmg_id = '1') and (pmg_emp_no = '"&emp_no&"') and (pmg_company = '"&emp_company&"')"
          Set Rs_give = DbConn.Execute(SQL)
          if not Rs_give.eof then
                 pmg_base_pay = int(Rs_give("pmg_base_pay"))
                 pmg_give_total = int(Rs_give("pmg_give_total"))

				 Sql = "select * from pay_month_deduct where (de_yymm = '"&pmg_yymm&"' ) and (de_id = '1') and (de_emp_no = '"&emp_no&"') and (de_company = '"&pmg_company&"')"
                 Set Rs_dct = DbConn.Execute(SQL)
                 if not Rs_dct.eof then
                        de_deduct_total = int(Rs_dct("de_deduct_total"))
                    else
                        de_deduct_total = 0
                 end if
                 Rs_dct.close()
				 pmg_curr_pay = pmg_give_total - de_deduct_total
			 else
			     pmg_base_pay = 0
                 pmg_give_total = 0
				 de_deduct_total = 0
				 pmg_curr_pay = 0

			 '기본급/식대등 가져오기
                 incom_family_cnt = 0
                 Sql = "SELECT * FROM pay_year_income where incom_emp_no = '"&emp_no&"' and incom_year = '"&rever_year&"'"
                 Set Rs_year = DbConn.Execute(SQL)
                 if not Rs_year.eof then
    	               pmg_base_pay = Rs_year("incom_base_pay")
		               pmg_meals_pay = Rs_year("incom_meals_pay")
		               pmg_overtime_pay = Rs_year("incom_overtime_pay")
		               if Rs_year("incom_month_amount") = 0 or isnull(Rs_year("incom_month_amount")) then
		                      incom_month_amount = Rs_year("incom_base_pay") + Rs_year("incom_overtime_pay")
		                  else
		                      incom_month_amount = Rs_year("incom_month_amount")
		               end if
		               incom_family_cnt = Rs_year("incom_family_cnt")
		               incom_nps_amount = Rs_year("incom_nps_amount")
		               incom_nhis_amount = Rs_year("incom_nhis_amount")
		               incom_nps = Rs_year("incom_nps")
		               incom_nhis = Rs_year("incom_nhis")
		               incom_wife_yn = int(Rs_year("incom_wife_yn"))
	           	       incom_age20 = Rs_year("incom_age20")
		               incom_age60 = Rs_year("incom_age60")
		               incom_old = Rs_year("incom_old")
		               incom_go_yn = Rs_year("incom_go_yn")
		               incom_long_yn = Rs_year("incom_long_yn")
                    else
		               pmg_base_pay = 0
		               pmg_meals_pay = 0
		               pmg_overtime_pay = 0
		               incom_month_amount = 0
		               incom_family_cnt = 0
		               incom_nps_amount = 0
		               incom_nhis_amount = 0
		               incom_nps = 0
		               incom_nhis = 0
		               incom_go_yn = "여"
		               incom_long_yn = "여"
		               incom_wife_yn = 0
		               incom_age20 = 0
		               incom_age60 = 0
		               incom_old = 0
                 end if
                 Rs_year.close()

                 pmg_tax_yes = pmg_base_pay + pmg_overtime_pay
                 pmg_tax_no = pmg_meals_pay
                 pmg_give_total = pmg_tax_yes + pmg_tax_no

		         'if incom_family_cnt = 0 then
                       incom_family_cnt = incom_wife_yn + incom_age20 + incom_age60 + incom_old + 1 '부양가족은 본인포함으로
                 'end if

			     '근로소득 간이세액 산출
                 inc_st_amt = 0
                 inc_incom = 0

				 Sql = "SELECT * FROM pay_income_amount where ('"&incom_month_amount&"' BETWEEN inc_from_amt and inc_to_amt) and (inc_yyyy = '"&rever_year&"')"
                 Set Rs_sod = DbConn.Execute(SQL)
                 if not Rs_sod.eof then
   	                  inc_st_amt = int(Rs_sod("inc_st_amt"))
	                  if incom_family_cnt = 1 then
	                       inc_incom = Rs_sod("inc_incom1")
	                  end if
	                  if incom_family_cnt = 2 then
	                       inc_incom = Rs_sod("inc_incom2")
	                  end if
	                  if incom_family_cnt = 3 then
	                       inc_incom = Rs_sod("inc_incom3")
	                  end if
	                  if incom_family_cnt = 4 then
	                       inc_incom = Rs_sod("inc_incom4")
	                  end if
	                  if incom_family_cnt = 5 then
	                       inc_incom = Rs_sod("inc_incom5")
	                  end if
	                  if incom_family_cnt = 6 then
	                       inc_incom = Rs_sod("inc_incom6")
	                  end if
	                  if incom_family_cnt = 7 then
	                       inc_incom = Rs_sod("inc_incom7")
	                  end if
	                  if incom_family_cnt = 8 then
	                       inc_incom = Rs_sod("inc_incom8")
	                  end if
	                  if incom_family_cnt = 9 then
	                       inc_incom = Rs_sod("inc_incom9")
	                  end if
	                  if incom_family_cnt = 10 then
	                       inc_incom = Rs_sod("inc_incom10")
	                  end if
	                  if incom_family_cnt = 11 then
	                       inc_incom = Rs_sod("inc_incom11")
	                  end if
                 end if
                 Rs_sod.close()

			     '소득세
                 de_income_tax = int(inc_incom)

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
                 if incom_long_yn = "여" then
                        long_amt = de_nhis_amt * (long_hap / 100)
                        long_amt = Int(long_amt)
                        'long_amt = long_amt / 2
                        de_longcare_amt = (Int(long_amt / 10)) * 10
                    else
                        de_longcare_amt = 0
                 end if

                 '고용보험 계산 : 비과세 포함한 금액으로 계산
                 if incom_go_yn = "여" then
                        'epi_amt = inc_st_amt * (epi_emp / 100)
		                epi_amt = pmg_give_tot * (epi_emp / 100)
                        epi_amt = int(epi_amt)
                        de_epi_amt = (int(epi_amt / 10)) * 10
                    else
		                de_epi_amt = 0
                 end if

                 '지방소득세
                 we_tax = inc_incom * (10 / 100)
                 we_tax = int(we_tax)
                 de_wetax = (int(we_tax / 10)) * 10

                 de_deduct_total = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt + de_income_tax + de_wetax
                 pmg_curr_pay = pmg_give_total - de_deduct_total
	      end if
    %>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td><%=rs("emp_name")%>&nbsp;</td>
                                <td><%=pmg_grade%>&nbsp;</td>
                                <td><%=pmg_position%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_in_date%>&nbsp;</td>
                                <td><%=pmg_org_name%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_base_pay,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_give_total,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(de_deduct_total,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_curr_pay,0)%>&nbsp;</td>
                                <td class="left"><%=pmg_company%>-<%=pmg_bonbu%>-<%=pmg_saupbu%>-<%=pmg_team%></td>
							</tr>
    <%
		rs.movenext()
	 loop
	 rs.close()
	%>
						</tbody>
					</table>
				</div>
	            <%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1

                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
                   	<td width="25%">
					<div class="btnleft">
                    <a href="insa_excel_pay_month_batch.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_yymm_to=<%=pmg_yymm_to%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_month_batch.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_yymm_to=<%=pmg_yymm_to%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                        <% if intstart > 1 then %>
                            <a href="insa_pay_month_batch.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_yymm_to=<%=pmg_yymm_to%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                        <% end if %>
                        <% for i = intstart to intend %>
                            <% if i = int(page) then %>
                                <b>[<%=i%>]</b>
                            <% else %>
                                <a href="insa_pay_month_batch.asp?page=<%=i%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_yymm_to=<%=pmg_yymm_to%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                            <% end if %>
                        <% next %>
                        <% if 	intend < total_page then %>
                            <a href="insa_pay_month_batch.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_yymm_to=<%=pmg_yymm_to%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_month_batch.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&pmg_yymm_to=<%=pmg_yymm_to%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                            <%	else %>
                            [다음]&nbsp;[마지막]
                        <% end if %>
                    </div>
                    </td>
                    <td width="25%">
					<div class="btnRight">
                    <a href="#" onClick="pay_month_transe('pmg_yymm','view_condi','pmg_yymm_to','to_date');return false;" class="btnType04">급여이월자료 등록</a>
					</div>
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="pmg_yymm1" value="<%=pmg_yymm%>" ID="Hidden1">
                  <input type="hidden" name="pmg_yymm_to1" value="<%=pmg_yymm_to%>" ID="Hidden1">
                  <input type="hidden" name="view_condi1" value="<%=view_condi%>" ID="Hidden1">
                  <input type="hidden" name="to_date1" value="<%=to_date%>" ID="Hidden1">
			</form>
            </form>
		</div>
	</div>
	</body>
</html>

