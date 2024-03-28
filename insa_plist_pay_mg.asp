<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

be_pg = "insa_plist_pay_mg.asp"
user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")
emp_company = request.cookies("nkpmg_user")("coo_emp_company")

if ck_sw = "y" then
	rever_yyyy=request("rever_yyyy")
  else
	rever_yyyy=Request.form("rever_yyyy")
end if

if rever_yyyy = "" then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	rever_yyyy = mid(cstr(from_date),1,4)
end if

' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

' 분기 테이블 생성
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "년 " + cstr(mid(quarter_tab(8,1),5,1)) + "/4분기"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "년 " + cstr(mid(quarter_tab(i,1),5,1)) + "/4분기"
next

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

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'Sql = "SELECT * FROM pay_month_give where pmg_yymm like '%"&rever_yyyy&"%' and pmg_emp_no = '"&emp_no&"' and pmg_id = '1' and (pmg_company = '"+emp_company+"')"
Sql = "SELECT * FROM pay_month_give where pmg_yymm like '%"&rever_yyyy&"%' and pmg_emp_no = '"&emp_no&"' and pmg_id = '1'"
Rs.Open Sql, Dbconn, 1

title_line = " 급여 조회 "

etc_code = "9999"

sql = "select * from emp_etc_code where emp_etc_code = '" + etc_code + "'"
Rs_etc.Open Sql, Dbconn, 1
emp_payend_date = Rs_etc("emp_payend_date")
emp_payend_yn = Rs_etc("emp_payend_yn")

Rs_etc.close()

if pmg_yymm > emp_payend_date then
       emp_payend = "N"
   else
	   emp_payend = "Y"
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
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
		</script>
		<script type="text/javascript">
		    $(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
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
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_plist_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_plist_pay_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                                <label>
								<strong>귀속년도 : </strong>
                                    <select name="rever_yyyy" id="rever_yyyy" type="text" value="<%=rever_yyyy%>" style="width:90px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If rever_yyyy = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                    </select>
  							    </label>
                                <label>
								<strong>사번 : </strong>
                                <input name="emp_no" type="text" value="<%=emp_no%>" style="width:70px" id="emp_no" readonly="true">
								</label>
                                <label>
								<strong>성명 : </strong>
                                <input name="emp_name" type="text" value="<%=user_name%>" style="width:90px" id="emp_name" readonly="true">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="*" >
							<col width="8%" >
							<col width="8%" >
                            <col width="8%" >
                            <col width="8%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="8%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">년월</th>
								<th scope="col">소속</th>
								<th scope="col">직급</th>
								<th scope="col">기본급</th>
                                <th scope="col">식대</th>
								<th scope="col">연장수당</th>
                                <th scope="col" style="background:#E0FFFF">지급액계</th>
                                <th scope="col" style="background:#E0FFFF">공제액계</th>
                                <th scope="col" style="background:#FFFFE6">차인지급액</th>
                                <th scope="col">상세조회</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
                              de_yymm = rs("pmg_yymm")
						  if emp_payend_date >= de_yymm then
							  de_company = rs("pmg_company")
							  pmg_give_tot = rs("pmg_give_total")

                              Sql = "SELECT * FROM pay_month_deduct where de_yymm = '"&de_yymm&"' and de_emp_no = '"&emp_no&"' and de_id = '1' and (de_company = '"+emp_company+"')"
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
									de_nps_amt = Rs_dct("de_nps_amt")
                                    de_nhis_amt = Rs_dct("de_nhis_amt")
                                    de_epi_amt = Rs_dct("de_epi_amt")
		                            de_longcare_amt = Rs_dct("de_longcare_amt")
                                    de_income_tax = Rs_dct("de_income_tax")
                                    de_wetax = Rs_dct("de_wetax")
									de_year_incom_tax = Rs_dct("de_year_incom_tax")
                                    de_year_wetax = Rs_dct("de_year_wetax")
                                    de_other_amt1 = Rs_dct("de_other_amt1")
		                            de_special_tax = Rs_dct("de_special_tax")
                                    de_saving_amt = Rs_dct("de_saving_amt")
                                    de_sawo_amt = Rs_dct("de_sawo_amt")
                                    de_johab_amt = Rs_dct("de_johab_amt")
                                    de_hyubjo_amt = Rs_dct("de_hyubjo_amt")
                                    de_school_amt = Rs_dct("de_school_amt")
                                    de_nhis_bla_amt = Rs_dct("de_nhis_bla_amt")
                                    de_long_bla_amt = Rs_dct("de_long_bla_amt")
		                            de_deduct_tot = Rs_dct("de_deduct_total")
	                             else
                                    de_nps_amt = 0
                                    de_nhis_amt = 0
                                    de_epi_amt = 0
		                            de_longcare_amt = 0
                                    de_income_tax = 0
                                    de_wetax = 0
									de_year_incom_tax = 0
                                    de_year_wetax = 0
                                    de_other_amt1 = 0
                                    de_sawo_amt = 0
                                    de_hyubjo_amt = 0
                                    de_school_amt = 0
                                    de_nhis_bla_amt = 0
                                    de_long_bla_amt = 0
		                            de_deduct_tot = 0
                              end if
                              Rs_dct.close()

							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  de_insu_hap = de_nps_amt + de_nhis_amt + de_epi_amt + de_longcare_amt
							  de_tax_hap = de_income_tax + de_wetax
							  de_other_hap = de_other_amt1 + de_sawo_amt + de_hyubjo_amt + de_school_amt
							  de_bla_amt = de_nhis_bla_amt + de_long_bla_amt

	           			%>
							<tr>
								<td class="first"><%=mid(cstr(rs("pmg_yymm")),1,4)%>년&nbsp;<%=mid(cstr(rs("pmg_yymm")),5,2)%>월&nbsp;</td>
                                <td class="left"><%=rs("pmg_company")%>&nbsp;-&nbsp;<%=rs("pmg_org_name")%>(<%=rs("pmg_org_code")%>)&nbsp;</td>
                                <td><%=rs("pmg_grade")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_base_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_meals_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_overtime_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("pmg_give_total"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(de_deduct_tot,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_curr_pay,0)%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('insa_pay_person_view.asp?emp_no=<%=emp_no%>&emp_name=<%=user_name%>&pmg_yymm=<%=rs("pmg_yymm")%>&pmg_date=<%=rs("pmg_date")%>&pmg_company=<%=rs("pmg_company")%>&pmg_org_code=<%=rs("pmg_org_code")%>&pmg_org_name=<%=rs("pmg_org_name")%>&pmg_grade=<%=rs("pmg_grade")%>&pmg_position=<%=rs("pmg_position")%>','insa_pay_person_pop','scrollbars=yes,width=750,height=700')">조회</a>&nbsp;</td>
							</tr>
						<%
						  end if
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</div>
	</body>
</html>

