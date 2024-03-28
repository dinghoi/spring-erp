<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

be_pg = "insa_pay_month_pay_mg.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")
owner_view=request("owner_view")

ck_sw=Request("ck_sw")

if ck_sw = "y" then
	view_condi = request("view_condi")
	owner_view=request("owner_view")
	condi = request("condi")
	pmg_yymm=request("pmg_yymm")
    to_date=request("to_date") 
else
	view_condi = request.form("view_condi")
	owner_view=Request.form("owner_view")
	condi = request.form("condi")
	pmg_yymm=Request.form("pmg_yymm")
    to_date=Request.form("to_date")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	condi = ""
	owner_view = "C"
	ck_sw = "n"
	curr_dd = cstr(datepart("d",now))
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
'매월 말일 구하기
   datYear = mid(cstr(pmg_yymm),1,4)
   datMonth = mid(cstr(pmg_yymm),5,2)
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
'   to_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + cstr(exec_LastDay)
   
   to_date = ""
end if

give_date = to_date '지급일
'당월 입사일이 15일 이전이면 당월 급여대상임
st_es_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + "01"
st_in_date = mid(cstr(pmg_yymm),1,4) + "-" + mid(cstr(pmg_yymm),5,2) + "-" + "16"

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

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'Sql = "SELECT * FROM k1_memb where "+condi_sql+"mg_group = '"+mg_group+"' ORDER BY user_name ASC"
'where_sql = " WHERE isNull(emp_end_date) or emp_end_date = '1900-01-01'"  emp_in_date >= '"+st_in_date+"' and emp_in_date <= '"+st_in_date+"'

if condi = "" then
      Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"')  and (emp_pay_id <> '5') and (emp_no < '900000')"
   else  
      if owner_view = "C" then 
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"') and (emp_pay_id <> '5') and (emp_name like '%"&condi&"%')"
         else
            Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"') and (emp_pay_id <> '5') and (emp_no = '"&condi&"')"
	  end if
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if condi = "" then
      Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"')  and (emp_pay_id <> '5') and (emp_no < '900000') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize 
   else  
      if owner_view = "C" then 
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"') and (emp_pay_id <> '5') and (emp_name like '%"&condi&"%') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize 
         else
            Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01' or emp_end_date >= '"&st_in_date&"') and (emp_in_date < '"&st_in_date&"') and (emp_company = '"&view_condi&"') and (emp_pay_id <> '5') and (emp_no = '"&condi&"') ORDER BY emp_in_date,emp_no ASC limit "& stpage & "," &pgsize 
	  end if
end if

'Response.write Sql
Rs.Open Sql, Dbconn, 1

title_line = " 급여자료 입력 "

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
		    $(function() {  $( "#datepicker" ).datepicker();
							$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
							$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {  $( "#datepicker1" ).datepicker();
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
			
			function pay_month_del(val, val2, val3, val4) {

            if (!confirm("정말 삭제하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.in_empno1.value = val;
			document.frm.in_name1.value = val2;
			document.frm.pmg_yymm1.value = val3;
			document.frm.view_condi1.value = val4;
		
            document.frm.action = "insa_pay_month_del.asp";
            document.frm.submit();
            }	
			
			function pay_month_tax_cal(val, val2, val3, val4, val5) {

            if (!confirm("급여 세금계산처리를 하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.pmg_yymm1.value = document.getElementById(val).value;
			document.frm.view_condi1.value = document.getElementById(val2).value;
			document.frm.in_empno1.value = val3;
			document.frm.in_name1.value = val4;
			document.frm.owner_view1.value = val5;
			
            document.frm.action = "insa_pay_month_tax_calcu.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_month_pay_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
								Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:130px">
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
								<strong>귀속년월 : </strong>
                                <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
								<label>
								<strong>지급일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">사번
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">성명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="condi" type="text" id="condi" value="<%=condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
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
                            <col width="3%" >
                            <col width="3%" >
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
                                <th scope="col">급여</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  emp_no = rs("emp_no")
	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                        <%
						      dt_ck = "0"
							  
							  Sql = "SELECT * FROM pay_month_give where pmg_yymm = '"&pmg_yymm&"' and pmg_emp_no = '"&emp_no&"' and pmg_id = '1' and (pmg_company = '"+view_condi+"')"
                              'Response.write Sql
                              Set rs_give = DbConn.Execute(SQL)
							  if not rs_give.eof then
                                    pmg_base_pay = rs_give("pmg_base_pay")
									pmg_give_tot = rs_give("pmg_give_total")
									dt_ck = "1"
	                             else
                                    pmg_base_pay = 0
								    pmg_give_tot = 0
                              end if
                              rs_give.close()
                        %>
                                <td class="right"><%=formatnumber(pmg_base_pay,0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pmg_give_tot,0)%>&nbsp;</td>
                        <%
                              Sql = "SELECT * FROM pay_month_deduct where de_yymm = '"&pmg_yymm&"' and de_emp_no = '"&emp_no&"' and de_id = '1' and (de_company = '"+view_condi+"')"
                              'Response.write Sql
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
									de_deduct_tot = Rs_dct("de_deduct_total")
	                             else
                                    de_deduct_tot = 0
                              end if
                              Rs_dct.close()
							  
							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
                        %>
                              
                                <td class="right"><%=formatnumber(de_deduct_tot,0)%>&nbsp;</td>
                                
                                <td class="right"><a href="#" onClick="pop_Window('insa_pay_person_view.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&pmg_yymm=<%=pmg_yymm%>&pmg_date=<%=give_date%>&pmg_company=<%=rs("emp_company")%>&pmg_org_code=<%=rs("emp_org_code")%>&pmg_org_name=<%=rs("emp_org_name")%>&pmg_grade=<%=rs("emp_grade")%>&pmg_position=<%=rs("emp_position")%>','insa_pay_person_pop','scrollbars=yes,width=750,height=700')"><%=formatnumber(pmg_curr_pay,0)%></a>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
                                
                   <% if emp_payend = "N" then 
						            if dt_ck = "0" then    
						       %>              
                          <td><a href="#" onClick="pop_Window('insa_pay_month_give_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&pmg_yymm=<%=pmg_yymm%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>&u_type=<%=""%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=700')">입력</a></td>
                   <%   else  %>
                          <td><a href="#" onClick="pop_Window('insa_pay_month_give_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&pmg_yymm=<%=pmg_yymm%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>&u_type=<%="U"%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=700')">수정</a></td>
                   <%   end if
				              else 
				           %>
                        <td>&nbsp;</td>
                   <% end if %>
                   
                   <% if emp_payend = "N" and dt_ck = "1"  then  %>
                      <td>
                      <a href="#" onClick="pay_month_del('<%=rs("emp_no")%>', '<%=rs("emp_name")%>', '<%=pmg_yymm%>', '<%=view_condi%>');return false;">삭제</a>
                      </td>
                   <% else %>
                      <td>&nbsp;</td>
                   <% end if %>                   
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
                  	<td width="15%">
					<div class="btnCenter">
                    <% 'insa_excel_pay_month_ledger %>
                    <a href="insa_excel_pay_transe_list.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&owner_view=<%=owner_view%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_month_pay_mg.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_month_pay_mg.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	          <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_month_pay_mg.asp?page=<%=i%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	          <% if	intend < total_page then %>
                        <a href="insa_pay_month_pay_mg.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_month_pay_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&owner_view=<%=owner_view%>&condi=<%=condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
					<div class="btnRight">
                    <% if emp_payend = "N" then 
					      if owner_view = "T" then 
                              emp_no = condi
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
							  end if
							  rs_emp.close()
				    %>
					<a href="#" onClick="pop_Window('insa_pay_month_give_add.asp?emp_no=<%=emp_no%>&emp_name=<%=emp_name%>&pmg_yymm=<%=pmg_yymm%>&give_date=<%=give_date%>&view_condi=<%=view_condi%>&u_type=<%=""%>','insa_pay_give_add_pop','scrollbars=yes,width=750,height=700')" class="btnType04">지급입력</a>
                    <a href="#" onClick="pay_month_tax_cal('pmg_yymm','view_condi','<%=emp_no%>','<%=emp_name%>','<%=owner_view%>');return false;" class="btnType04">급여 세금계산 처리</a>
                    <%     else 
					          if condi = "" then  %>   
                    <a href="#" onClick="pay_month_tax_cal('pmg_yymm','view_condi','in_empno','in_name','<%=owner_view%>');return false;" class="btnType04">급여 세금계산 일괄처리</a>
                    <%        end if
					     end if
					   end if %>   
					</div>                  
                    </td> 
			      </tr>
				  </table>
                  <input type="hidden" name="view_condi1" value="<%=view_condi%>" ID="Hidden1">
                  <input type="hidden" name="pmg_yymm1" value="<%=pmg_yymm%>" ID="Hidden1">
                  <input type="hidden" name="in_empno1" value="<%=emp_no%>" ID="Hidden1">
                  <input type="hidden" name="in_name1" value="<%=emp_name%>" ID="Hidden1">
                  <input type="hidden" name="owner_view1" value="<%=owner_view%>" ID="Hidden1">
        	</form>
		</div>				
	</div>        				
	</body>
</html>

