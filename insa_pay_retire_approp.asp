<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim pay_tab(5)
dim pay_pay(5)
dim bonus_tab(5)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

be_pg = "insa_pay_retire_approp.asp"

to_date=Request.form("to_date")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
    to_date=request("to_date")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	
	for i = 1 to 5
	    pay_tab(i) = ""
     	pay_pay(i) = 0
    	bonus_tab(i) = 0
    next
end if

target_date = to_date

t_year = int(mid(cstr(target_date),1,4))
t_month = int(mid(cstr(target_date),6,2))
t_day = int(mid(cstr(target_date),9,2))
tcal_month = mid(cstr(target_date),1,4) + mid(cstr(target_date),6,2)
tcal_day = cstr(t_day)

'pay_tab(3) = cstr(tcal_month)
'tcal_month = cstr(int(tcal_month) - 1)
'pay_tab(2) = cstr(tcal_month)
'tcal_month = cstr(int(tcal_month) - 1)
'pay_tab(1) = cstr(tcal_month)

'tcal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
pay_tab(3) = tcal_month
for i = 1 to 2
	tcal_month = cstr(int(tcal_month) - 1)
	if mid(tcal_month,5) = "00" then
		cal_year = cstr(int(mid(tcal_month,1,4)) - 1)
		tcal_month = cal_year + "12"
	end if	 
	j = 3 - i
	pay_tab(j) = tcal_month
next

tar1_date = cstr(mid(pay_tab(3),1,4) + "-" + mid(pay_tab(3),5,2) + "-" + tcal_day)
fir1_date = cstr(mid(pay_tab(1),1,4) + "-" + mid(pay_tab(1),5,2) + "-" + "01")
work1_cnt = int(datediff("d", fir1_date, tar1_date)) + 1
pay_tab(5) = work1_cnt


pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
   Sql = "select count(*) from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000')"
   else  
   Sql = "select count(*) from emp_master where emp_company='"+view_condi+"' and (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000')"
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_condi = "전체" then
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_company,emp_no ASC limit "& stpage & "," &pgsize 
   else  
   Sql = "select * from emp_master where emp_company = '"+view_condi+"' and (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_company,emp_no ASC limit "& stpage & "," &pgsize 
end if
Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - 퇴직급여 추계액(충당금)내역 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
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
				return "1 1";
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
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_end_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_retire_approp.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
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
								<strong>퇴직급여 추계기준일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
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
							<col width="8%" >
							<col width="10%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="9%" >
                            <col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="4%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="5%" >
                            <col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">회사</th>
								<th rowspan="2" scope="col">소속</th>
								<th rowspan="2" scope="col">성명</th>
								<th rowspan="2" scope="col">직급</th>
								<th rowspan="2" scope="col">직책</th>
								<th rowspan="2" scope="col">최초입사일</th>
                                <th colspan="3" scope="col" style=" border-bottom:1px solid #e3e3e3;">최근3개월급여</th>
                                <th rowspan="2" scope="col">일수</th>
                                <th rowspan="2" scope="col">평균임금</th>
                                <th rowspan="2" scope="col">월평균임금</th>
								<th rowspan="2" scope="col">근속연수</th>
								<th rowspan="2" scope="col">퇴직추계액</th>
							</tr>
                            <tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;"><%=mid(pay_tab(1),1,4)%>년&nbsp;<%=mid(pay_tab(1),5,2)%>월</th>
								<th scope="col"><%=mid(pay_tab(2),1,4)%>년&nbsp;<%=mid(pay_tab(2),5,2)%>월</th>
								<th scope="col"><%=mid(pay_tab(3),1,4)%>년&nbsp;<%=mid(pay_tab(3),5,2)%>월</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                            emp_no = rs("emp_no")
							emp_first_date = rs("emp_first_date")
                            if rs("emp_first_date") = "" then 
                                   emp_first_date = rs("emp_in_date")
                            end if
                            'target_date = "2015-02-20"
                            'emp_first_date = "2013-11-10"
						
					    	f_year = int(mid(cstr(emp_first_date),1,4))
                            f_month = int(mid(cstr(emp_first_date),6,2))
                            f_day = int(mid(cstr(emp_first_date),9,2))
                            fcal_day = cstr(f_day)
                            cf_date = emp_first_date '중간퇴직처리를 하기위한
						
						    year_cnt = datediff("yyyy", emp_first_date, target_date)
                            mon_cnt = datediff("m", emp_first_date, target_date)
                            day_cnt = datediff("d", emp_first_date, target_date) 

                            year_cnt = int(year_cnt) + 1
                            mon_cnt = int(mon_cnt) + 1
                            day_cnt = int(day_cnt) + 1
							if day_cnt < 365 then
							        gunsok_cnt = 0
							   else
									gunsok_cnt = formatnumber((day_cnt / 365),1)
							end if
							
							for i = 1 to 3
	                            p_yymm = pay_tab(i)
		                        if p_yymm <> "" then
		                              Sql = "select * from pay_month_give where (pmg_yymm = '"+p_yymm+"' ) and (pmg_id = '1') and (pmg_emp_no = '"+emp_no+"') and (pmg_company = '"+view_condi+"')"
                                      Rs_give.Open Sql, Dbconn, 1
                                      Set Rs_give = DbConn.Execute(SQL)
                                      if not Rs_give.eof then
                                             pmg_give_tot = int(Rs_give("pmg_give_total"))	
                                         else
                                             pmg_give_tot = 0
                                      end if
			                          Rs_give.close()
			                         
									  Sql = "select * from pay_month_deduct where (de_yymm = '"+p_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
                                      Set Rs_dct = DbConn.Execute(SQL)
                                      if not Rs_dct.eof then
                                             de_deduct_tot = int(Rs_dct("de_deduct_total"))	
                                         else
                                             de_deduct_tot = 0
                                      end if
                                      Rs_dct.close()
		                        	  pay_curr_amt = pmg_give_tot - de_deduct_tot
		                        	  pay_pay(i) = pay_curr_amt
	                           end if
                            next
							
							pay_sum = pay_pay(1)+pay_pay(2)+pay_pay(3)
							eot_average_pay = int(pay_sum / pay_tab(5))
							eot_month_pay = eot_average_pay * 30
							retire_pay = int(eot_month_pay * gunsok_cnt)

	           			%>
							<tr>
								<td class="first"><%=rs("emp_company")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_name")%>(<%=rs("emp_no")%>)&nbsp;</td> 
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=emp_first_date%>&nbsp;</td>
                                <td><%=formatnumber(pay_pay(1),0)%>&nbsp;</td>
                                <td><%=formatnumber(pay_pay(2),0)%>&nbsp;</td>
                                <td><%=formatnumber(pay_pay(3),0)%>&nbsp;</td>
                                <td><%=pay_tab(5)%>&nbsp;</td>
                                <td><%=formatnumber(eot_average_pay,0)%>&nbsp;</td>
                                <td><%=formatnumber(eot_month_pay,0)%>&nbsp;</td>
                                <td><%=formatnumber(gunsok_cnt,1)%>&nbsp;</td>
                                <td><%=formatnumber(retire_pay,0)%>&nbsp;</td>
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
                  	<td width="20%">
					<div class="btnCenter">
                    <a href="insa_excel_pay_retire_approp.asp?view_condi=<%=view_condi%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_pay_retire_approp.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_retire_approp.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_retire_approp.asp?page=<%=i%>&view_condi=<%=view_condi%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_retire_approp.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_retire_approp.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
				    <td width="20%">
					<div class="btnCenter">
                    <% if user_id = "900002" then %>
                    <a href="#" onClick="pop_Window('insa_pay_retire_approp_print.asp?view_condi=<%=view_condi%>&to_date=<%=to_date%>','pop_report','scrollbars=yes,width=1050,height=500')" class="btnType04">출력</a>
                    <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

