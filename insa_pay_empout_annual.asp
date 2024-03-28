<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

be_pg = "insa_pay_empout_annual.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

Page=Request("page")
from_date=Request.form("from_date")
to_date=Request.form("to_date")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "y" then
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date")
	pmg_yymm=request("pmg_yymm")
  else
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
	pmg_yymm=Request.form("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
end if

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
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
   Sql = "select count(*) from emp_master where emp_end_date >= '"+from_date+"' and emp_end_date <= '"+to_date+"'"
   else  
   Sql = "select count(*) from emp_master where emp_company='"+view_condi+"' and emp_end_date >= '"+from_date+"' and emp_end_date <= '"+to_date+"'"
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

if view_condi = "전체" then
   Sql = "select * from emp_master where emp_end_date >= '"+from_date+"' and emp_end_date <= '"+to_date+"' ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
   else  
   Sql = "select * from emp_master where emp_company = '"+view_condi+"' and emp_end_date >= '"+from_date+"' and emp_end_date <= '"+to_date+"' ORDER BY emp_no,emp_name ASC limit "& stpage & "," &pgsize 
end if
Rs.Open Sql, Dbconn, 1

title_line = " 퇴직자 연차수당 정산 "

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
			
			function yuncha_add(val, val2, val3) {

            if (!confirm("연차휴가일수를 처리하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.rever_yyyy.value = document.getElementById(val).value;
			document.frm.target_date.value = document.getElementById(val2).value;
            document.frm.view_condi.value = document.getElementById(val3).value;
			
            document.frm.action = "insa_year_leave_save.asp";
            document.frm.submit();
            }	
			
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_end_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_empout_annual.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈ 검색◈</dt>
                        <dd>
                            <p>
                             <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '회사' ORDER BY org_code ASC"
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
								<strong>퇴직일(From) : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>퇴직일(To) : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
								<label>
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="12%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성명</th>
								<th scope="col">직급</th>
                                <th scope="col">입사일</th>
                                <th scope="col">연차기산일</th>
                                <th scope="col">퇴직일</th>
                                <th scope="col">소속</th>
								<th scope="col">근속<br>년수</th>
                                <th scope="col">근속<br>개월</th>
                                <th scope="col">발생<br>연차일수</th>
                                <th scope="col">누적<br>연차</th>
								<th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                                <th scope="col">년차</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  
							  if rs("emp_yuncha_date") = "1900-01-01" or isNull(rs("emp_yuncha_date")) then
                                    emp_yuncha_date = rs("emp_in_date")
                                 else 
                                    emp_yuncha_date = rs("emp_yuncha_date")
                              end if
							   
							  ' 근속년수
							  target_date = rs("emp_end_date") + 1
                              year_cnt = datediff("yyyy", emp_yuncha_date, target_date)
							  
							  ' 연차일수
							  target_date = rs("emp_end_date")
							  if (datediff("d", emp_yuncha_date, target_date) + 1) / 365 < 1 then
							         yun_day = datediff("m", emp_yuncha_date, target_date) 
							     else
								     yun_day = round((((datediff("d", emp_yuncha_date, target_date) + 1) / 365) / 2),0) + 14
							  end if
							  
							  ' 누적연차수
							  if datediff("yyyy", emp_yuncha_date, target_date) mod 2 = 1 then
							          tot_yun = round(((year_cnt ^ 2 + 58 * year_cnt - 0) / 4),0)
								 else
							          tot_yun = year_cnt / 2 * (year_cnt / 2 + 1) + 14 * year_cnt
							  end if
							  
                              mon_cnt = datediff("m", emp_yuncha_date, target_date) 
	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_yuncha_date")%>&nbsp;</td>
                                <td><%=rs("emp_end_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td class="center"><%=year_cnt%>&nbsp;</td>
                                <td class="center"><%=mon_cnt%>&nbsp;</td>
                                <td class="center"><%=yun_day%>&nbsp;</td>
                                <td class="center"><%=tot_yun%>&nbsp;</td>
                                <td class="left"><%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
                                <td><a href="#" onClick="pop_Window('insa_pay_empout_annualadd.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=rs("emp_name")%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&u_type=<%=""%>','insa_pay_empout_annual_pop','scrollbars=yes,width=750,height=400')">정산</a></td>
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
                    <a href="insa_excel_empout_annual.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_empout_annual.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_empout_annual.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_empout_annual.asp?page=<%=i%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_empout_annual.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_empout_annual.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="target_date" value="<%=target_date%>" ID="Hidden1">
			</form>
            
		</div>				
	</div>        				
	</body>
</html>
