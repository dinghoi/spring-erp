<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

be_pg = "insa_pay_overtime_report.asp"

from_date=Request.form("from_date")
to_date=Request.form("to_date")

Page=Request("page")
view_condi = request("view_condi")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	from_date=Request.form("from_date")
    to_date=Request.form("to_date")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_condi = request("view_condi")
	from_date=request("from_date")
    to_date=request("to_date")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	'pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
	
	overtime_count = 0	
	sum_overtime_pay = 0	
end if

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
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
   Sql = "select count(*) from overtime where work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N'"
   else  
   Sql = "select count(*) from overtime where emp_company='"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N'"
end if
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from overtime where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,team,org_name,work_date,mg_ce_id ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    overtime_count = overtime_count + 1
    sum_overtime_pay = sum_overtime_pay + int(rs("overtime_amt"))
	rs.movenext()
loop
rs.close()

if view_condi = "전체" then
   Sql = "select * from overtime where work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,team,org_name,work_date,mg_ce_id ASC limit "& stpage & "," &pgsize 
   else  
   Sql = "select * from overtime where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' and cancel_yn = 'N' ORDER BY emp_company,team,org_name,work_date,mg_ce_id ASC limit "& stpage & "," &pgsize 
end if
Rs.Open Sql, Dbconn, 1

title_line = ""+ view_condi +" - 야·특근 현황(수당) "
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
				return "0 1";
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
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_overtime_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (org_level = '회사') ORDER BY org_code ASC"
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
								<strong>야·특근기간(From) : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>(To) : </strong>
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
							<col width="14%" >
							<col width="6%" >
							<col width="13%" >
							<col width="10%" >
							<col width="13%" >
                            <col width="8%" >
                            <col width="9%" >
                            <col width="5%" >
							<col width="6%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">소속</th>
								<th scope="col">구분</th>
								<th scope="col">작업일시</th>
								<th scope="col">고객사 명</th>
								<th scope="col">지점명</th>
								<th scope="col">작업자</th>
                                <th scope="col">전자결재No.</th>
                                <th scope="col">금액</th>
                                <th scope="col">AS No.</th>
								<th scope="col">작업내용</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

                              emp_no = rs("mg_ce_id")
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
							  end if
							  rs_emp.close()

	           			%>
							<tr>
								<td class="left"><%=rs("team")%>-<%=rs("org_name")%></td>
                                <td class="left"><%=rs("cost_detail")%></td>
                                <td class="left"><%=rs("work_date")%>&nbsp;<%=mid(rs("from_time"),1,2)%>:<%=mid(rs("from_time"),3,2)%>∼<%=mid(rs("to_time"),1,2)%>:<%=mid(rs("to_time"),3,2)%></td>
                                <td class="left"><%=rs("company")%></td>
                                <td class="left"><%=rs("dept")%></td>
                                <td><%=emp_name%>(<%=rs("mg_ce_id")%>)</td>
                                <td class="left">연장/휴일-<%=rs("sign_no")%></td>
                                <td class="right"><%=formatnumber(rs("overtime_amt"),0)%></td>
                                <td><%=rs("acpt_no")%></td>
                                <td class="left"><%=rs("work_gubun")%>-<%=rs("work_memo")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
                            <tr>
								<th colspan="4" class="first">합&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;계</th>
							    <th colspan="2" class="right"><%=formatnumber(overtime_count,0)%>&nbsp;명</th>
							    <th colspan="2" class="right"><%=formatnumber(sum_overtime_pay,0)%>&nbsp;원</th>
							    <th colspan="2">&nbsp;</th>
						    </tr>
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
                    <a href="insa_excel_pay_overtime.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                  <div id="paging">
                        <a href = "insa_pay_overtime_report.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_overtime_report.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_overtime_report.asp?page=<%=i%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_overtime_report.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_overtime_report.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
                    <td>
				    <td width="20%">
					<div class="btnCenter">
                    <a href="#" onClick="pop_Window('insa_pay_overtime_save.asp?view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&pmg_yymm=<%=pmg_yymm%>','pop_report','scrollbars=yes,width=1050,height=500')" class="btnType04">야특근수당 처리</a>
                    </div>
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

