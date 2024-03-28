<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(100,2)

be_pg = "insa_pay_albacost_report.asp"

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
    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
    to_date=request("to_date") 
end if

if view_condi = "" then
	view_condi = "케이원정보통신"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
	sum_alba_pay = 0
	sum_alba_trans = 0
	sum_alba_meals = 0
	sum_alba_other = 0
	sum_alba_other2 = 0
	sum_alba_give_total = 0
	sum_tax_amt1 = 0
	sum_tax_amt2 = 0
	sum_de_other = 0
	sum_pay_amount = 0
	
	pay_count = 0	
	
end if

give_date = to_date '지급일

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(100,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(100,2) = view_month
for i = 1 to 99
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 100 - i
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
Set Rs_alb = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_alco = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select count(*) from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    draft_no = rs("draft_no")
    alba_give_total = rs("alba_give_total")
    pay_count = pay_count + 1
				  
    sum_alba_pay = sum_alba_pay + int(rs("alba_pay"))
    sum_alba_trans = sum_alba_trans + int(rs("alba_trans"))
    sum_alba_meals = sum_alba_meals + int(rs("alba_meals"))
    sum_alba_other = sum_alba_other + int(rs("alba_other"))
    sum_alba_give_total = sum_alba_give_total + int(rs("alba_give_total"))
    sum_tax_amt1 = sum_tax_amt1 + int(rs("tax_amt1"))
    sum_tax_amt2 = sum_tax_amt2 + int(rs("tax_amt2"))
    sum_de_other = sum_de_other + int(rs("de_other"))
    sum_pay_amount = sum_pay_amount + int(rs("pay_amount"))
	sum_deduct_tot = sum_deduct_tot + (int(rs("tax_amt1")) + int(rs("tax_amt2")) + int(rs("de_other")))
	
	
	rs.movenext()
loop
rs.close()

Sql = "select * from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC limit "& stpage & "," &pgsize 

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 사업소득현황 "

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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_alba_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_albacost_report.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 100 to 1 step -1	%>
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
							<col width="8%" >
                            <col width="6%" >
                            <col width="*" >
                            <col width="8%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="8%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
							<col width="8%" > 
                            <col width="8%" >
                            <col width="3%" >
						</colgroup>
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">성명</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">등록일</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">구분</th>
				               <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">사업소득 및 제수당</th>
                               <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;">공제</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">차인지급액</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">조회</th>
			                </tr>
                            <tr>
								<td scope="col" style=" border-left:1px solid #e3e3e3;">사업소득</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">교통비</td>  
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">식대</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">기타</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">지급소계</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">소득세</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">지방소득세</td>
								<td scope="col" style=" border-bottom:1px solid #e3e3e3;">기타공제</td>
                                <td scope="col" style=" border-bottom:1px solid #e3e3e3;">공제소계</td>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
							  draft_no = rs("draft_no")
							  alba_give_total = rs("alba_give_total")

							  'sub_give_hap = int(rs("alba_pay")) + int(rs("alba_trans")) + int(rs("alba_meals")) + int(rs("alba_other"))
							  alba_give_total = rs("alba_give_total")
							  
							  Sql = "SELECT * FROM emp_alba_mst where draft_no = '"&draft_no&"'"
                              Set Rs_alb = DbConn.Execute(SQL)
		                      if not Rs_alb.eof then
		                    		draft_date = Rs_alb("draft_date")
	                             else
	                    			draft_date = ""
                              end if
                              Rs_alb.close()
							  
	           			 %>
							<tr>
								<td class="first"><%=rs("draft_man")%>(<%=rs("draft_no")%>)</td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=draft_date%></td>
                                <td style=" border-left:1px solid #e3e3e3;"><%=rs("draft_tax_id")%></td>
                                <td class="right"><%=formatnumber(rs("alba_pay"),0)%></td>
                                <td class="right"><%=formatnumber(rs("alba_trans"),0)%></td>
                                <td class="right"><%=formatnumber(rs("alba_meals"),0)%></td>
                                <td class="right"><%=formatnumber(rs("alba_other"),0)%></td>
                                <td class="right"><%=formatnumber(rs("alba_give_total"),0)%></td>
                         <%
							  sub_de_hap = int(rs("tax_amt1")) + int(rs("tax_amt2")) + int(rs("de_other"))
							  'pay_amount = alba_give_total - sub_de_hap
							  pay_amount = rs("pay_amount")

                         %>
                                <td class="right"><%=formatnumber(rs("tax_amt1"),0)%></td>
                                <td class="right"><%=formatnumber(rs("tax_amt2"),0)%></td>
                                <td class="right"><%=formatnumber(rs("de_other"),0)%></td>
                                <td class="right"><%=formatnumber(sub_de_hap,0)%></td>
                                <td class="right"><%=formatnumber(pay_amount,0)%></td>
                                <td class="right"><a href="#" onClick="pop_Window('insa_pay_albacost_view.asp?draft_no=<%=rs("draft_no")%>&draft_man=<%=rs("draft_man")%>&pmg_yymm=<%=pmg_yymm%>&pmg_date=<%=give_date%>&company=<%=rs("company")%>','insa_pay_albacost_pop','scrollbars=yes,width=750,height=500')">상세</a></td>
                                
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_alba_give_total - sum_deduct_tot
						
						%>
                          	<tr>
                                <th class="first">총계</th>
                                <th colspan="2" class="right"><%=formatnumber(pay_count,0)%>&nbsp;명</th>
                                <th class="right"><%=formatnumber(sum_alba_pay,0)%></th>
                                <th class="right"><%=formatnumber(sum_alba_trans,0)%></th>
                                <th class="right"><%=formatnumber(sum_alba_meals,0)%></th>
                                <th class="right"><%=formatnumber(sum_alba_other,0)%></th>
                                <th class="right"><%=formatnumber(sum_alba_give_total,0)%></th>
                                <th class="right"><%=formatnumber(sum_tax_amt1,0)%></th>
                                <th class="right"><%=formatnumber(sum_tax_amt2,0)%></th>
                                <th class="right"><%=formatnumber(sum_de_other,0)%></th>
                                <th class="right"><%=formatnumber(sum_deduct_tot,0)%></th>
                                <th class="right"><%=formatnumber(sum_pay_amount,0)%></th>
                                <th class="right">&nbsp;</th>
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
                  	<td width="15%">
					<div class="btnCenter">
                    <a href="insa_excel_pay_albacost_report.asp?view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href = "insa_pay_albacost_report.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_pay_albacost_report.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_pay_albacost_report.asp?page=<%=i%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_pay_albacost_report.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_pay_albacost_report.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&pmg_yymm=<%=pmg_yymm%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
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

