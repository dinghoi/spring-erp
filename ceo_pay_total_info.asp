<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

dim com_tab(5)
dim pay_count(5,3)
dim overtime_pay(5,3)
dim give_amt(5,3)
dim re_pay(5,3)
dim give_tot(5,3)

be_pg = "ceo_pay_total_info.asp"

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
'    to_date=Request.form("to_date")
  else
	view_condi = request("view_condi")
	pmg_yymm=request("pmg_yymm")
'    to_date=request("to_date") 
end if

if view_condi = "" then
	view_condi = "전체"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	'pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
	
  for i = 1 to 5
     com_tab(i) = ""
	 for j = 1 to 3
	    pay_count(i,j) = 0
		overtime_pay(i,j) = 0
		give_amt(i,j) = 0
		re_pay(i,j) = 0
		give_tot(i,j) = 0
     next
  next
	
end if

give_date = to_date '지급일

' 년월 테이블생성
cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
'cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
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
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

'당월급여 집계
if view_condi = "전체" then
          com_tab(1) = "케이원정보통신"
		  com_tab(2) = "휴디스"
		  com_tab(3) = "케이네트웍스"
		  com_tab(4) = "에스유에이치"
		  com_tab(5) = "합계"
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1')"
	else	  
		  com_tab(1) = view_condi
		  com_tab(5) = "합계"
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
      for i = 1 to 5
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,1) = pay_count(i,1) + 1
				 pay_count(5,1) = pay_count(5,1) + 1
		         overtime_pay(i,1) = overtime_pay(i,1) + int(rs("pmg_overtime_pay"))
				 overtime_pay(5,1) = overtime_pay(5,1) + int(rs("pmg_overtime_pay"))
		         give_amt(i,1) = give_amt(i,1) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(5,1) = give_amt(5,1) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,1) = re_pay(i,1) + int(rs("pmg_re_pay"))
				 re_pay(5,1) = re_pay(5,1) + int(rs("pmg_re_pay"))
		         give_tot(i,1) = give_tot(i,1) + int(rs("pmg_give_total"))
				 give_tot(5,1) = give_tot(5,1) + int(rs("pmg_give_total"))
		  end if		 
	  next			 
	rs.movenext()
loop
rs.close()		

'전월 급여
bef_month = mid(cstr(pmg_yymm),1,4) + mid(cstr(pmg_yymm),5,2)
bef_month = cstr(int(bef_month) - 1)
if mid(bef_month,5) = "00" then
	bef_year = cstr(int(mid(bef_month,1,4)) - 1)
	bef_month = bef_year + "12"
end if	

if view_condi = "전체" then
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_month+"' ) and (pmg_id = '1')"
	else	  
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_month+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
      for i = 1 to 5
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,2) = pay_count(i,2) + 1
				 pay_count(5,2) = pay_count(5,2) + 1
		         overtime_pay(i,2) = overtime_pay(i,2) + int(rs("pmg_overtime_pay"))
				 overtime_pay(5,2) = overtime_pay(5,2) + int(rs("pmg_overtime_pay"))
		         give_amt(i,2) = give_amt(i,2) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(5,2) = give_amt(5,2) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,2) = re_pay(i,2) + int(rs("pmg_re_pay"))
				 re_pay(5,2) = re_pay(5,2) + int(rs("pmg_re_pay"))
		         give_tot(i,2) = give_tot(i,2) + int(rs("pmg_give_total"))
				 give_tot(5,2) = give_tot(5,2) + int(rs("pmg_give_total"))
		  end if		 
	  next			 
	rs.movenext()
loop
rs.close()		

'전년 급여
bef_yearmon = cstr(int(mid(cstr(pmg_yymm),1,4)) - 1) + mid(cstr(pmg_yymm),5,2)
if view_condi = "전체" then
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_yearmon+"' ) and (pmg_id = '1')"
	else	  
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+bef_yearmon+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
      for i = 1 to 5
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,3) = pay_count(i,3) + 1
				 pay_count(5,3) = pay_count(5,3) + 1
		         overtime_pay(i,3) = overtime_pay(i,3) + int(rs("pmg_overtime_pay"))
				 overtime_pay(5,3) = overtime_pay(5,3) + int(rs("pmg_overtime_pay"))
		         give_amt(i,3) = give_amt(i,3) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(5,3) = give_amt(5,3) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,3) = re_pay(i,3) + int(rs("pmg_re_pay"))
				 re_pay(5,3) = re_pay(5,3) + int(rs("pmg_re_pay"))
		         give_tot(i,3) = give_tot(i,3) + int(rs("pmg_give_total"))
				 give_tot(5,3) = give_tot(5,3) + int(rs("pmg_give_total"))
		  end if		 
	  next			 
	rs.movenext()
loop
rs.close()		

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 급여 전월비교분석"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>임원 정보 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
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
			
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
            <!--#include virtual = "/include/ceo_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="ceo_pay_total_info.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                <table cellpadding="0" cellspacing="0">
				  <tr>
                   	<td>
      				<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
                            <col width="15%" >
							<col width="15%" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="20%" >
						</colgroup>
						<thead>
							<tr>
								<th colspan="2" class="first" scope="col">구&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;분</th>
								<th scope="col"><%=mid(pmg_yymm,1,4)%>년&nbsp;<%=mid(pmg_yymm,5,2)%>월</th>
                                <th scope="col"><%=mid(bef_month,1,4)%>년&nbsp;<%=mid(bef_month,5,2)%>월</th>
                                <th scope="col"><%=mid(bef_yearmon,1,4)%>년&nbsp;<%=mid(bef_yearmon,5,2)%>월</th>
                                <th scope="col">비고</th>
							</tr>  
                        </thead>
                    </table>
                    </DIV>
					</td>
                  </tr>
                  <tr>
                    <td valign="top">
				    <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll;overflow-x:hidden;" onscroll="scrollAll()">
					<table cellpadding="0" cellspacing="0" class="scrollList">
                        <colgroup>
							<col width="*" >
                            <col width="15%" >
							<col width="15%" >
                            <col width="15%" >
                            <col width="15%" >
                            <col width="20%" >
						</colgroup>                        
                        <tbody>
                        <%
						b_pay_count = 0
		                b_overtime_pay = 0
		                b_give_amt = 0
		                b_re_pay = 0
		                b_give_tot = 0
						
						y_pay_count = 0
		                y_overtime_pay = 0
		                y_give_amt = 0
		                y_re_pay = 0
		                y_give_tot = 0
						
                        for i = 1 to 5 
                        	if	com_tab(i) <> "" then
						%>	
							<tr>
								<td class="first" rowspan="5"><%=com_tab(i)%></td>
                                <td>인원</td>
								<td class="right"><%=formatnumber(pay_count(i,1),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(pay_count(i,2),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(pay_count(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                        	<tr>
								<td style=" border-left:1px solid #e3e3e3;">야특근</td>
								<td class="right"><%=formatnumber(overtime_pay(i,1),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(overtime_pay(i,2),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(overtime_pay(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>   
                            <tr>
								<td style=" border-left:1px solid #e3e3e3;">급여</td>
								<td class="right"><%=formatnumber(give_amt(i,1),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(give_amt(i,2),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(give_amt(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>   
                            <tr>
								<td style=" border-left:1px solid #e3e3e3;">소급</td>
								<td class="right"><%=formatnumber(re_pay(i,1),0)%>&nbsp;</td>
								<td class="right"><%=formatnumber(re_pay(i,2),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(re_pay(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>    
                            <tr>
								<th style=" border-left:1px solid #e3e3e3;">합계</th>
								<th class="right"><%=formatnumber(give_tot(i,1),0)%>&nbsp;</th>
								<th class="right"><%=formatnumber(give_tot(i,2),0)%>&nbsp;</th>
                                <th class="right"><%=formatnumber(give_tot(i,3),0)%>&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>    
                        <%
							end if
						next
						        b_pay_count = pay_count(5,1) - pay_count(5,2)
		                        b_overtime_pay = overtime_pay(5,1) - overtime_pay(5,2)
		                        b_give_amt = give_amt(5,1) - give_amt(5,2)
		                        b_re_pay = re_pay(5,1) - re_pay(5,2)
		                        b_give_tot = give_tot(5,1) - give_tot(5,2)
								
								y_pay_count = pay_count(5,1) - pay_count(5,3)
		                        y_overtime_pay = overtime_pay(5,1) - overtime_pay(5,3)
		                        y_give_amt = give_amt(5,1) - give_amt(5,3)
		                        y_re_pay = re_pay(5,1) - re_pay(5,3)
		                        y_give_tot = give_tot(5,1) - give_tot(5,3)
                      %>    
                            <tr>
								<td class="first" rowspan="5" style=" border-top:2px solid #515254;">전월대비증가</td>
                                <td style=" border-top:2px solid #515254;">인원</td>
								<td colspan="3" class="right" style=" border-top:2px solid #515254;"><%=formatnumber(b_pay_count,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;">&nbsp;</td>
							</tr>
                        	<tr>
								<td style=" border-left:1px solid #e3e3e3;">야특근</td>
								<td colspan="3" class="right"><%=formatnumber(b_overtime_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td style=" border-left:1px solid #e3e3e3;">급여</td>
								<td colspan="3" class="right"><%=formatnumber(b_give_amt,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td style=" border-left:1px solid #e3e3e3;">소급</td>
								<td colspan="3" class="right"><%=formatnumber(b_re_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<th style=" border-left:1px solid #e3e3e3;">가감액</th>
								<th colspan="3" class="right"><%=formatnumber(b_give_tot,0)%>&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>                
                            <tr>
								<td class="first" rowspan="5">전년대비증가</td>
                                <td style=" border-left:1px solid #e3e3e3;">인원</td>
								<td colspan="3" class="right"><%=formatnumber(y_pay_count,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                        	<tr>
								<td style=" border-left:1px solid #e3e3e3;">야특근</td>
								<td colspan="3" class="right"><%=formatnumber(y_overtime_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td style=" border-left:1px solid #e3e3e3;">급여</td>
								<td colspan="3" class="right"><%=formatnumber(y_give_amt,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td style=" border-left:1px solid #e3e3e3;">소급</td>
								<td colspan="3" class="right"><%=formatnumber(y_re_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<th style=" border-left:1px solid #e3e3e3;">가감액</th>
								<th colspan="3" class="right"><%=formatnumber(y_give_tot,0)%>&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>                
						</tbody>
					</table>
                    </DIV>
					</td>
                  </tr>
				</table>
                <table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
	</body>
</html>

