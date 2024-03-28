<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim month_tab(24,2)

dim com_tab(6)
dim pay_count(6,3)
dim overtime_pay(6,3)
dim give_amt(6,3)
dim re_pay(6,3)
dim give_tot(6,3)

view_condi=Request("view_condi")
pmg_yymm=request("pmg_yymm")
to_date=request("to_date")

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

savefilename = pmg_yymm + "월 급여 전월비교분석.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

if view_condi = "" then
	view_condi = "전체"
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	
  for i = 1 to 6
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
		  com_tab(5) = "코리아디엔씨"
		  com_tab(6) = "합계"
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1')"
	else	  
		  com_tab(1) = view_condi
		  com_tab(6) = "합계"
		  Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof
      for i = 1 to 6
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,1) = pay_count(i,1) + 1
				 pay_count(6,1) = pay_count(6,1) + 1
		         overtime_pay(i,1) = overtime_pay(i,1) + int(rs("pmg_overtime_pay"))
				 overtime_pay(6,1) = overtime_pay(6,1) + int(rs("pmg_overtime_pay"))
		         give_amt(i,1) = give_amt(i,1) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(6,1) = give_amt(6,1) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,1) = re_pay(i,1) + int(rs("pmg_re_pay"))
				 re_pay(6,1) = re_pay(6,1) + int(rs("pmg_re_pay"))
		         give_tot(i,1) = give_tot(i,1) + int(rs("pmg_give_total"))
				 give_tot(6,1) = give_tot(6,1) + int(rs("pmg_give_total"))
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
      for i = 1 to 6
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,2) = pay_count(i,2) + 1
				 pay_count(6,2) = pay_count(6,2) + 1
		         overtime_pay(i,2) = overtime_pay(i,2) + int(rs("pmg_overtime_pay"))
				 overtime_pay(6,2) = overtime_pay(6,2) + int(rs("pmg_overtime_pay"))
		         give_amt(i,2) = give_amt(i,2) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(6,2) = give_amt(6,2) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,2) = re_pay(i,2) + int(rs("pmg_re_pay"))
				 re_pay(6,2) = re_pay(6,2) + int(rs("pmg_re_pay"))
		         give_tot(i,2) = give_tot(i,2) + int(rs("pmg_give_total"))
				 give_tot(6,2) = give_tot(6,2) + int(rs("pmg_give_total"))
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
      for i = 1 to 6
	      if com_tab(i) = rs("pmg_company") then
	             pay_count(i,3) = pay_count(i,3) + 1
				 pay_count(6,3) = pay_count(6,3) + 1
		         overtime_pay(i,3) = overtime_pay(i,3) + int(rs("pmg_overtime_pay"))
				 overtime_pay(6,3) = overtime_pay(6,3) + int(rs("pmg_overtime_pay"))
		         give_amt(i,3) = give_amt(i,3) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
				 give_amt(6,3) = give_amt(6,3) + (int(rs("pmg_give_total")) - int(rs("pmg_overtime_pay")) - int(rs("pmg_re_pay")))
		         re_pay(i,3) = re_pay(i,3) + int(rs("pmg_re_pay"))
				 re_pay(6,3) = re_pay(6,3) + int(rs("pmg_re_pay"))
		         give_tot(i,3) = give_tot(i,3) + int(rs("pmg_give_total"))
				 give_tot(6,3) = give_tot(6,3) + int(rs("pmg_give_total"))
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
		<title>급여관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th colspan="2" class="first" scope="col">구&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;분</th>
								<th scope="col"><%=mid(pmg_yymm,1,4)%>년&nbsp;<%=mid(pmg_yymm,5,2)%>월</th>
                                <th scope="col"><%=mid(bef_month,1,4)%>년&nbsp;<%=mid(bef_month,5,2)%>월</th>
                                <th scope="col"><%=mid(bef_yearmon,1,4)%>년&nbsp;<%=mid(bef_yearmon,5,2)%>월</th>
                                <th scope="col">비고</th>
							</tr>  
                        </thead>
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
						
                        for i = 1 to 6 
                        	if	com_tab(i) <> "" then
						%>	
							<tr>
								<td class="first" rowspan="5"><%=com_tab(i)%></td>
                                <td>인원</td>
								<td align="right"><%=formatnumber(pay_count(i,1),0)%>&nbsp;</td>
								<td align="right"><%=formatnumber(pay_count(i,2),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(pay_count(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                        	<tr>
								<td >야특근</td>
								<td align="right"><%=formatnumber(overtime_pay(i,1),0)%>&nbsp;</td>
								<td align="right"><%=formatnumber(overtime_pay(i,2),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(overtime_pay(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>   
                            <tr>
								<td>급여</td>
								<td align="right"><%=formatnumber(give_amt(i,1),0)%>&nbsp;</td>
								<td align="right"><%=formatnumber(give_amt(i,2),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(give_amt(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>   
                            <tr>
								<td >소급</td>
								<td align="right"><%=formatnumber(re_pay(i,1),0)%>&nbsp;</td>
								<td align="right"><%=formatnumber(re_pay(i,2),0)%>&nbsp;</td>
                                <td align="right"><%=formatnumber(re_pay(i,3),0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>    
                            <tr>
								<th >합계</th>
								<th align="right"><%=formatnumber(give_tot(i,1),0)%>&nbsp;</th>
								<th align="right"><%=formatnumber(give_tot(i,2),0)%>&nbsp;</th>
                                <th align="right"><%=formatnumber(give_tot(i,3),0)%>&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>    
                        <%
							end if
						next
						        b_pay_count = pay_count(6,1) - pay_count(6,2)
		                        b_overtime_pay = overtime_pay(6,1) - overtime_pay(6,2)
		                        b_give_amt = give_amt(6,1) - give_amt(6,2)
		                        b_re_pay = re_pay(6,1) - re_pay(6,2)
		                        b_give_tot = give_tot(6,1) - give_tot(6,2)
								
								y_pay_count = pay_count(6,1) - pay_count(6,3)
		                        y_overtime_pay = overtime_pay(6,1) - overtime_pay(6,3)
		                        y_give_amt = give_amt(6,1) - give_amt(6,3)
		                        y_re_pay = re_pay(6,1) - re_pay(6,3)
		                        y_give_tot = give_tot(6,1) - give_tot(6,3)
                      %>    
                            <tr>
								<td class="first" rowspan="5" style=" border-top:2px solid #515254;">전월대비증가</td>
                                <td style=" border-top:2px solid #515254;">인원</td>
								<td colspan="3" align="right" style=" border-top:2px solid #515254;"><%=formatnumber(b_pay_count,0)%>&nbsp;</td>
                                <td class="right" style=" border-top:2px solid #515254;">&nbsp;</td>
							</tr>
                        	<tr>
								<td>야특근</td>
								<td colspan="3" align="right"><%=formatnumber(b_overtime_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td>급여</td>
								<td colspan="3" align="right"><%=formatnumber(b_give_amt,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td>소급</td>
								<td colspan="3" align="right"><%=formatnumber(b_re_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<th >가감액</th>
								<th colspan="3" align="right"><%=formatnumber(b_give_tot,0)%>&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>                
                            <tr>
								<td class="first" rowspan="5">전년대비증가</td>
                                <td >인원</td>
								<td colspan="3" align="right"><%=formatnumber(y_pay_count,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>
                        	<tr>
								<td >야특근</td>
								<td colspan="3" align="right"><%=formatnumber(y_overtime_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td>급여</td>
								<td colspan="3" align="right"><%=formatnumber(y_give_amt,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<td>소급</td>
								<td colspan="3" align="right"><%=formatnumber(y_re_pay,0)%>&nbsp;</td>
                                <td class="right">&nbsp;</td>
							</tr>  
                            <tr>
								<th>가감액</th>
								<th colspan="3" align="right"><%=formatnumber(y_give_tot,0)%>&nbsp;</th>
                                <th class="right">&nbsp;</th>
							</tr>                
						</tbody>
					</table>
            </div>	
	 	 </div>				
	  </div>        				
	</body>
</html>

