<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Dim Rs
Dim stay_name
dim pay_tab(5)
dim pay_pay(5)
dim bonus_tab(5)

view_condi=Request("view_condi")
to_date=request("to_date")

for i = 1 to 5
	    pay_tab(i) = ""
     	pay_pay(i) = 0
    	bonus_tab(i) = 0
next
sum_retire_pay = 0
curr_date = datevalue(mid(cstr(now()),1,10))

target_date = to_date

t_year = int(mid(cstr(target_date),1,4))
t_month = int(mid(cstr(target_date),6,2))
t_day = int(mid(cstr(target_date),9,2))
tcal_month = mid(cstr(target_date),1,4) + mid(cstr(target_date),6,2)
tcal_day = cstr(t_day)

pay_tab(3) = cstr(tcal_month)
tcal_month = cstr(int(tcal_month) - 1)
pay_tab(2) = cstr(tcal_month)
tcal_month = cstr(int(tcal_month) - 1)
pay_tab(1) = cstr(tcal_month)

tar1_date = cstr(mid(pay_tab(3),1,4) + "-" + mid(pay_tab(3),5,2) + "-" + tcal_day)
fir1_date = cstr(mid(pay_tab(1),1,4) + "-" + mid(pay_tab(1),5,2) + "-" + "01")
work1_cnt = int(datediff("d", fir1_date, tar1_date)) + 1
pay_tab(5) = work1_cnt

savefilename = "퇴직급여 추계액 -- "+ to_date +" "+ view_condi +".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set Rs_stay = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi = "전체" then
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_company,emp_no ASC"
   else  
   Sql = "select * from emp_master where emp_company = '"+view_condi+"' and (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_no < '900000') ORDER BY emp_company,emp_no ASC"
end if
Rs.Open Sql, Dbconn, 1

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
													
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<style type="text/css">
<!--
.style1 {font-size: 12px}
.style2 {
	font-size: 14px;
	font-weight: bold;
}
-->
</style>
</head>
<body>
<table  border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="#EFEFEF" class="style11">
    <td colspan="13" bgcolor="#FFFFFF"><div align="left" class="style2"><%=to_date%> &nbsp;퇴직급여 추계액 현황>&nbsp;(<%=view_condi%>)</div></td>
  </tr>
  <tr bgcolor="#EFEFEF" class="style11">
    <td><div align="center" class="style1">회사</div></td>
    <td><div align="center" class="style1">소속</div></td>
    <td><div align="center" class="style1">사번</div></td>
    <td><div align="center" class="style1">성명</div></td>
    <td><div align="center" class="style1">직급</div></td>
    <td><div align="center" class="style1">직책</div></td>
    <td><div align="center" class="style1">최초입사일</div></td>
    <td><div align="center" class="style1"><%=mid(pay_tab(1),1,4)%>년&nbsp;<%=mid(pay_tab(1),5,2)%>월</div></td>
    <td><div align="center" class="style1"><%=mid(pay_tab(2),1,4)%>년&nbsp;<%=mid(pay_tab(2),5,2)%>월</div></td>
    <td><div align="center" class="style1"><%=mid(pay_tab(3),1,4)%>년&nbsp;<%=mid(pay_tab(3),5,2)%>월</div></td>
    <td><div align="center" class="style1">일수</div></td>
    <td><div align="center" class="style1">평균임금</div></td>
    <td><div align="center" class="style1">월평균임금</div></td>
    <td><div align="center" class="style1">근속연수</div></td>
    <td><div align="center" class="style1">퇴직추계액</div></td>
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
			sum_retire_pay = sum_retire_pay + retire_pay

	%>
  <tr valign="middle" class="style11">
    <td width="145"><div align="center" class="style1"><%=rs("emp_company")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_org_name")%></div></td>
    <td width="115"><div align="center" class="style1"><%=rs("emp_no")%></div></td>
    <td width="145"><div align="center" class="style1"><%=rs("emp_name")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_grade")%></div></td>
    <td width="59"><div align="center" class="style1"><%=rs("emp_position")%></div></td>
    <td width="145"><div align="center" class="style1"><%=emp_first_date%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(pay_pay(1),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(pay_pay(2),0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(pay_pay(3),0)%></div></td>
    <td width="145"><div align="center" class="style1"><%=pay_tab(5)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(eot_average_pay,0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(eot_month_pay,0)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(gunsok_cnt,1)%></div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(retire_pay,0)%></div></td>
    <% 'response.write(rs("emp_stay_code"))
	   'response.End %>
  </tr>
	<%
	Rs.MoveNext()
	loop
	%>
  <tr valign="middle" class="style11">	
    <td colspan="14"><div align="right" class="style1">퇴직추계액 총계</div></td>
    <td width="145"><div align="right" class="style1"><%=formatnumber(sum_retire_pay,0)%></div></td>
  </tr>	  
</table>
</body>
</html>
<%
Rs.Close()
Set Rs = Nothing
%>
