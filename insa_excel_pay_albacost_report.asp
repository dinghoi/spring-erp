<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

view_condi=Request("view_condi")
to_date=request("to_date")
pmg_yymm=request("pmg_yymm")
	
curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = pmg_yymm + "월 사업소득 지급현황 -- " + view_condi + ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename	
	
	
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

give_date = to_date '지급일

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_alb = Server.CreateObject("ADODB.Recordset")
Set Rs_year = Server.CreateObject("ADODB.Recordset")
Set Rs_alco = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


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

Sql = "select * from pay_alba_cost where (rever_yymm = '"+pmg_yymm+"' ) and (company = '"+view_condi+"') ORDER BY company,draft_no ASC"

Rs.Open Sql, Dbconn, 1

curr_yyyy = mid(cstr(pmg_yymm),1,4)
curr_mm = mid(cstr(pmg_yymm),5,2)
title_line = cstr(curr_yyyy) + "년 " + cstr(curr_mm) + "월 " + " 사업소득 지급현황 "

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
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
				               <th rowspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">성명</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">등록일</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">구분</th>
				               <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#FFFFE6;">사업소득 및 제수당</th>
                               <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3; background:#E0FFFF;">공제</th>
                               <th rowspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">차인지급액</th>
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
                                <td align="right"><%=formatnumber(rs("alba_pay"),0)%></td>
                                <td align="right"><%=formatnumber(rs("alba_trans"),0)%></td>
                                <td align="right"><%=formatnumber(rs("alba_meals"),0)%></td>
                                <td align="right"><%=formatnumber(rs("alba_other"),0)%></td>
                                <td align="right"><%=formatnumber(rs("alba_give_total"),0)%></td>
                         <%
							  sub_de_hap = int(rs("tax_amt1")) + int(rs("tax_amt2")) + int(rs("de_other"))
							  'pay_amount = alba_give_total - sub_de_hap
							  pay_amount = rs("pay_amount")

                         %>
                                <td align="right"><%=formatnumber(rs("tax_amt1"),0)%></td>
                                <td align="right"><%=formatnumber(rs("tax_amt2"),0)%></td>
                                <td align="right"><%=formatnumber(rs("de_other"),0)%></td>
                                <td align="right"><%=formatnumber(sub_de_hap,0)%></td>
                                <td align="right"><%=formatnumber(pay_amount,0)%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						sum_curr_pay = sum_alba_give_total - sum_deduct_tot
						
						%>
                          	<tr>
                                <th class="first">총계</th>
                                <th colspan="2" align="right"><%=formatnumber(pay_count,0)%>&nbsp;명</th>
                                <th align="right"><%=formatnumber(sum_alba_pay,0)%></th>
                                <th align="right"><%=formatnumber(sum_alba_trans,0)%></th>
                                <th align="right"><%=formatnumber(sum_alba_meals,0)%></th>
                                <th align="right"><%=formatnumber(sum_alba_other,0)%></th>
                                <th align="right"><%=formatnumber(sum_alba_give_total,0)%></th>
                                <th align="right"><%=formatnumber(sum_tax_amt1,0)%></th>
                                <th align="right"><%=formatnumber(sum_tax_amt2,0)%></th>
                                <th align="right"><%=formatnumber(sum_de_other,0)%></th>
                                <th align="right"><%=formatnumber(sum_deduct_tot,0)%></th>
                                <th align="right"><%=formatnumber(sum_pay_amount,0)%></th>
                                <th align="right">&nbsp;</th>
							</tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

