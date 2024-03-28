<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
Dim from_date
Dim to_date
Dim win_sw
	 
view_condi=Request("view_condi")
from_date=request("from_date")
to_date=request("to_date")
pmg_yymm=request("pmg_yymm")

curr_date = datevalue(mid(cstr(now()),1,10))

savefilename = "야·특근 수당(일반경비 잡비포함) -- "+ view_condi +"(" + from_date + "∼" + to_date + ").xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from pay_overtime_cost where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' ORDER BY emp_company,team,org_name,mg_ce_id,work_date ASC"
Rs.Open Sql, Dbconn, 1
do until rs.eof
    overtime_count = overtime_count + 1
    sum_overtime_pay = sum_overtime_pay + int(rs("overtime_amt"))
	rs.movenext()
loop
rs.close()

Sql = "select * from pay_overtime_cost where emp_company = '"+view_condi+"' and work_date >= '"+from_date+"' and work_date <= '"+to_date+"' ORDER BY emp_company,team,org_name,mg_ce_id,work_date ASC"
Rs.Open Sql, Dbconn, 1
	
title_line = "야·특근 수당(일반경비 잡비포함) -- "+ view_condi +"(" + from_date + "∼" + to_date + ")"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
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
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
						<%
						     sum_overtime_cnt = 0	 
						     sum_overtime_cost = 0
							 
							 tot_overtime_cnt = 0	 
						     tot_overtime_cost = 0
							 
						if rs.eof or rs.bof then
							bi_team = ""
							bi_ce = ""
					      else						  
							if isnull(rs("team")) or rs("team") = "" then	
								bi_team = ""
							  else
								bi_team = rs("team")
							end if
							if isnull(rs("mg_ce_id")) or rs("mg_ce_id") = "" then	
								bi_mg_ce_id = ""
							  else
								bi_mg_ce_id = rs("mg_ce_id")
							end if
						end if

						do until rs.eof

                              if isnull(rs("team")) or rs("team") = "" then
								     emp_team = ""
							     else
							  	     emp_team = rs("team")
							  end if
							  if isnull(rs("mg_ce_id")) or rs("mg_ce_id") = "" then
								     mg_ce_id = ""
							     else
							  	     mg_ce_id = rs("mg_ce_id")
							  end if
							
							  if bi_mg_ce_id <> mg_ce_id then
							     emp_no = bi_mg_ce_id
							     Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                                 Set rs_emp = DbConn.Execute(SQL)
							     if not Rs_emp.eof then
								      ce_name = rs_emp("emp_name")
							     end if
							     rs_emp.close()
					  %>
                                 <tr>
								    <td colspan="4" bgcolor="#EEFFFF" align="center"><%=ce_name%>&nbsp;&nbsp;&nbsp;계</td>
							        <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_overtime_cnt,0)%>&nbsp;건</td>
							        <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_overtime_cost,0)%>&nbsp;원</td>
							        <td colspan="3" bgcolor="#EEFFFF" >&nbsp;</th>
						         </tr>
                      <%
							     sum_overtime_cnt = 0	 
								 sum_overtime_cost = 0
								 bi_mg_ce_id = mg_ce_id
							  end if
							  
							  if bi_team <> emp_team then
					  %>
                                 <tr>
								    <td colspan="4" bgcolor="#EEFFFF" align="center"><%=bi_team%>&nbsp;&nbsp;&nbsp;계</td>
							        <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(tot_overtime_cnt,0)%>&nbsp;건</td>
							        <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(tot_overtime_cost,0)%>&nbsp;원</td>
							        <td colspan="3" bgcolor="#EEFFFF" >&nbsp;</th>
						         </tr>                      
                      
                      <%    
							     tot_overtime_cnt = 0	 
								 tot_overtime_cost = 0
								 bi_team = emp_team
							  end if
							  					                    							  
							  emp_no = rs("mg_ce_id")
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
								   emp_end_date = rs_emp("emp_end_date")
							  end if
							  rs_emp.close()
                              
							  if isNull(emp_end_date) or emp_end_date = "1900-01-01" or emp_end_date = "0000-00-00" then
							          emp_end = ""
								 else 
								      emp_end = "퇴직"
							  end if
							  
							  sum_overtime_cnt = sum_overtime_cnt + 1	 
							  sum_overtime_cost = sum_overtime_cost + int(rs("overtime_amt"))
							  
							  tot_overtime_cnt = tot_overtime_cnt + 1	 
							  tot_overtime_cost = tot_overtime_cost + int(rs("overtime_amt"))

	           			%>
							<tr>
								<td class="left"><%=rs("team")%>-<%=rs("org_name")%></td>
                                <td class="left"><%=rs("cost_detail")%></td>
                                <td class="left"><%=rs("work_date")%>&nbsp;<%=mid(rs("from_time"),1,2)%>:<%=mid(rs("from_time"),3,2)%>∼<%=mid(rs("to_time"),1,2)%>:<%=mid(rs("to_time"),3,2)%></td>
                                <td class="left"><%=rs("company")%></td>
                                <td class="left"><%=rs("dept")%></td>
                                <td><%=emp_name%>(<%=rs("mg_ce_id")%>)</td>
                                <td class="left"><%=rs("sign_no")%></td>
                                <td class="right"><%=formatnumber(rs("overtime_amt"),0)%></td>
                                <td><%=rs("acpt_no")%></td>
                                <td class="left"><%=rs("work_gubun")%>-<%=rs("work_memo")%></td>
                                <td><%=emp_end%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()

						         emp_no = bi_mg_ce_id
							     Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                                 Set rs_emp = DbConn.Execute(SQL)
							     if not Rs_emp.eof then
								      ce_name = rs_emp("emp_name")
							     end if
							     rs_emp.close()
						
						%>
                            <tr>
								<td colspan="4" bgcolor="#EEFFFF" align="center"><%=ce_name%>&nbsp;&nbsp;&nbsp;계</td>
							    <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_overtime_cnt,0)%>&nbsp;건</td>
							    <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(sum_overtime_cost,0)%>&nbsp;원</td>
							    <td colspan="3" bgcolor="#EEFFFF" >&nbsp;</th>
						    </tr>
                            
                            <tr>
								<td colspan="4" bgcolor="#EEFFFF" align="center"><%=bi_team%>&nbsp;&nbsp;&nbsp;계</td>
							    <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(tot_overtime_cnt,0)%>&nbsp;건</td>
							    <td colspan="2" bgcolor="#EEFFFF" align="right"><%=formatnumber(tot_overtime_cost,0)%>&nbsp;원</td>
							    <td colspan="3" bgcolor="#EEFFFF" >&nbsp;</th>
						    </tr>  

                            <tr>
								<th colspan="4" align="center">합&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;계</th>
							    <th colspan="2" align="right"><%=formatnumber(overtime_count,0)%>&nbsp;건</th>
							    <th colspan="2" align="right"><%=formatnumber(sum_overtime_pay,0)%>&nbsp;원</th>
							    <th colspan="3">&nbsp;</th>
						    </tr>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>
