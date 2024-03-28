<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim year_tab(5)
dim sum_amt(13)
dim tot_amt(13)
dim cost_tab
cost_tab = array("인건비","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

cost_year=Request.form("cost_year")
com_yn=Request.form("com_yn")
emp_company=Request.form("emp_company")
bonbu=Request.form("bonbu")

if cost_year = "" then
	cost_year = mid(cstr(now()),1,4)
	base_year = cost_year
	com_yn = "Y"
	emp_company = "전체"
	bonbu = ""
end If

be_year = int(cost_year) - 1
for i = 1 to 5
	year_tab(i) = int(cost_year) - i + 1
next

'if bonbu = "" then
'	sql="select * from emp_org_mst where org_company = '케이원정보통신' and org_level='본부' order by  org_name asc"
'	rs_org.Open Sql, Dbconn, 1
'	bonbu = rs_org("org_name")
'	rs_org.close()
'end if

for i = 0 to 13
	sum_amt(i) = 0
	tot_amt(i) = 0
next

emp_company_view = ""
bonbu_view = ""
if com_yn = "Y" then
	if emp_company = "전체" then
		cond_sql = ""
		emp_company_view = ""
	  else	
	  	condi_sql = " and emp_company = '" + emp_company + "'"
		emp_company_view = emp_company
	end if
  else
  	condi_sql = " and bonbu = '" + bonbu + "'"
	bonbu_view = bonbu
end if

'sql_org="select * from emp_org_mst where org_level='사업부' and org_company ='케이원정보통신'"
'rs_org.Open sql_org, Dbconn, 1
'do until rs_org.eof

'	sql="select * from saupbu_cost_account"
'	rs.Open sql, Dbconn, 1
'	do until rs.eof
'		sql_etc = "select * from org_cost where cost_year='"&cost_year&"' and emp_company='"&rs_org("org_company")&"' and bonbu='"&rs_org("org_bonbu")&"' and saupbu='"&rs_org("org_saupbu")&"' and org_name='"&rs_org("org_name")&"' and cost_id='"&rs("cost_id")&"' and cost_detail='"&rs("cost_detail")&"'"
'		set rs_etc=dbconn.execute(sql_etc)
'		if rs_etc.eof or rs_etc.bof then
'			sql="insert into org_cost (cost_year,emp_company,bonbu,saupbu,org_name,cost_id,cost_detail) values ('"&cost_year&"','"&rs_org("org_company")&"','"&rs_org("org_bonbu")&"','"&rs_org("org_saupbu")&"','"&rs_org("org_name")&"','"&rs("cost_id")&"','"&rs("cost_detail")&"')"
'			dbconn.execute(sql)
'		end if
'		rs_etc.close()
		
'		rs.movenext()
'	loop
'	rs.close()
''	rec_cnt = rec_cnt + 1
'	rs_org.movenext()
'loop
'rs_org.close()

title_line = "전체 비용 현황"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.cost_year.value == "") {
					alert ("조회년을 입력하세요.");
					return false;
				}	
				return true;
			}
			function condi_view() {

				if (eval("document.frm.com_yn[0].checked")) {
					document.getElementById('emp_company_view').style.display = '';
					document.getElementById('bonbu_view').style.display = 'none';
				}	
				if (eval("document.frm.com_yn[1].checked")) {
					document.getElementById('emp_company_view').style.display = 'none';
					document.getElementById('bonbu_view').style.display = '';
				}	
			}
			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body onload="condi_view();">
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="total_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
						<dd>
							<p>
								<label>
									&nbsp;&nbsp;<strong>조회년&nbsp;</strong> : 
									<select name="cost_year" id="cost_year" style="width:70px">
										<% for i = 1 to 5 %>
										<option value="<%=year_tab(i)%>" <% if cost_year=year_tab(i) then %>selected<% end if %>>&nbsp;<%=year_tab(i)%></option>
										<% next	%>
									</select>
								</label>
								<label>
									<input type="radio" name="com_yn" value="Y" <% if com_yn = "Y" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="condi_view()"><strong>회사별</strong>
									<input type="radio" name="com_yn" value="N" <% if com_yn = "N" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="condi_view()"><strong>본부별</strong>
								</label>
								<label>
									<select name="emp_company" id="emp_company_view" style="width:150px">
										<option value="전체" <% if emp_company = "전체" then %>selected<% end if %>>전체</option>
										<%
                      Sql="select * from org_cost where cost_year ='"&cost_year&"' group by emp_company order by emp_company asc"
                      rs_org.Open Sql, Dbconn, 1
                      do until rs_org.eof
                    %>
                    <option value='<%=rs_org("emp_company")%>' <%If emp_company = rs_org("emp_company") then %>selected<% end if %>><%=rs_org("emp_company")%></option>
                    <%
                    		rs_org.movenext()
                      loop
                      rs_org.close()						
                    %>
                  </select>
								</label>
								<label>
									<select name="bonbu" id="bonbu_view" style="width:150px; display:none">
                  	<option value='' <%If bonbu = "" then %>selected<% end if %>>직할본부</option>
                    <%
                    	Sql="select * from org_cost where bonbu <> '' and cost_year ='"&cost_year&"' group by bonbu order by  bonbu asc"
                      rs_org.Open Sql, Dbconn, 1
                      do until rs_org.eof
                    %>
                    <option value='<%=rs_org("bonbu")%>' <%If bonbu = rs_org("bonbu") then %>selected<% end if %>><%=rs_org("bonbu")%></option>
                    <%
                    		rs_org.movenext()
                      loop
                      rs_org.close()						
                    %>
                  </select>
								</label>
                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
              </p>
						</dd>
					</dl>
				</fieldset>
				<div style="text-align:right"><strong>금액단위 : 천원</strong></div>
				<table cellpadding="0" cellspacing="0">
					<tr>
						<td>
							<DIV id="topLine2" style="width:1200px;overflow:hidden;">
								<div class="gView">
									<table cellpadding="0" cellspacing="0" class="tableList">
										<colgroup>
											<col width="5%" >
											<col width="*" >
											<col width="6%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="6%" >
											<col width="6%" >
										</colgroup>
										<thead>
											<tr>
												<th class="first" scope="col">비용항목</th>
												<th scope="col">세부내역</th>
												<th scope="col">전년</th>
												<% for i = 1 to 12	%>
												<th scope="col"><%=i%>월</th>
												<% next	%>
												<th scope="col">합계</th>
												<th scope="col">전년대비</th>
              				</tr>
              			</thead>
              		</table>
                </DIV>
							</td>
						</tr>
						<tr>
							<td valign="top">
								<DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
									<table cellpadding="0" cellspacing="0" class="scrollList">
										<colgroup>
											<col width="5%" >
											<col width="*" >
											<col width="6%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="5%" >
											<col width="6%" >
											<col width="6%" >
										</colgroup>
										<tbody>
											<%
												for jj = 0 to 9
													rec_cnt = 0
													sql = "select cost_detail from org_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail"
													rs.Open sql, Dbconn, 1
													
													do until rs.eof
														rec_cnt = rec_cnt + 1
														rs.movenext()
													loop
													rs.close()
													
													if rec_cnt <> 0 then
														if cost_tab(jj) = "인건비" then
															sql = "select org_cost.cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost inner join saupbu_cost_account on org_cost.cost_id = saupbu_cost_account.cost_id and org_cost.cost_detail = saupbu_cost_account.cost_detail where org_cost.cost_year ='"&cost_year&"' and org_cost.cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by org_cost.cost_detail order by saupbu_cost_account.view_seq"
							  						else
							  							sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
														end if
														rs.Open sql, Dbconn, 1
														
														tot_cost_amt = cdbl(rs("cost_amt_01")) + cdbl(rs("cost_amt_02")) + cdbl(rs("cost_amt_03")) + cdbl(rs("cost_amt_04")) + cdbl(rs("cost_amt_05")) + cdbl(rs("cost_amt_06")) + cdbl(rs("cost_amt_07")) + cdbl(rs("cost_amt_08")) + cdbl(rs("cost_amt_09")) + cdbl(rs("cost_amt_10")) + cdbl(rs("cost_amt_11")) + cdbl(rs("cost_amt_12")) 
														sum_amt(0) = sum_amt(0) + tot_cost_amt
														sum_amt(1) = sum_amt(1) + cdbl(rs("cost_amt_01"))
														sum_amt(2) = sum_amt(2) + cdbl(rs("cost_amt_02"))
														sum_amt(3) = sum_amt(3) + cdbl(rs("cost_amt_03"))
														sum_amt(4) = sum_amt(4) + cdbl(rs("cost_amt_04"))
														sum_amt(5) = sum_amt(5) + cdbl(rs("cost_amt_05"))
														sum_amt(6) = sum_amt(6) + cdbl(rs("cost_amt_06"))
														sum_amt(7) = sum_amt(7) + cdbl(rs("cost_amt_07"))
														sum_amt(8) = sum_amt(8) + cdbl(rs("cost_amt_08"))
														sum_amt(9) = sum_amt(9) + cdbl(rs("cost_amt_09"))
														sum_amt(10) = sum_amt(10) + cdbl(rs("cost_amt_10"))
														sum_amt(11) = sum_amt(11) + cdbl(rs("cost_amt_11"))
														sum_amt(12) = sum_amt(12) + cdbl(rs("cost_amt_12"))

														' 전년 자료
														sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost where cost_year ='"&be_year&"' and cost_detail ='"&rs("cost_detail")&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
														set rs_etc=DbConn.Execute(sql)
														
														if rs_etc.eof or rs_etc.bof then
															be_cost_amt = 0
														else
															be_cost_amt = cdbl(rs_etc("cost_amt_01")) + cdbl(rs_etc("cost_amt_02")) + cdbl(rs_etc("cost_amt_03")) + cdbl(rs_etc("cost_amt_04")) + cdbl(rs_etc("cost_amt_05")) + cdbl(rs_etc("cost_amt_06")) + cdbl(rs_etc("cost_amt_07")) + cdbl(rs_etc("cost_amt_08")) + cdbl(rs_etc("cost_amt_09")) + cdbl(rs_etc("cost_amt_10")) + cdbl(rs_etc("cost_amt_11")) + cdbl(rs_etc("cost_amt_12")) 
														end if
														
														rs_etc.close()
														sum_amt(13) = sum_amt(13) + be_cost_amt
											%>
											<tr>
												<td rowspan="<%=rec_cnt + 1%>" class="first">
													<% 
														if jj = 1 or jj = 2 then	
															Response.write cost_tab(jj)&"<br>(현금사용)"
														else
															Response.write cost_tab(jj)
														end if
													%>
                        </td>
                        <td class="left"><%=rs("cost_detail")%></td>
                        <td class="right" bgcolor="#FFFFCC"><%=formatnumber(be_cost_amt/1000,0)%></td>
                        <%	
                        	for k = 1 to 12
                        		if k < 10 then
                        			kk = "0" + cstr(k)
                        		else
                        			kk = cstr(k)
														end if
														
														cost = "cost_amt_" + cstr(kk)
														cost_amt = rs(cost)
														
														if cost_amt = "0" then
															cost_amt = 0
								  					else
								  						cost_amt = cdbl(cost_amt) / 1000
														end if
												%>
												<td class="right"><%=formatnumber(cost_amt,0)%></td>
												<%	
													next	
													
													if be_cost_amt = 0 then
														cr_pro = 100
							  					else
							  						cr_pro = tot_cost_amt / be_cost_amt * 100
													end if
													
													if be_cost_amt = 0  and tot_cost_amt = 0 then
														cr_pro = 0
													end if
												%>
												<td class="right"><%=formatnumber(tot_cost_amt/1000,0)%></td>
												<td class="right"><%=formatnumber(cr_pro,2)%>%</td>
											</tr>
											<%
												rs.movenext()
												
												do until rs.eof
													tot_cost_amt = cdbl(rs("cost_amt_01")) + cdbl(rs("cost_amt_02")) + cdbl(rs("cost_amt_03")) + cdbl(rs("cost_amt_04")) + cdbl(rs("cost_amt_05")) + cdbl(rs("cost_amt_06")) + cdbl(rs("cost_amt_07")) + cdbl(rs("cost_amt_08")) + cdbl(rs("cost_amt_09")) + cdbl(rs("cost_amt_10")) + cdbl(rs("cost_amt_11")) + cdbl(rs("cost_amt_12")) 

													sum_amt(0) = sum_amt(0) + tot_cost_amt
													sum_amt(1) = sum_amt(1) + cdbl(rs("cost_amt_01"))
													sum_amt(2) = sum_amt(2) + cdbl(rs("cost_amt_02"))
													sum_amt(3) = sum_amt(3) + cdbl(rs("cost_amt_03"))
													sum_amt(4) = sum_amt(4) + cdbl(rs("cost_amt_04"))
													sum_amt(5) = sum_amt(5) + cdbl(rs("cost_amt_05"))
													sum_amt(6) = sum_amt(6) + cdbl(rs("cost_amt_06"))
													sum_amt(7) = sum_amt(7) + cdbl(rs("cost_amt_07"))
													sum_amt(8) = sum_amt(8) + cdbl(rs("cost_amt_08"))
													sum_amt(9) = sum_amt(9) + cdbl(rs("cost_amt_09"))
													sum_amt(10) = sum_amt(10) + cdbl(rs("cost_amt_10"))
													sum_amt(11) = sum_amt(11) + cdbl(rs("cost_amt_11"))
													sum_amt(12) = sum_amt(12) + cdbl(rs("cost_amt_12"))
													
													' 전년 자료
													sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost where cost_year ='"&be_year&"' and cost_detail ='"&rs("cost_detail")&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
													set rs_etc=DbConn.Execute(sql)
													
													if rs_etc.eof or rs_etc.bof then
														be_cost_amt = 0
								  				else
								  					be_cost_amt = cdbl(rs_etc("cost_amt_01")) + cdbl(rs_etc("cost_amt_02")) + cdbl(rs_etc("cost_amt_03")) + cdbl(rs_etc("cost_amt_04")) + cdbl(rs_etc("cost_amt_05")) + cdbl(rs_etc("cost_amt_06")) + cdbl(rs_etc("cost_amt_07")) + cdbl(rs_etc("cost_amt_08")) + cdbl(rs_etc("cost_amt_09")) + cdbl(rs_etc("cost_amt_10")) + cdbl(rs_etc("cost_amt_11")) + cdbl(rs_etc("cost_amt_12")) 
													end if
													rs_etc.close()
													sum_amt(13) = sum_amt(13) + be_cost_amt
											%>
											<tr>
												<td class="left" style=" border-left:1px solid #e3e3e3;"><%=rs("cost_detail")%></td>
												<td class="right" bgcolor="#FFFFCC"><%=formatnumber(be_cost_amt/1000,0)%></td>
												<%	
													for k = 1 to 12
														if k < 10 then
															kk = "0" + cstr(k)
								  					else
								  						kk = cstr(k)
														end if
														
														cost = "cost_amt_" + cstr(kk)
														cost_amt = rs(cost)
														
														if cost_amt = "0" then
															cost_amt = 0
								  					else
															cost_amt = cdbl(cost_amt) / 1000
														end if
												%>
												<td class="right"><%=formatnumber(cost_amt,0)%></td>
												<%	
													next
													
													if be_cost_amt = 0 then
														cr_pro = 100
							  					else
							  						cr_pro = tot_cost_amt / be_cost_amt * 100
													end if
													
													if be_cost_amt = 0  and tot_cost_amt = 0 then
														cr_pro = 0
														end if
												%>
												<td class="right"><%=formatnumber(tot_cost_amt/1000,0)%></td>
												<td class="right"><%=formatnumber(cr_pro,2)%>%</td>
											</tr>
											<%
													rs.movenext()
												loop
												rs.close()
											%>
											<tr>
												<td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
												<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(13)/1000,0)%></td>
												<%
													for i = 1 to 12
												%>
												<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(i)/1000,0)%></td>
												<%
													next
													
													if sum_amt(13) = 0 then
														cr_pro = 100
							  					else
							  						cr_pro = sum_amt(0) / sum_amt(13) * 100
													end if
													if sum_amt(13) = 0  and sum_amt(0) = 0 then
														cr_pro = 0
													end if
												%>
												<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(0)/1000,0)%></td>
												<td class="right" bgcolor="#EEFFFF"><%=formatnumber(cr_pro,2)%>%</td>
											</tr>
											<%
												for i = 0 to 13
													tot_amt(i) = tot_amt(i) + sum_amt(i)
													sum_amt(i) = 0
												next
											end if
										next
										%>
							<tr bgcolor="#FFDFDF">
							  <td colspan="2" class="first" scope="col">합계</td>
							  <td class="right"><%=formatnumber(tot_amt(13)/1000,0)%></td>
						<%
' 합계
						for i = 1 to 12
						%>
							  <td scope="col" class="right"><%=formatnumber(tot_amt(i)/1000,0)%></td>
						<%
                        next

						if tot_amt(13) = 0 then
							cr_pro = 100 
						  else
						  	cr_pro = tot_amt(0) / tot_amt(13) * 100
						end if
						if tot_amt(13) = 0  and tot_amt(0) = 0 then
							cr_pro = 0
						end if
						%>
							  <td scope="col" class="right"><%=formatnumber(tot_amt(0)/1000,0)%></td>
							  <td class="right"><%=formatnumber(cr_pro,2)%>%</td>
                          </tr>
						</tbody>
					</table>
                        </DIV>
						</td>
                    </tr>
					</table>
				
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

