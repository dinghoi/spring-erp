<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim year_tab(5)
dim sum_amt(13)
dim tot_amt(13)
dim cost_tab
cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

cost_year=Request("cost_year")
view_sw=Request("view_sw")
reside=Request("reside")
common=Request("common")
direct=Request("direct")

for i = 0 to 13
	sum_amt(i) = 0
	tot_amt(i) = 0
next

savefilename = cost_year + "년 비용유형별.xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// 엑셀로 지정
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
	<style type="text/css">
    <!--
    	.style10 {font-size: 10px; font-family: "굴림체", "굴림체", Seoul; }
        .style10B {font-size: 10px; font-weight: bold; font-family: "굴림체", "굴림체", Seoul; }
    -->
    </style>
	</head>
	<body>
		<div id="wrap">			
			<div id="container">
		<%
			if view_sw = "0" then
            	Sql="select cost_year from company_cost where cost_year = '"&cost_year&"' group by cost_year order by cost_year asc"
			end if
			if view_sw = "1" then
            	Sql="select company from company_cost where (cost_center = '상주직접비') and cost_year = '"&cost_year&"' group by company order by company asc"
			end if
			if view_sw = "2" then
            	Sql="select company from company_cost where cost_year = '"&cost_year&"' and cost_center = '"&common&"' group by company order by company asc"
			end if
			if view_sw = "3" then
                Sql="select saupbu from company_cost where cost_center = '직접비' and cost_year = '"&cost_year&"' group by saupbu order by saupbu asc"
			end if
			if view_sw = "4" then
                Sql="select saupbu from company_cost where cost_center = '상주직접비' and cost_year = '"&cost_year&"' group by saupbu order by saupbu asc"
			end if
			rs_org.Open Sql, Dbconn, 1
			do until rs_org.eof
				if view_sw = "0" then
					condi_sql = " "
					title_line = cost_year + "년 전체 비용 현황"
				end if
				if view_sw = "1" then
					condi_sql = " and cost_center = '상주직접비' and company = '"&rs_org("company")&"'"
					if rs_org("company") = "" then						
						title_line = cost_year + "년 " + " 상주처가 없는 상주처 비용 현황"
					  else
						title_line = cost_year + "년 " + rs_org("company") + " 상주처 비용 현황"
					end if
				end if
				if view_sw = "2" then
					condi_sql = " and cost_center = '"&common&"'"
					title_line = cost_year + "년 " + common + " 비용 현황"
				end if
				if view_sw = "3" then
					condi_sql = " and cost_center = '직접비' and saupbu = '"&rs_org("saupbu")&"'"
					title_line = cost_year + "년 " + rs_org("saupbu") + " 직접비 비용 현황"
				end if
				if view_sw = "4" then
					condi_sql = " and cost_center = '상주직접비' and saupbu = '"&rs_org("saupbu")&"'"
					title_line = cost_year + "년 " + rs_org("saupbu") + " 상주직접비 비용 현황"
				end if
        %>
				<h3 class="tit"><%=title_line%></h3>
                <div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
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
						<tbody>
					<%
					for jj = 0 to 10
						rec_cnt = 0
						sql = "select cost_detail from company_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail"
						rs.Open sql, Dbconn, 1
						do until rs.eof
							rec_cnt = rec_cnt + 1
							rs.movenext()
						loop
						rs.close()

						if rec_cnt <> 0 then
							if cost_tab(jj) = "인건비" then
								sql = "select company_cost.cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost inner join saupbu_cost_account on company_cost.cost_id = saupbu_cost_account.cost_id and company_cost.cost_detail = saupbu_cost_account.cost_detail where company_cost.cost_year ='"&cost_year&"' and company_cost.cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by company_cost.cost_detail order by saupbu_cost_account.view_seq"
							  else
								sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
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
							sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost where cost_year ='"&be_year&"' and cost_detail ='"&rs("cost_detail")&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
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
						<% if jj = 2 or jj = 3 then	%>
                        	  <%=cost_tab(jj)%><br>(현금사용)
						<%   else	%>
                        	  <%=cost_tab(jj)%>
                        <% end if	%>
                              </td>
								<td class="left"><%=rs("cost_detail")%></td>
								<td class="right" bgcolor="#FFFFCC"><%=formatnumber(be_cost_amt,0)%></td>
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
									cost_amt = cdbl(cost_amt)
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
                                <td class="right"><%=formatnumber(tot_cost_amt,0)%></td>
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
								sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost where cost_year ='"&be_year&"' and cost_detail ='"&rs("cost_detail")&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
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
								<td class="right" bgcolor="#FFFFCC"><%=formatnumber(be_cost_amt,0)%></td>
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
									cost_amt = cdbl(cost_amt)
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
								<td class="right"><%=formatnumber(tot_cost_amt,0)%></td>
								<td class="right"><%=formatnumber(cr_pro,2)%>%</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(13),0)%></td>
						<%
							for i = 1 to 12
						%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(i),0)%></td>
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
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(0),0)%></td>
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
							  <td class="right"><%=formatnumber(tot_amt(13),0)%></td>
						<%
' 합계
						for i = 1 to 12
						%>
							  <td scope="col" class="right"><%=formatnumber(tot_amt(i),0)%></td>
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
							  <td scope="col" class="right"><%=formatnumber(tot_amt(0),0)%></td>
							  <td class="right"><%=formatnumber(cr_pro,2)%>%</td>
                          </tr>
						</tbody>
					</table>				
			</div>				
		<%
               	rs_org.movenext()
            loop
            rs_org.close()						
'			end if
        %>
	</div>        				
	</body>
</html>

