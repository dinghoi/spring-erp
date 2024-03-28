<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim year_tab(5)
dim sum_amt(12)
dim tot_amt(12)
dim cost_tab
cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","전산")

cost_month=Request("cost_month")
saupbu=Request("saupbu")

if cost_month = "" then
	be_date = dateadd("m",-1,now())
	be_month = mid(cstr(be_date),1,4) + mid(cstr(be_date),6,2)
	cost_month = be_month
end If

cost_year = mid(cost_month,1,4)
mm = int(mid(cost_month,5,2))

for i = 0 to 12
	sum_amt(i) = 0
	tot_amt(i) = 0
next

sql = "select * from saupbu_memo where cost_month ='"&cost_month&"' and saupbu ='"&saupbu&"'"
set rs_etc = dbconn.execute(sql)
if rs_etc.eof or rs_etc.bof then
	saupbu_memo = "의견 없음"
	reg_name = ""
	reg_date = ""
  else
  	saupbu_memo = rs_etc("saupbu_memo")
	reg_name = rs_etc("saupbu_reg_name")
	reg_date = rs_etc("saupbu_reg_date")
end if

sql = "select * from cost_end where end_month ='"&cost_month&"' and saupbu ='"&saupbu&"'"
set rs_etc = dbconn.execute(sql)
if rs_etc.eof or rs_etc.bof then
	ceo_yn = "Y"
	bonbu_yn = "Y"
  else
  	ceo_yn = rs_etc("ceo_yn")
  	bonbu_yn = rs_etc("bonbu_yn")
end if

title_line = cost_year + "년" + cstr(mm) + "월 " + saupbu + " 비용 사용 현황 "

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
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("조회년월을 입력하세요.");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_approval_ok.asp" method="post" name="frm">
				<div  style="text-align:right">
				<strong>금액단위 : 천원</strong>
				</div>
                <div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="*" >
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
							<col width="5%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">비용항목</th>
							  <th rowspan="2" scope="col">세부내역</th>
						<% for i = 1 to 12	%>
							  <th scope="col"><%=i%>월</th>
						<% next	%>
							  <th rowspan="2" scope="col">합계</th>
							  <th rowspan="2" scope="col">전월대비</th>
							  <th rowspan="2" scope="col">증감율</th>
                          </tr>
							<tr>
							  <td scope="col" style=" border-left:1px solid #e3e3e3;">
						<% if (position = "본부장" or (user_id = "100031" and saupbu = "KAL지원사업부") or (user_id = "100031" and saupbu = "공항지원사업부")) and mm = 1 then	%>
						<%   if bonbu_yn = "N" then	%>   
                              <a href="#" onClick="pop_Window('saupbu_memo_add.asp?cost_year=<%=cost_year%>&cost_mm=<%=1%>&memo_sw=<%="등록"%>&saupbu=<%=saupbu%>','saupbu_memo_add_pop','scrollbars=yes,width=800,height=600')">의견등록</a>
						<%     else	%>
                              <a href="#" onClick="pop_Window('saupbu_memo_add.asp?cost_year=<%=cost_year%>&cost_mm=<%=1%>&memo_sw=<%="조회"%>&saupbu=<%=saupbu%>','saupbu_memo_add_pop','scrollbars=yes,width=800,height=400')">의견조회</a>
                        <%   end if	%>
						<%   else	%>
                              <a href="#" onClick="pop_Window('saupbu_memo_add.asp?cost_year=<%=cost_year%>&cost_mm=<%=1%>&memo_sw=<%="조회"%>&saupbu=<%=saupbu%>','saupbu_memo_add_pop','scrollbars=yes,width=800,height=400')">의견조회</a>
						<% end if	%>
                              </td>
						<% for i = 2 to 12	%>
							  <td scope="col">
						<% if (position = "본부장" or (user_id = "100031" and saupbu = "KAL지원사업부") or (user_id = "100031" and saupbu = "공항지원사업부")) and mm = i then	%>
						<%   if bonbu_yn = "N" then	%>   
                              <a href="#" onClick="pop_Window('saupbu_memo_add.asp?cost_year=<%=cost_year%>&cost_mm=<%=i%>&memo_sw=<%="등록"%>&saupbu=<%=saupbu%>','saupbu_memo_add_pop','scrollbars=yes,width=800,height=600')">의견등록</a>
						<%     else	%>
                              <a href="#" onClick="pop_Window('saupbu_memo_add.asp?cost_year=<%=cost_year%>&cost_mm=<%=i%>&memo_sw=<%="조회"%>&saupbu=<%=saupbu%>','saupbu_memo_add_pop','scrollbars=yes,width=800,height=400')">의견조회</a>
                        <%   end if	%>
						<%   else	%>
                              <a href="#" onClick="pop_Window('saupbu_memo_add.asp?cost_year=<%=cost_year%>&cost_mm=<%=i%>&memo_sw=<%="조회"%>&saupbu=<%=saupbu%>','saupbu_memo_add_pop','scrollbars=yes,width=800,height=400')">의견조회</a>
						<% end if	%>
                              </td>
						<% next	%>
						  </tr>
						</thead>
						<tbody>
					<%
					for jj = 0 to 10
						rec_cnt = 0
						sql = "select cost_detail from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_tab(jj)&"' group by cost_detail"
						rs.Open sql, Dbconn, 1
						do until rs.eof
							rec_cnt = rec_cnt + 1
							rs.movenext()
						loop
						rs.close()

						if rec_cnt <> 0 then
							if cost_tab(jj) = "인건비" then
								sql = "select org_cost.cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost inner join saupbu_cost_account on org_cost.cost_id = saupbu_cost_account.cost_id and org_cost.cost_detail = saupbu_cost_account.cost_detail where org_cost.cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and org_cost.cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by org_cost.cost_detail order by saupbu_cost_account.view_seq"
							  else
								sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
							end if
'							sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_tab(jj)&"' group by cost_detail order by sort_seq"
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
						<%	
							for k = 1 to 12
								if k < 10 then
									kk = "0" + cstr(k)
								  else
								  	kk = cstr(k)
								end if	
								cost = "cost_amt_" + cstr(kk)
								cost_amt = rs(cost)
								if k = mm -1 then
									be_cost = cdbl(cost_amt)
								end if
								if k = mm then
									curr_cost = cdbl(cost_amt)
								end if								
								if cost_amt = "0" then
									cost_amt = 0
								  else
									cost_amt = cdbl(cost_amt) / 1000
								end if
						%>
						<% if k = mm then	%>
								<td class="right" bgcolor="#FFFFCC">
                                <a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a>
                                </td>
						<%   else	%>
								<td class="right">
                                <a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a>
                                </td>
                        <% end if	%>
						<%	
							next	
							cr_cost = curr_cost - be_cost							
							
							if cr_cost = 0 then
								cr_pro = 0
							  elseif bi_cost = 0 then
							  	cr_pro = 100 
							  else
							  	cr_pro = cr_cost / be_cost
							end if
						%>								
                                <td class="right"><%=formatnumber(tot_cost_amt/1000,0)%></td>
								<td class="right"><%=formatnumber(cr_cost/1000,0)%></td>
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
						%>
                        	<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;"><%=rs("cost_detail")%></td>
						<%	
							for k = 1 to 12
								if k < 10 then
									kk = "0" + cstr(k)
								  else
								  	kk = cstr(k)
								end if	
								cost = "cost_amt_" + cstr(kk)
								cost_amt = rs(cost)
								if k = mm -1 then
									be_cost = cdbl(cost_amt)
								end if
								if k = mm then
									curr_cost = cdbl(cost_amt)
								end if								
								if cost_amt = "0" then
									cost_amt = 0
								  else
									cost_amt = cdbl(cost_amt) / 1000
								end if
						%>
						<% if k = mm then	%>
								<td class="right" bgcolor="#FFFFCC"><a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a></td>
						<%   else	%>
								<td class="right"><a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a></td>
						<% end if	%>
						<%	
							next
							cr_cost = curr_cost - be_cost							
							if cr_cost = 0 then
								cr_pro = 0
							  elseif bi_cost = 0 then
							  	cr_pro = 100 
							  else
							  	cr_pro = cr_cost / be_cost
							end if
						%>
								<td class="right"><%=formatnumber(tot_cost_amt/1000,0)%></td>
								<td class="right"><%=formatnumber(cr_cost/1000,0)%></td>
								<td class="right"><%=formatnumber(cr_pro,2)%>%</td>
							</tr>
						<%
								rs.movenext()
							loop
							rs.close()
						%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
						<%
							for i = 1 to 12
						%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(i)/1000,0)%></td>
						<%
							next
							cr_cost = sum_amt(mm) - sum_amt(mm-1)
							if cr_cost = 0 then
								cr_pro = 0
							  elseif sum_amt(mm-1) = 0 then
								cr_pro = 100
							  else
								cr_pro = cr_cost / sum_amt(mm-1)
							end if
						%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(0)/1000,0)%></td>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(cr_cost/1000,0)%></td>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(cr_pro,2)%>%</td>
							</tr>
						<%
							for i = 0 to 12
								tot_amt(i) = tot_amt(i) + sum_amt(i)
								sum_amt(i) = 0
							next
						end if
					next
					%>
							<tr>
							  <th colspan="2" class="first" scope="col">합계</th>
						<%
' 합계
						for i = 1 to 12
						%>
							  <th scope="col" class="right"><%=formatnumber(tot_amt(i)/1000,0)%></th>
						<%
                        next
						cr_cost = tot_amt(mm) - tot_amt(mm-1)
						if cr_cost = 0 then
							cr_pro = 0
						  elseif tot_amt(mm-1) = 0 then
							cr_pro = 100
						  else
							cr_pro = cr_cost / tot_amt(mm-1)
						end if
						%>
							  <th scope="col" class="right"><%=formatnumber(tot_amt(0)/1000,0)%></th>
							  <th class="right"><%=formatnumber(cr_cost/1000,0)%></th>
							  <th class="right"><%=formatnumber(cr_pro,2)%>%</th>
                          </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
				<%	if (emp_no = "100001" and ceo_yn = "N") then	%>
                    <span class="btnType01"><input type="button" value="비용승인" onclick="javascript:frmcheck();" NAME="Button1"></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="닫기" onclick="javascript:goAction();"></span>
                </div>
				<br>
				<input type="hidden" name="cost_month" value="<%=cost_month%>" ID="Hidden1">
				<input type="hidden" name="saupbu" value="<%=saupbu%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

