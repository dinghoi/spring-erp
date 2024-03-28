<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim year_tab(5)
dim sum_amt(12)
dim tot_amt(12)
dim cost_tab
cost_tab = array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

cost_month=Request.form("cost_month")
if cost_grade = "0" then
	saupbu=Request.form("saupbu")
end if
if cost_month = "" then
	be_date = dateadd("m",-1,now())
	be_month = mid(cstr(be_date),1,4) + mid(cstr(be_date),6,2)
	cost_month = be_month
'	saupbu = saupbu
end If

sql="select * from emp_org_mst where org_company = '케이원정보통신' and org_level='사업부' and org_empno ='"&emp_no&"'"
set rs_etc=dbconn.execute(sql)
if rs_etc.eof or rs_etc.bof then
	saupbu = saupbu
  else
  	saupbu = rs_etc("org_saupbu")
end if
rs_etc.close()

'if saupbu = "" then
'	saupbu = "KAL지원사업부"
'end if

Sql="select * from cost_end where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
Set rs=DbConn.Execute(Sql)
if rs.eof or rs.bof then
	bonbu_yn = "N"
	batch_yn = "N"
  else
	bonbu_yn = rs("bonbu_yn")
	batch_yn = rs("batch_yn")
end if
rs.close()

cost_year = mid(cost_month,1,4)
mm = int(mid(cost_month,5,2))

for i = 0 to 12
	sum_amt(i) = 0
	tot_amt(i) = 0
next

'sql_org="select * from emp_org_mst where org_level='사업부' and org_saupbu ='"&saupbu&"'"
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
'	rec_cnt = rec_cnt + 1
'	rs_org.movenext()
'loop
'rs_org.close()

title_line = "사업부별 비용 총괄 현황 "

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
				if (document.frm.cost_month.value == "") {
					alert ("조회년월을 입력하세요.");
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="saupbu_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
							<label>
							&nbsp;&nbsp;<strong>조회년월&nbsp;</strong> : 
                            <input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
							</label>
							<label>
							&nbsp;&nbsp;<strong>사업부&nbsp;</strong> : 
							</label>
						<% if cost_grade < "2" then	%>
						<%   if cost_grade = "0" then
                             	  Sql="select org_name from emp_org_mst where org_level = '사업부' group by org_name order by org_name asc"
                                  rs_org.Open Sql, Dbconn, 1
							    else
                             	  Sql="select org_name from emp_org_mst where org_level = '사업부' and org_bonbu = '"&bonbu&"' group by org_name order by org_name asc"
                                  rs_org.Open Sql, Dbconn, 1
							end if
	                    %>
							<label>
                            <select name="saupbu" id="saupbu_view" style="width:150px">
                              <option value="" <%If saupbu = "" then %>selected<% end if %>>직할팀</option>
                              <%
                                do until rs_org.eof
                                %>
                              <option value='<%=rs_org("org_name")%>' <%If saupbu = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                              <%
                                    rs_org.movenext()
                                loop
                                rs_org.close()						
                                %>
                            </select>
							</label>
						<%   else	%>
							<%=saupbu%>
                        <% end if	%>
                            <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div  style="text-align:right">
				<strong>금액단위 : 천원</strong>
				</div>
					<table cellpadding="0" cellspacing="0">
					<tr>
                    	<td>
      					<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
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
						<% if (position = "사업부장" or position = "본부장") and mm = 1 then	%>
						<%   if batch_yn = "N" then	%>   
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
						<% if (position = "사업부장" or position = "본부장") and mm = i then	%>
						<%   if batch_yn = "N" then	%>   
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
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="6%" >
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
						<tbody>

					<%
					for jj = 0 to 10
						rec_cnt = 0
						sql = "select cost_detail from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_tab(jj)&"' group by cost_detail"
'Response.write sql&"; <br>"						
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
						<% if jj = 2 then	%>
                        	  <%=cost_tab(jj)%><br>(현금사용)
						<%   elseif jj = 3 then	%>
                        	  <%=cost_tab(jj)%><br>(주유카드<br>현금사용)
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
						<%    if rs("cost_detail") = "4대보험" or rs("cost_detail") = "소득세종업원분" or rs("cost_detail") = "연차수당" or rs("cost_detail") = "퇴직충당금"  then	%>
								<%=formatnumber(cost_amt,0)%>
						<%		else	%>		
                                <a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a>
						<%	  end if	%>
                                </td>
						<%   else	%>
								<td class="right">
						<%    if rs("cost_detail") = "4대보험" or rs("cost_detail") = "소득세종업원분" or rs("cost_detail") = "연차수당" or rs("cost_detail") = "퇴직충당금"  then	%>
								<%=formatnumber(cost_amt,0)%>
						<%		else	%>		
                                <a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a>
						<%	  end if	%>
                                </td>
                        <% end if	%>
						<%	
							next	
							cr_cost = curr_cost - be_cost							
							
							if cr_cost = 0 then
								cr_pro = 0
							  elseif be_cost = 0 then
							  	cr_pro = 100 
							  else
							  	cr_pro = cr_cost / be_cost * 100
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
								<td class="right" bgcolor="#FFFFCC">
						<%    if rs("cost_detail") = "4대보험" or rs("cost_detail") = "소득세종업원분" or rs("cost_detail") = "연차수당" or rs("cost_detail") = "퇴직충당금"  then	%>
								<%=formatnumber(cost_amt,0)%>
						<%		else	%>		
                                <a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a>
						<%	  end if	%>
                                </td>
						<%   else	%>
								<td class="right">
						<%    if rs("cost_detail") = "4대보험" or rs("cost_detail") = "소득세종업원분" or rs("cost_detail") = "연차수당" or rs("cost_detail") = "퇴직충당금"  then	%>
								<%=formatnumber(cost_amt,0)%>
						<%		else	%>		
                                <a href="#" onClick="pop_Window('team_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&saupbu=<%=saupbu%>','team_cost_view_pop','scrollbars=yes,width=450,height=450')"><%=formatnumber(cost_amt,0)%></a>
						<%	  end if	%>
                                </td>
                        <% end if	%>
						<%	
							next
							cr_cost = curr_cost - be_cost							
							if cr_cost = 0 then
								cr_pro = 0
							  elseif be_cost = 0 then
							  	cr_pro = 100 
							  else
							  	cr_pro = cr_cost / be_cost * 100
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
								cr_pro = cr_cost / sum_amt(mm-1) * 100
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
							<tr bgcolor="#FFE8E8">
							  <td colspan="2" class="first" scope="col">합계</td>
						<%
' 합계
						for i = 1 to 12
						%>
							  <td scope="col" class="right"><%=formatnumber(tot_amt(i)/1000,0)%></td>
						<%
                        next
						cr_cost = tot_amt(mm) - tot_amt(mm-1)
						if cr_cost = 0 then
							cr_pro = 0
						  elseif tot_amt(mm-1) = 0 then
							cr_pro = 100
						  else
							cr_pro = cr_cost / tot_amt(mm-1) * 100
						end if
						%>
							  <td scope="col" class="right"><%=formatnumber(tot_amt(0)/1000,0)%></td>
							  <td class="right"><%=formatnumber(cr_cost/1000,0)%></td>
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

