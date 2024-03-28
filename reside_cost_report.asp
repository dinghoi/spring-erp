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

cost_year=Request.form("cost_year")
view_sw=Request.form("view_sw")
reside=Request.form("reside")
common=Request.form("common")
direct=Request.form("direct")

if cost_year = "" then
	cost_year = mid(cstr(now()),1,4)
	base_year = cost_year
	view_sw = "0"
end If

be_year = int(cost_year) - 1
for i = 1 to 5
	year_tab(i) = int(cost_year) - i + 1
next

for i = 0 to 13
	sum_amt(i) = 0
	tot_amt(i) = 0
next

if view_sw = "0" then
	condi_sql = ""
end if

if view_sw = "1" then
	condi_sql = " and cost_center = '상주직접비' and company = '"&reside&"'"
end if
if view_sw = "2" then
	condi_sql = " and cost_center = '"&common&"'"
end if
if view_sw = "3" then
	condi_sql = " and cost_center = '직접비' and saupbu = '"&direct&"'"
end if
if view_sw = "4" then
	condi_sql = " and cost_center = '상주직접비' and saupbu = '"&direct&"'"
end if

title_line = "비용 유형별 현황"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>
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
				console.log(document.frm.view_sw[0].checked);
				if (eval("document.frm.view_sw[0].checked")) {
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = 'none';
				}
				if (eval("document.frm.view_sw[1].checked")) {
					document.getElementById('reside_view').style.display = '';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = 'none';
				}
				if (eval("document.frm.view_sw[2].checked")) {
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = '';
					document.getElementById('direct_view').style.display = 'none';
				}
				if (eval("document.frm.view_sw[3].checked")) {
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = '';
				}
				if (eval("document.frm.view_sw[4].checked")) {
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = '';
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
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="reside_cost_report.asp" method="post" name="frm">
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
							<input type="radio" name="view_sw" value="0" <% if view_sw = "0" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="condi_view()"><strong>총괄</strong>
							<input type="radio" name="view_sw" value="1" <% if view_sw = "1" then %>checked<% end if %> style="width:30px" id="Radio3" onClick="condi_view()"><strong>상주처별</strong>
							<!--<input type="radio" name="view_sw" value="2" <% if view_sw = "2" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="condi_view()"><strong>공통비</strong>-->
							<input type="radio" name="view_sw" value="3" <% if view_sw = "3" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="condi_view()"><strong>직접비</strong>
							<input type="radio" name="view_sw" value="4" <% if view_sw = "4" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="condi_view()"><strong>상주직접비</strong>
							</label>
							<label>
                            <select name="reside" id="reside_view" style="width:150px">
                              <option value="선택" <% if reside = "" then %>selected<% end if %>>선택</option>
                              <%
                                Sql="select company from company_cost where (cost_center = '상주직접비') group by company order by company asc"
                                rs_org.Open Sql, Dbconn, 1
                                do until rs_org.eof
                                %>
                              <option value='<%=rs_org("company")%>' <%If reside = rs_org("company") then %>selected<% end if %>><%=rs_org("company")%></option>
                              <%
                                    rs_org.movenext()
                                loop
                                rs_org.close()
                                %>
                            </select>
                            <select name="common" id="common_view" style="width:150px">
                              <option value="부문공통비" <% if common = "부문공통비" then %>selected<% end if %>>부문공통비</option>
                              <option value="전사공통비" <% if common = "전사공통비" then %>selected<% end if %>>전사공통비</option>
                              <option value="회사간거래" <% if common = "회사간거래" then %>selected<% end if %>>회사간거래</option>
                            </select>
                            <select name="direct" id="direct_view" style="width:150px; display:none">
                              <option value="" <% if direct = "" then %>selected<% end if %>>사업부미지정</option>
                              <%
								sql = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
                                rs_org.Open Sql, Dbconn, 1
                                do until rs_org.eof
                                %>
                              <option value='<%=rs_org("saupbu")%>' <%If direct = rs_org("saupbu") then %>selected<% end if %>><%=rs_org("saupbu")%></option>
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
								<td class="right">
						<%	if view_sw = "0" or view_sw = "1" or view_sw = "3" or view_sw = "4" or cost_tab(jj) = "인건비" then	%>
								<%=formatnumber(cost_amt,0)%>
						<%	  else	%>
						<%		if	view_sw = "2" then	%>
                                <a href="#" onClick="pop_Window('person_company_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&view_sw=<%=view_sw%>&reside=<%=reside%>&common=<%=common%>&direct=<%=direct%>','person_company_cost_view_pop','scrollbars=yes,width=800,height=500')"><%=formatnumber(cost_amt,0)%></a>
						<%		  else	%>
								<%=formatnumber(cost_amt,0)%>
                        <%		end if	%>
						<%	end if	%>
                                </td>
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
								<td class="right">
						<%	if view_sw = "0" or view_sw = "1" or view_sw = "3" or view_sw = "4" or cost_tab(jj) = "인건비" then	%>
								<%=formatnumber(cost_amt,0)%>
						<%	  else	%>
						<%		if	view_sw = "2" then	%>
                                <a href="#" onClick="pop_Window('person_company_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rs("cost_detail")%>&view_sw=<%=view_sw%>&reside=<%=reside%>&common=<%=common%>&direct=<%=direct%>','person_company_cost_view_pop','scrollbars=yes,width=800,height=500')"><%=formatnumber(cost_amt,0)%></a>
						<%		  else	%>
								<%=formatnumber(cost_amt,0)%>
                        <%		end if	%>
						<%	end if	%>
                                </td>
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
				    <td width="25%">
					<div class="btnCenter">
                    <a href="reside_cost_excel.asp?cost_year=<%=cost_year%>&view_sw=<%=view_sw%>&reside=<%=reside%>&common=<%=common%>&direct=<%=direct%>" class="btnType04">엑셀다운로드</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>
	</div>
	</body>
</html>

