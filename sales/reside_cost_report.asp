<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim year_tab(5)
Dim sum_amt(13)
Dim tot_amt(13)
Dim cost_tab
Dim cost_year, view_sw, common, direct, be_year, title_line, base_year
Dim condi_sql, i

cost_tab = Array("인건비", "야특근", "일반경비", "교통비", "법인카드", "임차료", "외주비", "자재", "장비", "운반비", "상각비")

cost_year = f_Request("cost_year")
view_sw = f_Request("view_sw")
reside = f_Request("reside")
common = f_Request("common")
direct = f_Request("direct")

title_line = "비용 유형별 현황"

If cost_year = "" Then
	cost_year = Mid(CStr(Now()), 1, 4)
	base_year = cost_year
	view_sw = "0"
End If

be_year = Int(cost_year) - 1

For i = 1 To 5
	year_tab(i) = Int(cost_year) - i + 1
Next

For i = 0 To 13
	sum_amt(i) = 0
	tot_amt(i) = 0
Next

If view_sw = "0" Then
	condi_sql = ""
End If

If view_sw = "1" Then
	condi_sql = "AND cost_center = '상주직접비' AND company = '"&reside&"' "
End If

If view_sw = "2" Then
	condi_sql = "AND cost_center = '"&common&"' "
End If

If view_sw = "3" Then
	condi_sql = "AND cost_center = '직접비' AND saupbu = '"&direct&"' "
End If

If view_sw = "4" Then
	condi_sql = "AND cost_center = '상주직접비' AND saupbu = '"&direct&"' "
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.cost_year.value == ""){
					alert ("조회년을 입력하세요.");
					return false;
				}
				return true;
			}

			function condi_view(){
				console.log(document.frm.view_sw[0].checked);

				if(eval("document.frm.view_sw[0].checked")){
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = 'none';
				}

				if(eval("document.frm.view_sw[1].checked")){
					document.getElementById('reside_view').style.display = '';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = 'none';
				}

				if(eval("document.frm.view_sw[2].checked")){
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = '';
					document.getElementById('direct_view').style.display = 'none';
				}

				if(eval("document.frm.view_sw[3].checked")){
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = '';
				}

				if(eval("document.frm.view_sw[4].checked")){
					document.getElementById('reside_view').style.display = 'none';
					document.getElementById('common_view').style.display = 'none';
					document.getElementById('direct_view').style.display = '';
				}
			}

			function scrollAll(){
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
				<form action="/sales/reside_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건 검색</dt>
                        <dd>
                            <p>
							<label>
								&nbsp;&nbsp;<strong>조회년&nbsp;</strong> :
								<select name="cost_year" id="cost_year" style="width:70px">
								<%For i = 1 To 5 %>
								  <option value="<%=year_tab(i)%>" <%If cost_year=year_tab(i) Then %> selected <%End If %>>&nbsp;<%=year_tab(i)%></option>
								<%Next	%>
								</select>
								</label>
								<label>
								<input type="radio" name="view_sw" value="0" <%If view_sw = "0" Then %> checked <%End If %> style="width:30px" id="Radio3" onClick="condi_view()">
								<strong>총괄</strong>

								<input type="radio" name="view_sw" value="1" <%If view_sw = "1" Then %> checked <%End If %> style="width:30px" id="Radio3" onClick="condi_view()">
								<strong>상주처별</strong>

								<!--<input type="radio" name="view_sw" value="2" <% if view_sw = "2" then %>checked<% end if %> style="width:30px" id="Radio4" onClick="condi_view()"><strong>공통비</strong>-->

								<input type="radio" name="view_sw" value="3" <%If view_sw = "3" Then %> checked <%End If %> style="width:30px" id="Radio4" onClick="condi_view()">
								<strong>직접비</strong>

								<input type="radio" name="view_sw" value="4" <%If view_sw = "4" Then %> checked <%End If %> style="width:30px" id="Radio4" onClick="condi_view()">
								<strong>상주직접비</strong>
							</label>
							<label>
								<select name="reside" id="reside_view" style="width:150px">
									<option value="선택" <% if reside = "" then %>selected<% end if %>>선택</option>
									<%
									Dim rs_org, rsSaupbu

									'Sql="select company from company_cost where (cost_center = '상주직접비') group by company order by company asc"
									objBuilder.Append "SELECT company FROM company_cost "
									objBuilder.Append "WHERE cost_center = '상주직접비' "
									objBuilder.Append "GROUP BY company "
									objBuilder.Append "ORDER BY company ASC"

									Set rs_org = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									Do Until rs_org.EOF
									%>
									<option value='<%=rs_org("company")%>' <%If reside = rs_org("company") Then %> selected <%End If %>><%=rs_org("company")%></option>
									<%
										rs_org.MoveNext()
									Loop
									rs_org.Close() : Set rs_org = Nothing
									%>
								</select>

								<select name="common" id="common_view" style="width:150px">
									<option value="부문공통비" <%If common = "부문공통비" Then %> selected <%End If %>>부문공통비</option>
									<option value="전사공통비" <%If common = "전사공통비" Then %> selected <%End If %>>전사공통비</option>
									<option value="회사간거래" <%If common = "회사간거래" Then %> selected <%End If %>>회사간거래</option>
								</select>

								<select name="direct" id="direct_view" style="width:150px; display:none;">
									<option value="" <%If direct = "" Then %> selected <%End If %>>사업부미지정</option>
									<%
									'sql = "select saupbu from sales_org where sales_year='" & cost_year & "' order by sort_seq"
									objBuilder.Append "SELECT saupbu FROM sales_org WHERE sales_year = '"&cost_year&"' "
									objBuilder.Append "ORDER BY sort_seq "

									Set rsSaupbu = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									Do Until rsSaupbu.EOF
									%>
									<option value='<%=rsSaupbu("saupbu")%>' <%If direct = rsSaupbu("saupbu") Then %> selected <%End If %>><%=rsSaupbu("saupbu")%></option>
									<%
										rsSaupbu.MoveNext()
									Loop
									rsSaupbu.Close() : Set rsSaupbu = Nothing
									%>
								</select>
							</label>
                            <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="검색"></a>
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
      					<div id="topLine2" style="width:1200px;overflow:hidden;">
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
						Dim jj, rec_cnt, rsDetail, rsTotal
						Dim tot_cost_amt, rs_etc, be_cost_amt, k, kk, cost, cost_amt
						Dim cr_pro

						For jj = 0 To 10
							rec_cnt = 0

							'sql = "select cost_detail from company_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail"
							objBuilder.Append "SELECT cost_detail FROM company_cost "
							objBuilder.Append "WHERE cost_year = '"&cost_year&"' AND cost_id = '"&cost_tab(jj)&"' "
							objBuilder.Append condi_sql
							objBuilder.Append "GROUP BY cost_detail "

							'rs.Open sql, Dbconn, 1
							Set rsDetail = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rsDetail.EOF
								rec_cnt = rec_cnt + 1

								rsDetail.MoveNext()
							Loop
							rsDetail.Close() : Set rsDetail = Nothing

							If rec_cnt <> 0 Then
								If cost_tab(jj) = "인건비" Then
									'sql = "select company_cost.cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost inner join saupbu_cost_account on company_cost.cost_id = saupbu_cost_account.cost_id and company_cost.cost_detail = saupbu_cost_account.cost_detail where company_cost.cost_year ='"&cost_year&"' and company_cost.cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by company_cost.cost_detail order by saupbu_cost_account.view_seq"
									objBuilder.Append "select coct.cost_detail, sum(cost_amt_01) as cost_amt_01, sum(cost_amt_02) as cost_amt_02, "
									objBuilder.Append "	sum(cost_amt_03) as cost_amt_03, sum(cost_amt_04) as cost_amt_04, sum(cost_amt_05) as cost_amt_05, "
									objBuilder.Append "	sum(cost_amt_06) as cost_amt_06, sum(cost_amt_07) as cost_amt_07, sum(cost_amt_08) as cost_amt_08, "
									objBuilder.Append "	sum(cost_amt_09) as cost_amt_09, sum(cost_amt_10) as cost_amt_10, sum(cost_amt_11) as cost_amt_11, "
									objBuilder.Append "	sum(cost_amt_12) as cost_amt_12 "
									objBuilder.Append "from company_cost AS coct "
									objBuilder.Append "inner join saupbu_cost_account AS scat on coct.cost_id = scat.cost_id "
									objBuilder.Append "	and coct.cost_detail = scat.cost_detail "
									objBuilder.Append "where coct.cost_year ='"&cost_year&"' and coct.cost_id ='"&cost_tab(jj)&"'"&condi_sql&" "
									objBuilder.Append "group by coct.cost_detail "
									objBuilder.Append "order by scat.view_seq "
								Else
									'sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
									objBuilder.Append "select cost_detail, sum(cost_amt_01) as cost_amt_01, sum(cost_amt_02) as cost_amt_02, "
									objBuilder.Append "	sum(cost_amt_03) as cost_amt_03, sum(cost_amt_04) as cost_amt_04, sum(cost_amt_05) as cost_amt_05, "
									objBuilder.Append "	sum(cost_amt_06) as cost_amt_06, sum(cost_amt_07) as cost_amt_07, sum(cost_amt_08) as cost_amt_08, "
									objBuilder.Append "	sum(cost_amt_09) as cost_amt_09, sum(cost_amt_10) as cost_amt_10, sum(cost_amt_11) as cost_amt_11, "
									objBuilder.Append "	sum(cost_amt_12) as cost_amt_12 "
									objBuilder.Append "from company_cost AS coct "
									objBuilder.Append "where cost_year ='"&cost_year&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" "
									objBuilder.Append "group by cost_detail "
									objBuilder.Append "order by cost_detail "
								End If

								'rs.Open sql, Dbconn, 1
								Set rsTotal = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								tot_cost_amt = cdbl(rsTotal("cost_amt_01")) + cdbl(rsTotal("cost_amt_02")) + cdbl(rsTotal("cost_amt_03")) + cdbl(rsTotal("cost_amt_04"))
								tot_cost_amt = tot_cost_amt  + cdbl(rsTotal("cost_amt_05")) + cdbl(rsTotal("cost_amt_06")) + cdbl(rsTotal("cost_amt_07"))
								tot_cost_amt = tot_cost_amt + cdbl(rsTotal("cost_amt_08")) + cdbl(rsTotal("cost_amt_09")) + cdbl(rsTotal("cost_amt_10"))
								tot_cost_amt = tot_cost_amt + cdbl(rsTotal("cost_amt_11")) + cdbl(rsTotal("cost_amt_12"))

								sum_amt(0) = sum_amt(0) + tot_cost_amt
								sum_amt(1) = sum_amt(1) + cdbl(rsTotal("cost_amt_01"))
								sum_amt(2) = sum_amt(2) + cdbl(rsTotal("cost_amt_02"))
								sum_amt(3) = sum_amt(3) + cdbl(rsTotal("cost_amt_03"))
								sum_amt(4) = sum_amt(4) + cdbl(rsTotal("cost_amt_04"))
								sum_amt(5) = sum_amt(5) + cdbl(rsTotal("cost_amt_05"))
								sum_amt(6) = sum_amt(6) + cdbl(rsTotal("cost_amt_06"))
								sum_amt(7) = sum_amt(7) + cdbl(rsTotal("cost_amt_07"))
								sum_amt(8) = sum_amt(8) + cdbl(rsTotal("cost_amt_08"))
								sum_amt(9) = sum_amt(9) + cdbl(rsTotal("cost_amt_09"))
								sum_amt(10) = sum_amt(10) + cdbl(rsTotal("cost_amt_10"))
								sum_amt(11) = sum_amt(11) + cdbl(rsTotal("cost_amt_11"))
								sum_amt(12) = sum_amt(12) + cdbl(rsTotal("cost_amt_12"))

								' 전년 자료
								'sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost where cost_year ='"&be_year&"' and cost_detail ='"&rs("cost_detail")&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
								objBuilder.Append "select cost_detail, sum(cost_amt_01) as cost_amt_01, sum(cost_amt_02) as cost_amt_02, "
								objBuilder.Append "	sum(cost_amt_03) as cost_amt_03, sum(cost_amt_04) as cost_amt_04, sum(cost_amt_05) as cost_amt_05, "
								objBuilder.Append "	sum(cost_amt_06) as cost_amt_06, sum(cost_amt_07) as cost_amt_07, sum(cost_amt_08) as cost_amt_08, "
								objBuilder.Append "	sum(cost_amt_09) as cost_amt_09, sum(cost_amt_10) as cost_amt_10, sum(cost_amt_11) as cost_amt_11, "
								objBuilder.Append "	sum(cost_amt_12) as cost_amt_12 "
								objBuilder.Append "from company_cost "
								objBuilder.Append "where cost_year ='"&be_year&"' and cost_detail ='"&rsTotal("cost_detail")&"' "
								objBuilder.Append "	and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" "
								objBuilder.Append "group by cost_detail "
								objBuilder.Append "order by cost_detail "

								set rs_etc = DbConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								if rs_etc.eof or rs_etc.bof then
									be_cost_amt = 0
								else
									be_cost_amt = cdbl(rs_etc("cost_amt_01")) + cdbl(rs_etc("cost_amt_02")) + cdbl(rs_etc("cost_amt_03")) + cdbl(rs_etc("cost_amt_04")) + cdbl(rs_etc("cost_amt_05")) + cdbl(rs_etc("cost_amt_06")) + cdbl(rs_etc("cost_amt_07")) + cdbl(rs_etc("cost_amt_08")) + cdbl(rs_etc("cost_amt_09")) + cdbl(rs_etc("cost_amt_10")) + cdbl(rs_etc("cost_amt_11")) + cdbl(rs_etc("cost_amt_12"))
								end If

								rs_etc.close() : Set rs_etc = Nothing

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
									<td class="left"><%=rsTotal("cost_detail")%></td>
									<td class="right" bgcolor="#FFFFCC"><%=formatnumber(be_cost_amt/1000,0)%></td>
							<%
								for k = 1 to 12
									if k < 10 then
										kk = "0" + cstr(k)
									  else
										kk = cstr(k)
									end if
									cost = "cost_amt_" + cstr(kk)
									cost_amt = rsTotal(cost)
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
									<a href="#" onClick="pop_Window('person_company_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rsTotal("cost_detail")%>&view_sw=<%=view_sw%>&reside=<%=reside%>&common=<%=common%>&direct=<%=direct%>','person_company_cost_view_pop','scrollbars=yes,width=800,height=500')"><%=formatnumber(cost_amt,0)%></a>
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
								rsTotal.movenext()

								do until rsTotal.eof
									'tot_cost_amt = cdbl(rs("cost_amt_01")) + cdbl(rs("cost_amt_02")) + cdbl(rs("cost_amt_03")) + cdbl(rs("cost_amt_04")) + cdbl(rs("cost_amt_05")) + cdbl(rs("cost_amt_06")) + cdbl(rs("cost_amt_07")) + cdbl(rs("cost_amt_08")) + cdbl(rs("cost_amt_09")) + cdbl(rs("cost_amt_10")) + cdbl(rs("cost_amt_11")) + cdbl(rs("cost_amt_12"))
									tot_cost_amt = cdbl(rsTotal("cost_amt_01")) + cdbl(rsTotal("cost_amt_02")) + cdbl(rsTotal("cost_amt_03")) + cdbl(rsTotal("cost_amt_04"))
									tot_cost_amt = tot_cost_amt  + cdbl(rsTotal("cost_amt_05")) + cdbl(rsTotal("cost_amt_06")) + cdbl(rsTotal("cost_amt_07"))
									tot_cost_amt = tot_cost_amt + cdbl(rsTotal("cost_amt_08")) + cdbl(rsTotal("cost_amt_09")) + cdbl(rsTotal("cost_amt_10"))
									tot_cost_amt = tot_cost_amt + cdbl(rsTotal("cost_amt_11")) + cdbl(rsTotal("cost_amt_12"))

									sum_amt(0) = sum_amt(0) + tot_cost_amt
									sum_amt(1) = sum_amt(1) + cdbl(rsTotal("cost_amt_01"))
									sum_amt(2) = sum_amt(2) + cdbl(rsTotal("cost_amt_02"))
									sum_amt(3) = sum_amt(3) + cdbl(rsTotal("cost_amt_03"))
									sum_amt(4) = sum_amt(4) + cdbl(rsTotal("cost_amt_04"))
									sum_amt(5) = sum_amt(5) + cdbl(rsTotal("cost_amt_05"))
									sum_amt(6) = sum_amt(6) + cdbl(rsTotal("cost_amt_06"))
									sum_amt(7) = sum_amt(7) + cdbl(rsTotal("cost_amt_07"))
									sum_amt(8) = sum_amt(8) + cdbl(rsTotal("cost_amt_08"))
									sum_amt(9) = sum_amt(9) + cdbl(rsTotal("cost_amt_09"))
									sum_amt(10) = sum_amt(10) + cdbl(rsTotal("cost_amt_10"))
									sum_amt(11) = sum_amt(11) + cdbl(rsTotal("cost_amt_11"))
									sum_amt(12) = sum_amt(12) + cdbl(rsTotal("cost_amt_12"))
									' 전년 자료
									'sql = "select cost_detail,sum(cost_amt_01) as cost_amt_01,sum(cost_amt_02) as cost_amt_02,sum(cost_amt_03) as cost_amt_03,sum(cost_amt_04) as cost_amt_04,sum(cost_amt_05) as cost_amt_05,sum(cost_amt_06) as cost_amt_06,sum(cost_amt_07) as cost_amt_07,sum(cost_amt_08) as cost_amt_08,sum(cost_amt_09) as cost_amt_09,sum(cost_amt_10) as cost_amt_10,sum(cost_amt_11) as cost_amt_11,sum(cost_amt_12) as cost_amt_12 from company_cost where cost_year ='"&be_year&"' and cost_detail ='"&rs("cost_detail")&"' and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
									objBuilder.Append "select cost_detail,sum(cost_amt_01) as cost_amt_01, sum(cost_amt_02) as cost_amt_02, "
									objBuilder.Append "	sum(cost_amt_03) as cost_amt_03, sum(cost_amt_04) as cost_amt_04, sum(cost_amt_05) as cost_amt_05, "
									objBuilder.Append "	sum(cost_amt_06) as cost_amt_06, sum(cost_amt_07) as cost_amt_07, sum(cost_amt_08) as cost_amt_08, "
									objBuilder.Append "	sum(cost_amt_09) as cost_amt_09, sum(cost_amt_10) as cost_amt_10, sum(cost_amt_11) as cost_amt_11, "
									objBuilder.Append "	sum(cost_amt_12) as cost_amt_12 "
									objBuilder.Append "from company_cost "
									objBuilder.Append "where cost_year ='"&be_year&"' and cost_detail ='"&rsTotal("cost_detail")&"' "
									objBuilder.Append "	and cost_id ='"&cost_tab(jj)&"'"&condi_sql&" "
									objBuilder.Append "group by cost_detail "
									objBuilder.Append "order by cost_detail "

									set rs_etc = DbConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									if rs_etc.eof or rs_etc.bof then
										be_cost_amt = 0
									else
										be_cost_amt = cdbl(rs_etc("cost_amt_01")) + cdbl(rs_etc("cost_amt_02")) + cdbl(rs_etc("cost_amt_03")) + cdbl(rs_etc("cost_amt_04")) + cdbl(rs_etc("cost_amt_05")) + cdbl(rs_etc("cost_amt_06")) + cdbl(rs_etc("cost_amt_07")) + cdbl(rs_etc("cost_amt_08")) + cdbl(rs_etc("cost_amt_09")) + cdbl(rs_etc("cost_amt_10")) + cdbl(rs_etc("cost_amt_11")) + cdbl(rs_etc("cost_amt_12"))
									end if
									rs_etc.close()
									sum_amt(13) = sum_amt(13) + be_cost_amt
							%>
								<tr>
									<td class="left" style=" border-left:1px solid #e3e3e3;"><%=rsTotal("cost_detail")%></td>
									<td class="right" bgcolor="#FFFFCC"><%=formatnumber(be_cost_amt/1000,0)%></td>
							<%
								for k = 1 to 12
									if k < 10 then
										kk = "0" + cstr(k)
									  else
										kk = cstr(k)
									end if
									cost = "cost_amt_" + cstr(kk)
									cost_amt = rsTotal(cost)
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
									<a href="#" onClick="pop_Window('person_company_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=k%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=rsTotal("cost_detail")%>&view_sw=<%=view_sw%>&reside=<%=reside%>&common=<%=common%>&direct=<%=direct%>','person_company_cost_view_pop','scrollbars=yes,width=800,height=500')"><%=formatnumber(cost_amt,0)%></a>
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
									rsTotal.movenext()
								loop
								rsTotal.close() : Set rsTotal = Nothing
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

