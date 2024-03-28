<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim sum_amt(10)
Dim tot_amt(10)
Dim detail_tab(30)
Dim cost_amt(30,10)

Dim cost_tab

Dim sales_saupbu, cost_year, cost_mm, cost_month
Dim before_year, before_mm, before_month, c_month, b_month
Dim condi_sql

Dim i
Dim rsPreCostSum, before_sales_amt
Dim rsCurrCostSum, curr_sales_amt

Dim title_line

cost_tab = Array("인건비","야특근","일반경비","교통비","법인카드","임차료","외주비","자재","장비","운반비","상각비")

sales_saupbu = Request("sales_saupbu")
cost_year = Request("cost_year")
cost_mm = Right("0" & CStr(Request("cost_mm")), 2)
cost_month = CStr(cost_year) & CStr(cost_mm)

If cost_mm = "01" Then
	before_year = CStr(Int(cost_year) - 1)
	before_mm = "12"
Else
	before_year = cost_year
	before_mm = Right("0" & CStr(Int(cost_mm) - 1),2)
End If

before_month = CStr(before_year) & CStr(before_mm)	'이전 년도(yyyyMM)

c_month = CStr(cost_year) & "-" & CStr(cost_mm)		'당월 년도(yyyy-MM)
b_month = CStr(before_year) & "-" & CStr(before_mm)	'이전 년도(yyyy-MM)

'If sales_saupbu = "전체" Then
'	condi_sql = ""
'Else
'	condi_sql = " AND saupbu ='"&sales_saupbu&"' "
'End If

'If sales_saupbu = "기타사업부" Then
'	condi_sql = " AND (saupbu ='' OR saupbu = '기타사업부') "
'End If

'If sales_saupbu = "한진" OR sales_saupbu = "한진그룹" Then
'	condi_sql = " AND saupbu IN ('한진', '한진그룹') "
'End If

Select Case sales_saupbu
	Case "전체"
		condi_sql = ""
	Case "기타사업부"
		condi_sql = " AND (saupbu ='' OR saupbu = '기타사업부') "
	Case "한진", "한진그룹"
		condi_sql = " AND saupbu IN ('한진', '한진그룹') "
	Case Else
		condi_sql = " AND saupbu ='"&sales_saupbu&"' "
End Select

For i = 0 To 10
	sum_amt(i) = 0
	tot_amt(i) = 0
Next

'매출계(전월)
'sql = "SELECT SUM(cost_amt) AS sales_amt "&_
'	  "  FROM saupbu_sales "&_
'	  " WHERE SUBSTRING(SALES_DATE,1,7) = '"&b_month&"'"&condi_sql
'Set rs_sum = Dbconn.Execute(sql)
objBuilder.Append "SELECT SUM(cost_amt) AS sales_amt "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(SALES_DATE, 1, 7) = '"&b_month&"'"&condi_sql

Set rsPreCostSum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsPreCostSum(0)) Then
	before_sales_amt = 0
Else
	before_sales_amt = CDbl(rsPreCostSum(0))
End If

rsPreCostSum.Close()
Set rsPreCostSum = Nothing

'매출계(당월)
'sql = "SELECT SUM(cost_amt) AS sales_amt "&_
'	  "  FROM saupbu_sales "&_
'	  " WHERE SUBSTRING(sales_date,1,7) = '"&c_month&"'"&condi_sql
'Set rs_sum = Dbconn.Execute (sql)
objBuilder.Append "SELECT SUM(cost_amt) AS sales_amt "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&c_month&"'"&condi_sql

Set rsCurrCostSum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsCurrCostSum(0)) Then
	curr_sales_amt = 0
Else
	curr_sales_amt = CDbl(rsCurrCostSum(0))
End If

rsCurrCostSum.Close()
Set rsCurrCostSum = Nothing

title_line = sales_saupbu + " 손익 현황"

If sales_saupbu = "" Then
	title_line = "기타사업부 손익 현황"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
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
				return "2 1";
			}

			function frmcheck(){
				if (chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if (document.frm.cost_month.value == ""){
					alert ("조회년을 입력하세요.");
					return false;
				}
				return true;
			}

			function scrollAll(){
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
						<td>
							<div id="topLine2" style="width:1200px;overflow:hidden;">
								<div class="gView">
									<table cellpadding="0" cellspacing="0" class="tableList">
										<colgroup>
											<col width="4%" >
											<col width="*" >
											<col width="8%" >
											<col width="6%" >
											<col width="6%" >
											<col width="7%" >
											<col width="9%" >
											<col width="8%" >
											<col width="6%" >
											<col width="6%" >
											<col width="7%" >
											<col width="9%" >
											<col width="7%" >
											<col width="5%" >
											<col width="1%" >
										</colgroup>
										<thead>
											<tr>
											  <th rowspan="2" class="first" scope="col">비용항목</th>
											  <th rowspan="2" scope="col">세부내역</th>
											  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">전 월&nbsp;(<%=before_year%>년<%=before_mm%>월)</th>
											  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">당 월&nbsp;(<%=cost_year%>년<%=cost_mm%>월)</th>
											  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">증감</th>
											  <th rowspan="2" scope="col"></th>
										  	</tr>
											<tr>
											  <th scope="col" style="border-left:1px solid #e3e3e3;">상주직접비</th>
											  <th scope="col">직접비</th>
											  <th scope="col">전사공통비</th>
											  <th scope="col">부문공통비</th>
											  <th scope="col">계</th>
											  <th scope="col">상주직접비</th>
											  <th scope="col">직접비</th>
											  <th scope="col">전사공통비</th>
											  <th scope="col">부문공통비</th>
											  <th scope="col">계</th>
											  <th scope="col">금액</th>
											  <th scope="col">율</th>
				              				</tr>
										</thead>
									</table>
								</div>
							</div>
						</td>
					</tr>
					<tr>
          				<td valign="top">
				    	<DIV id="mainDisplay2" style="width:1200;height:470px;overflow:scroll" onscroll="scrollAll()">
				    	<table cellpadding="0" cellspacing="0" class="scrollList">
				    		<colgroup>
								<col width="6%" >
								<col width="*" >
								<col width="8%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="9%" >
								<col width="8%" >
								<col width="6%" >
								<col width="6%" >
								<col width="7%" >
								<col width="9%" >
								<col width="7%" >
								<col width="5%" >
								<col width="1%" >
							</colgroup>
							<tbody>
								<tr bgcolor="#FFFFCC">
									<td colspan="2" class="first" scope="col"><strong>매출계</strong></td>
									<td colspan="5" scope="col"><strong><%=FormatNumber(before_sales_amt, 0)%></strong></td>
									<td colspan="5" scope="col"><strong><%=FormatNumber(curr_sales_amt, 0)%></strong></td>
									<%
									Dim incr_amt, incr_per

									incr_amt = curr_sales_amt - before_sales_amt

									If before_sales_amt = 0 And curr_sales_amt = 0 Then
										incr_per = 0
							  		ElseIf before_sales_amt = 0 Then
										incr_per = 100
							  		Else
						   				incr_per = incr_amt / before_sales_amt * 100
						   			End If
									%>
									<td scope="col" class="right"><%=FormatNumber(incr_amt, 0)%></td>
									<td scope="col" class="right"><%=FormatNumber(incr_per, 2)%>%</td>
									<td scope="col" class="right">&nbsp;</td>
                    			</tr>
								<%
								' cost_tab = array(0"인건비",1"야특근",2"일반경비",3"교통비",4"법인카드",5"임차료",6"외주비",7"자재",8"장비",9"운반비",10"상각비")
								Dim jj, rec_cnt, j
								Dim rsCostDetail, rsCostSum

								Set rsCostDetail = Server.CreateObject("ADODB.RecordSet")
								Set rsPreCostSum = Server.CreateObject("ADODB.RecordSet")
								Set rsCurrCostSum = Server.CreateObject("ADODB.RecordSet")

								For jj = 0 To 10
									rec_cnt = 0

									For i = 1 To 30
										detail_tab(i) = ""

										For j = 1 To 10
											cost_amt(i, j) = 0
											sum_amt(j) = 0
										Next
									Next

									If cost_tab(jj) = "인건비" Then
										'sql = "   SELECT cost_detail "&_
										'	  "     FROM SAUPBU_COST_ACCOUNT "&_
										'	  "    WHERE cost_id = '인건비' "&_
										'	  " ORDER BY view_seq"
										'rs.Open sql, Dbconn, 1
										objBuilder.Append "SELECT cost_detail "
										objBuilder.Append "FROM SAUPBU_COST_ACCOUNT "
										objBuilder.Append "WHERE cost_id = '인건비' "
										objBuilder.Append "ORDER BY view_seq "

										rsCostDetail.Open objBuilder.ToString(), DBConn, 1
										objBuilder.Clear()

										Do Until rsCostDetail.EOF
											rec_cnt = rec_cnt + 1
											detail_tab(rec_cnt) = rsCostDetail("cost_detail")

											rsCostDetail.MoveNext()
										Loop

										rsCostDetail.Close()
									Else
										'sql = "   SELECT cost_detail "&_
										'	  "     FROM SAUPBU_PROFIT_LOSS "&_
										'	  "    WHERE (cost_year = '"& cost_year &"' OR cost_year = '"& before_year &"') "&_
										'	  "      AND cost_id ='"& cost_tab(jj) &"'"& condi_sql &_
										'	  " GROUP BY cost_detail "&_
										'	  " ORDER BY cost_detail"
										'rs.Open sql, Dbconn, 1
										objBuilder.Append "SELECT cost_detail "
										objBuilder.Append "FROM SAUPBU_PROFIT_LOSS "
										objBuilder.Append "WHERE (cost_year = '"& cost_year &"' OR cost_year = '"& before_year &"') "
										objBuilder.Append "	AND cost_id ='"& cost_tab(jj) &"'"& condi_sql
										objBuilder.Append "GROUP BY cost_detail "
										objBuilder.Append "ORDER BY cost_detail "

										rsCostDetail.Open objBuilder.ToString(), DBConn, 1
										objBuilder.Clear()

										Do Until rsCostDetail.EOF
											rec_cnt = rec_cnt + 1
											detail_tab(rec_cnt) = rsCostDetail("cost_detail")

											rsCostDetail.MoveNext()
										Loop

										rsCostDetail.Close()
									End If

									If rec_cnt <> 0 Then
										' 전월 금액 SUM
										'sql = "  SELECT cost_center "&_
										' 	  "       , cost_detail "&_
										'	  "       , SUM(cost_amt_"& before_mm &") AS cost " &_
										'	  "    FROM SAUPBU_PROFIT_LOSS "&_
										'	  "   WHERE cost_year = '"& before_year &"' "&_
										'	  "     AND cost_id   = '"& cost_tab(jj) &"'"&condi_sql &_
										'	  " GROUP BY cost_center, cost_detail "&_
										'	  " ORDER BY cost_center, cost_detail"
										'rs.Open sql, Dbconn, 1
										objBuilder.Append "SELECT cost_center, cost_detail, SUM(cost_amt_"& before_mm &") AS cost "
										objBuilder.Append "FROM SAUPBU_PROFIT_LOSS "
										objBuilder.Append "WHERE cost_year = '"& before_year &"' "
										objBuilder.Append "	AND cost_id = '"& cost_tab(jj) &"'"&condi_sql
										objBuilder.Append "GROUP BY cost_center, cost_detail "
										objBuilder.Append "ORDER BY cost_center, cost_detail "

										rsPreCostSum.Open objBuilder.ToString(), DBConn, 1
										objBuilder.Clear()

										Do Until rsPreCostSum.EOF
											For i = 1 To 30
												' 전월에는 있지만 detail_tab에 없다면 cost_detail은 나오지 않는다..
												If rsPreCostSum("cost_detail") = detail_tab(i) Then
													Select Case rsPreCostSum("cost_center")
														Case "상주직접비" : j = 1
														Case "직접비"     : j = 2
														Case "전사공통비" : j = 3
														Case "부문공통비" : j = 4
													End Select

													cost_amt(i, j) = cost_amt(i, j) + CDbl(rsPreCostSum("cost"))
													cost_amt(i, 5) = cost_amt(i, 5) + CDbl(rsPreCostSum("cost"))
													sum_amt(j) = sum_amt(j) + CDbl(rsPreCostSum("cost"))
													sum_amt(5) = sum_amt(5) + CDbl(rsPreCostSum("cost"))
													tot_amt(j) = tot_amt(j) + CDbl(rsPreCostSum("cost"))
													tot_amt(5) = tot_amt(5) + CDbl(rsPreCostSum("cost"))

													Exit For
												End If
											Next

											rsPreCostSum.MoveNext()
										Loop

										rsPreCostSum.Close()

										' 당월 금액 SUM
										'sql = "    SELECT cost_center "&_
										'	  "         , cost_detail "&_
										'	  "         , SUM(cost_amt_"&cost_mm&") AS cost "&_
										'	  "      FROM  SAUPBU_PROFIT_LOSS "&_
										'	  "     WHERE  cost_year ='"& cost_year &"' "&_
										'	  "       AND  cost_id   ='"& cost_tab(jj) &"'"&condi_sql&" "&_
										'	  " GROUP  BY cost_center, cost_detail "&_
										'	  " ORDER  BY cost_center, cost_detail"
										'rs.Open sql, Dbconn, 1
										objBuilder.Append "SELECT cost_center, cost_detail, SUM(cost_amt_"&cost_mm&") AS cost "
										objBuilder.Append "FROM  SAUPBU_PROFIT_LOSS "
										objBuilder.Append "WHERE  cost_year ='"& cost_year &"' "
										objBuilder.Append "	AND cost_id ='"& cost_tab(jj) &"' "&condi_sql
										objBuilder.Append "GROUP  BY cost_center, cost_detail "
										objBuilder.Append "ORDER  BY cost_center, cost_detail "

										rsCurrCostSum.Open objBuilder.ToString(), DBConn, 1
										objBuilder.Clear()

										Do Until rsCurrCostSum.EOF
											For i = 1 To 30
												' 전월에는 있지만 detail_tab에 없다면 cost_detail은 나오지 않는다..
												If rsCurrCostSum("cost_detail") = detail_tab(i) Then
													Select Case rsCurrCostSum("cost_center")
														Case "상주직접비"	: j = 6
														Case "직접비"	    : j = 7
														Case "전사공통비"	: j = 8
														Case "부문공통비"	: j = 9
													End Select

													cost_amt(i, j) = cost_amt(i, j) + CDbl(rsCurrCostSum("cost"))
													cost_amt(i, 10) = cost_amt(i, 10) + CDbl(rsCurrCostSum("cost"))
													sum_amt(j) = sum_amt(j) + CDbl(rsCurrCostSum("cost"))
													sum_amt(10) = sum_amt(10) + CDbl(rsCurrCostSum("cost"))
													tot_amt(j) = tot_amt(j) + CDbl(rsCurrCostSum("cost"))
													tot_amt(10) = tot_amt(10) + CDbl(rsCurrCostSum("cost"))

													Exit For
												End If
											Next

											rsCurrCostSum.MoveNext()
										Loop

										rsCurrCostSum.Close()
										%>
										<tr>
							  				<td rowspan="<%=rec_cnt + 1%>" class="first">
											<%
											If jj = 2 Or jj = 3 Then
												Response.Write cost_tab(jj) & "<BR/>(현금사용)"
											Else
												Response.Write cost_tab(jj)
											End If
											%>
                  							</td>
											<td class="left"><%=detail_tab(1)%></td>

											<%
											For j = 1 To 10
												If j = 5 Or j = 10 Then
													Response.write "<td class='right'><strong>"&FormatNumber(cost_amt(1, j), 0)&"</strong></td>"
												Else
													Response.write "<td class='right'>" ' [["&jj&"]][[cost_amt(1,"&j&")="&cost_amt(1,j)&"]]

													If jj < 2 Then
														Response.Write FormatNumber(cost_amt(1, j), 0)
													Else
														If(j = 1 Or j = 2 Or j = 6 Or j = 7) And jj > 1 And cost_amt(1,j) <> 0 Then
														%>
			                  								<a href="#" onClick="pop_Window('profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(1)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')">
																<%=FormatNumber(cost_amt(1, j), 0)%>
															</a>
														<%
														Else
			                  								Response.Write FormatNumber(cost_amt(1, j), 0)
			                  							End If
			                  						End If
			                  						%>
		                  							</td>
												<%
												End If
											Next

											incr_amt = cost_amt(1, 10) - cost_amt(1, 5)

											If cost_amt(1, 5) = 0 And cost_amt(1, 10) = 0 Then
												incr_per = 0
											ElseIf cost_amt(1, 5) = 0 Then
												incr_per = 100
											Else
												incr_per = incr_amt / cost_amt(1, 5) * 100
											End If
											%>
											<td class="right"><%=FormatNumber(incr_amt, 0)%></td>
											<td class="right"><%=FormatNumber(incr_per, 2)%>%</td>
											<td class="right">&nbsp;</td>
										</tr>
										<%For i = 2 To rec_cnt%>
										<tr>
											<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
											<%
											For j = 1 To 10
												If j = 5 Or j = 10 Then
													Response.Write "<td class='right'><strong>"&FormatNumber(cost_amt(i, j), 0)&"</strong></td>"
												Else
											%>
											<td class="right">
												<%If jj < 2	Then	'//2016-08-23 알바비 상세조회 링크 추가
													If detail_tab(i) = "알바비" Then
													%>
														<a href="#" onClick="pop_Window('profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(i)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')">
															<%=FormatNumber(cost_amt(i, j), 0)%>
														</a>
													<%
													Else
														Response.Write FormatNumber(cost_amt(i, j), 0)
													End IF
													%>
												<%Else 	%>
													<%
													If (j = 1 Or j = 2 Or j = 6 Or j = 7) And jj > 1 And cost_amt(i, j) <>  0 Then%>
														<a href="#" onClick="pop_Window('profit_loss_detail_view.asp?cost_month=<%=cost_month%>&before_month=<%=before_month%>&cost_id=<%=cost_tab(jj)%>&cost_detail=<%=detail_tab(i)%>&j=<%=j%>&mg_saupbu=<%=sales_saupbu%>','profit_loss_detail_view_pop','scrollbars=yes,width=1000,height=600')">
															<%=FormatNumber(cost_amt(i, j), 0)%>
														</a>
													<%
													Else
													%>
														<%=FormatNumber(cost_amt(i, j), 0)%>
													<%
													End If	%>
												<%End If%>
											</td>
											<%
												End If
											Next

											incr_amt = cost_amt(i, 10) - cost_amt(i, 5)

											If cost_amt(i, 5) = 0 And cost_amt(i, 10) = 0 Then
													incr_per = 0
											ElseIf cost_amt(i, 5) = 0 Then
												incr_per = 100
											Else
												incr_per = incr_amt / cost_amt(i,5) * 100
											End If
											%>
											<td class="right"><%=FormatNumber(incr_amt, 0)%></td>
											<td class="right"><%=FormatNumber(incr_per, 2)%>%</td>
											<td class="right">&nbsp;</td>
										</tr>
										<%Next	%>

										<!--=== 소계 ===-->
										<tr>
											<td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">소계</td>
											<%
											For j = 1 To 10
												If j = 5 Or j = 10 Then
											%>
											<td class="right" bgcolor="#EEFFFF"><strong><%=FormatNumber(sum_amt(j), 0)%></strong></td>
											<%
												Else
											%>
											<td class="right" bgcolor="#EEFFFF"><%=FormatNumber(sum_amt(j), 0)%></td>
											<%
												End If
											Next

											incr_amt = sum_amt(10) - sum_amt(5)

											If sum_amt(5) = 0 And sum_amt(10) = 0 Then
												incr_per = 0
											ElseIf sum_amt(5) = 0 Then
												incr_per = 100
											Else
												incr_per = incr_amt / sum_amt(5) * 100
											End If
											%>
											<td class="right" bgcolor="#EEFFFF"><%=FormatNumber(incr_amt, 0)%></td>
											<td class="right" bgcolor="#EEFFFF"><%=FormatNumber(incr_per, 2)%>%</td>
											<td class="right" bgcolor="#EEFFFF">&nbsp;</td>
										</tr>
									<%
									End If
								Next

								Set rsCostDetail = Nothing
								Set rsPreCostSum = Nothing
								Set rsCurrCostSum = Nothing

								DBConn.Close
								Set DBConn = Nothing
								%>

								<!--=====	비용합계	=====-->
								<tr bgcolor="#FFFFCC">
									<td colspan="2" class="first" scope="col"><strong>비용합계</strong></td>
									<%
									For j = 1 To 10
										If j = 5 Or j = 10 Then
									%>
									<td class="right"><strong><%=formatnumber(tot_amt(j),0)%></strong></td>
									<%
										Else
									%>
									<td class="right"><%=formatnumber(tot_amt(j),0)%></td>
									<%
										End If
									Next

									incr_amt = tot_amt(10) - tot_amt(5)

									If tot_amt(5) = 0 And tot_amt(10) = 0 Then
										incr_per = 0
									ElseIf tot_amt(5) = 0 Then
										incr_per = 100
									Else
										incr_per = incr_amt / tot_amt(5) * 100
									End if
									%>
									<td scope="col" class="right"><%=FormatNumber(incr_amt, 0)%></td>
									<td scope="col" class="right"><%=FormatNumber(incr_per, 2)%>%</td>
									<td scope="col" class="right">&nbsp;</td>
								</tr>

								<!--=====	손익	=====-->
								<tr bgcolor="#FFDFDF">
									<td colspan="2" bgcolor="#FFDFDF" class="first" scope="col"><strong>손익</strong></td>
									<%
										Dim be_profit_loss, curr_profit_loss

										be_profit_loss = before_sales_amt - tot_amt(5)
										curr_profit_loss = curr_sales_amt - tot_amt(10)
										incr_amt = curr_profit_loss - be_profit_loss

										If be_profit_loss = 0 And curr_profit_loss = 0 Then
											incr_per = 0
										ElseIf be_profit_loss = 0 Then
											incr_per = 100
										Else
											incr_per = incr_amt / be_profit_loss * 100
										End If

										If be_profit_loss < 0 Then
											incr_per = incr_per * -1
										End If
									%>
									<td scope="col" colspan="5"><strong><%=FormatNumber(be_profit_loss, 0)%></strong></td>
									<td scope="col" colspan="5"><strong><%=FormatNumber(curr_profit_loss, 0)%></strong></td>
									<td scope="col" class="right"><%=FormatNumber(incr_amt, 0)%></td>
									<td scope="col" class="right"><%=FormatNumber(incr_per, 2)%>%</td>
									<td scope="col" class="right">&nbsp;</td>
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
							<div class="btnCenter">
			            		<a href="/sales/excel/saupbu_profit_loss_excel_old.asp?cost_year=<%=cost_year%>&cost_mm=<%=cost_mm%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">화면 엑셀다운로드</a>
			            		<a href="/sales/excel/cost_center_detail_excel_old.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">상주비/직접비 엑셀다운로드</a>
			            		<a href="/sales/excel/saupbu_sales_detail_excel2_old.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">매출액 엑셀다운로드</a>
								<%If sales_grade = "0" Then	%>
			            			<a href="/sales/excel/cost_center_detail_excel_old.asp?cost_month=<%=cost_month%>&sales_saupbu=<%="전사공통비"%>" class="btnType04">전사공통비 엑셀다운로드</a>
			          				<a href="/sales/excel/cost_center_detail_excel_old.asp?cost_month=<%=cost_month%>&sales_saupbu=<%="부문공통비"%>" class="btnType04">부문공통비 엑셀다운로드</a>
								<%End If%>
							</div>
            			</td>
			    	</tr>
				  	</table>
					<br>
				</form>
			</div>
		</div>
	</body>
</html>
