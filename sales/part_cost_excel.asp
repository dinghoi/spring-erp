<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'on Error resume next
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
Dim cost_month, sales_saupbu, before_date, condi_sql
Dim mm, cost_year
Dim title_line, savefilename
Dim rsCostTotal, tot_part_cost
Dim rsComCost, rsAsTot, tot_part_cnt

Dim from_date, end_date, to_date

cost_month = Request("cost_month")
sales_saupbu = Request("sales_saupbu")

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date), 6, 2)
	sales_saupbu = "전체"
End If

if sales_saupbu = "전체" then
	condi_sql = ""
  else
  	condi_sql = " AND saupbu ='"&sales_saupbu&"'"
end If

from_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

mm = Mid(cost_month, 5, 2)
cost_year = Mid(cost_month, 1, 4)

title_line = cost_year & "년 " & mm & "월 " & sales_saupbu & " 부문 공통비 배분 현황"
savefilename = title_line & ".xls"

Call ViewExcelType(savefilename)

'부문공통비 전체 비용
objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS tot_cost "
objBuilder.Append "FROM company_cost "
objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
objBuilder.Append "	AND cost_center = '부문공통비' "

Set rsCostTotal = DbConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsCostTotal("tot_cost")) Then
	tot_part_cost = 0
Else
	tot_part_cost = CLng(rsCostTotal("tot_cost"))
End If

rsCostTotal.close() : Set rsCostTotal = Nothing

If sales_saupbu = "전체" Then
	condi_sql = ""
Else
  	condi_sql = " AND trat.saupbu ='"& sales_saupbu &"' "
End If

'A/S 전체 카운트
objBuilder.Append "SELECT SUM(as_total) AS tot_cnt "
objBuilder.Append "FROM as_acpt_status "
objBuilder.Append "WHERE as_month = '"&cost_month&"' " & condi_sql

Set rsAsTot = DBconn.Execute(objBuilder.ToString())
objBuilder.Clear()

tot_part_cnt = rsAsTot("tot_cnt")

rsAsTot.Close() : Set rsAsTot = Nothing

Dim rsPart, part_tot_cost, as_tot_cnt, set_tot_cost
'부문공통비(배분) - 설치공사 제외
objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_detail = '설치공사')) AS 'part_tot_cost', "
objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&cost_year&mm&"') AS 'as_tot_cnt', "
objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
objBuilder.Append "	AND cost_detail = '설치공사') AS 'set_tot_cost' "
objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '부문공통비' "

Set rsPart = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'부문공통비(배분)
as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS 총 건수
set_tot_cost = CDbl(f_toString(rsPart("set_tot_cost"), 0))	'설치공사 비용

rsPart.Close() : Set rsPart = Nothing

'A/S 장애처리 현황
objBuilder.Append "SELECT as_company, as_set, set_time, as_error, as_testing, as_collect, total_time, "
objBuilder.Append "	trade_name, saupbu, "
'objBuilder.Append	tot_part_cost&" / "&tot_part_cnt&" * as_total AS 'as_cost' "
objBuilder.Append	part_tot_cost&" / "&as_tot_cnt&" * (as_total - as_set) AS 'as_cost', "
objBuilder.Append set_tot_cost&" * set_time / (SELECT SUM(set_time) FROM as_acpt_status WHERE as_month = '"&cost_month&"' AND set_time > 0) AS 'as_set_cost' "
objBuilder.Append "FROM as_acpt_status AS aast "
objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
objBuilder.Append "	AND trdt.trade_id = '매출' "
objBuilder.Append "WHERE as_month = '"&cost_month&"' " & condi_sql

Set rsComCost = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>영업 관리 시스템</title>

		<style type="text/css">
			.first{
				text-align:center;
			}

			.right{
				text-align:right;
			}
		</style>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
						<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="15%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="10%" >
							<col width="15%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">거래처명</th>
								<!--<th scope="col">설치/공사</th>-->
								<th scope="col">설치/공사(시간)</th>
								<th scope="col">설치/공사비</th>
								<th scope="col">장애</th>
								<th scope="col">예방점검</th>
								<th scope="col">장비회수</th>
								<th scope="col">관리본부</th>
								<th scope="col">부문공통비</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim i, set_sum, set_time_sum, error_sum, testing_sum, collect_sum, total_time_sum, as_cost_sum
						Dim as_set_sum

						set_sum = 0
						set_time_sum = 0
						error_sum = 0
						testing_sum = 0
						collect_sum = 0
						total_time_sum = 0
						as_cost_sum = 0
						as_set_sum = 0

						i = 0

						Do Until  rsComCost.EOF
							i = i + 1

							'항목 별 총 합계
							set_sum = set_sum + CLng(rsComCost("as_set"))
							set_time_sum = set_time_sum + CLng(rsComCost("set_time"))
							error_sum = error_sum + CLng(rsComCost("as_error"))
							testing_sum = testing_sum + CLng(rsComCost("as_testing"))
							collect_sum = collect_sum + CLng(rsComCost("as_collect"))
							total_time_sum = total_time_sum + CLng(rsComCost("total_time"))

							as_cost_sum = as_cost_sum + CDbl(rsComCost("as_cost"))
							as_set_sum = as_set_sum + CDbl(rsComCost("as_set_cost"))
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=rsComCost("as_company")%></td>
								<!--<td class="right"><%=FormatNumber(rsComCost("as_set"), 0)%>&nbsp;</td>-->
								<td class="right"><%=FormatNumber(rsComCost("set_time"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_set_cost"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_error"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_testing"), 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(rsComCost("as_collect"), 0)%>&nbsp;</td>
								<td><%=rsComCost("saupbu")%></td>
								<td class="right"><%=FormatNumber(rsComCost("as_cost"), 0)%>&nbsp;</td>
							</tr>
						<%
							rsComCost.movenext()
						Loop

						rsComCost.Close() : Set rsComCost = Nothing
						DBConn.Close() : Set DBConn = Nothing

						Dim dist_part, dist_cost, part_cost

						'설치공사 비율 = 총 시간 / 설치공사 총 시간 * 100
						'dist_part = FormatNumber(set_time_sum / total_time_sum * 100, 1)

						'설치공사 비중  = 총 부문공통비 * 설치공사 비율
						'dist_cost = FormatNumber(as_cost_sum * dist_part / 100, 0)
						dist_cost = FormatNumber(set_tot_cost, 0)

						'부문공통비(배분) = 총 부문공통비 - 설치공사 비중
						'part_cost = FormatNumber(as_cost_sum - dist_cost, 0)
						part_cost = FormatNumber(tot_part_cost - dist_cost, 0)
						%>
							<tr>
								<td colspan="2" bgcolor="#FFE8E8" class="first">총계</td>
								<!--<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(set_sum, 0)%>&nbsp;건</td>-->
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(set_time_sum, 0)%>&nbsp;시간</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(as_set_sum, 0)%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(error_sum, 0)%>&nbsp;건</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(testing_sum, 0)%>&nbsp;건</td>
								<td bgcolor="#FFE8E8" class="right"><%=FormatNumber(collect_sum, 0)%>&nbsp;건</td>
								<td colspan="2" bgcolor="#FFE8E8" class="right">
									<div style="font-weight:bold;">설치공사 비중 : <%=dist_cost%></div>
									<div style="font-weight:bold;">부문공통비(배분) : <%=part_cost%></div>
								</td>
							</tr>
						</tbody>
					</table>
				<br>
		</div>
	</div>
	</body>
</html>