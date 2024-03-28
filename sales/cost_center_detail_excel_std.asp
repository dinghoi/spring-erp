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
Dim cost_month, sales_saupbu, slip_month, title_line, savefilename, i
Dim rs_etc, insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per

cost_month = f_Request("cost_month")
sales_saupbu = f_Request("sales_saupbu")

'If sales_saupbu = "기타사업부" Then
'	sales_saupbu = ""
'End If

slip_month = mid(cost_month,1,4) & "-" & mid(cost_month,5,2)

title_line = cost_month & "월 " & sales_saupbu & " 비용세부 내역(표준)"
savefilename = title_line & ".xls"

'엑셀 다운로드 설정
Call ViewExcelType(savefilename)

objBuilder.Append "SELECT insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per "
objBuilder.Append "FROM insure_per WHERE insure_year = '"&Mid(cost_month, 1, 4)&"' "

Set rs_etc = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

insure_tot_per = rs_etc("insure_tot_per")
income_tax_per = rs_etc("income_tax_per")
annual_pay_per = rs_etc("annual_pay_per")
retire_pay_per = rs_etc("retire_pay_per")

rs_etc.Close() : Set rs_etc = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">비용구분</th>
								<th scope="col">세부유형</th>
								<th scope="col">비용유형</th>
								<th scope="col">담당영업사업부</th>
								<th scope="col">계산서유무</th>
								<th scope="col">비용회사</th>
								<th scope="col">본부</th>
								<th scope="col">사업부</th>
								<th scope="col">팀</th>
								<th scope="col">조직명</th>
								<th scope="col">상주처</th>
								<th scope="col">고객사</th>
								<th scope="col">담당자</th>
								<th scope="col">발행일자</th>
								<th scope="col">발행순번</th>
								<th scope="col">외주업체</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">발행내역</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim rsPay, arrPay, cost_center, mg_saupbu, pmg_id, pmg_emp_name, pmg_yymm, tax_bill_yn
						Dim gubun, account, pmg_emp_no, pmg_give_total, org_company, org_bonbu, org_saupbu, org_team
						Dim emp_org_name, emp_reside_place, emp_reside_company, customer, cost_vat, slip_memo, num

						If (saupbu = sales_saupbu And position = "사업부장") Or (saupbu = sales_saupbu And position = "본부장") Or sales_grade = "0" Then
							' 인건비 > 급여
                            objBuilder.Append "SELECT pmgt.cost_center, pmgt.mg_saupbu, pmgt.pmg_id, "
                            objBuilder.Append "   pmgt.pmg_emp_name, pmgt.pmg_yymm,pmgt.pmg_emp_no, pmgt.pmg_give_total, "
                            objBuilder.Append "   eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
                            'SQL = SQL & "   eomt.org_name, eomt.org_reside_place, eomt.org_reside_company "
							objBuilder.Append "   emmt.emp_org_name, emmt.emp_reside_place, emmt.emp_reside_company "
                            objBuilder.Append "FROM pay_month_give AS pmgt "
                            objBuilder.Append "LEFT OUTER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
                            objBuilder.Append "   AND emmt.emp_month = '"&cost_month&"' "
                            objBuilder.Append "LEFT OUTER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
							objBuilder.Append "	AND eomt.org_month = '"&cost_month&"' "
                            objBuilder.Append "WHERE pmgt.pmg_id <>'4' AND pmgt.pmg_yymm = '"&cost_month&"' AND pmgt.cost_center <> '손익제외' "

							If sales_saupbu = "전사공통비" Or sales_saupbu = "부문공통비" Then
                                objBuilder.Append "   AND pmgt.cost_center ='"&sales_saupbu&"' "
							Else
                                objBuilder.Append "   AND (pmgt.cost_center ='직접비' OR pmgt.cost_center ='상주직접비') "

								'SQL = SQL & "	AND pmgt.mg_saupbu ='"&sales_saupbu&"' "
								If sales_saupbu = "기타사업부" Then
									objBuilder.Append "	AND pmgt.mg_saupbu IN ('"&sales_saupbu&"', '') "
								Else
									objBuilder.Append "	AND pmgt.mg_saupbu = '"&sales_saupbu&"' "
								End If
							End If
                            objBuilder.Append "ORDER BY pmgt.pmg_id, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
                            'SQL = SQL & "   eomt.org_name, eomt.org_reside_place, eomt.org_reside_company, "
							objBuilder.Append "   emmt.emp_org_name, emmt.emp_reside_place, emmt.emp_reside_company, "
							objBuilder.Append "	pmgt.pmg_emp_name "

							Set rsPay = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							If Not rsPay.EOF Then
								arrPay = rsPay.getRows()
							End If
							rsPay.Close() : Set rsPay = Nothing

							If IsArray(arrPay) Then
								tax_bill_yn = "일반"
								gubun = "인건비"
								account = "미지정"
								num = 0

								For i = LBound(arrPay) To UBound(arrPay, 2)
									cost_center = arrPay(0, i)
									mg_saupbu = arrPay(1, i)
									pmg_id = arrPay(2, i)
									pmg_emp_name = arrPay(3, i)
									pmg_yymm = arrPay(4, i)
									pmg_emp_no = arrPay(5, i)
									pmg_give_total = arrPay(6, i)
									org_company = arrPay(7, i)
									org_bonbu = arrPay(8, i)
									org_saupbu = arrPay(9, i)
									org_team = arrPay(10, i)
									emp_org_name = arrPay(11, i)
									emp_reside_place = arrPay(12, i)
									emp_reside_company = arrPay(13, i)

									Select Case pmg_id
										Case "1" : account = "급여"
										Case "2" : account = "상여"
										Case "3" : account = "추천인센티브"
									End Select

									customer     = ""
									cost_vat     = 0
									slip_memo    = ""

									num = num + 1
						%>
							<tr>
								<td class="first"><%=num%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=org_company%></td>
								<td><%=org_bonbu%></td>
								<td><%=org_saupbu%></td>
								<td><%=org_team%></td>
								<td><%=org_name%></td>
								<td><%=emp_reside_place%></td>
								<td><%=emp_reside_company%></td>
								<td><%=pmg_emp_name%></td>
								<td><%=pmg_yymm%></td>
								<td><%=pmg_emp_no%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=FormatNumber(pmg_give_total, 0)%></td>
							  	<td class="right"><%=FormatNumber(pmg_give_total, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost_vat, 0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
								Next
							End If

							' 인건비 > 4대보험, 소득세종업원분, 연차수당, 퇴직충당금
							Dim rsInsure, arrInsure, insure_tot, income_tax, annual_pay, retire_pay
							Dim tot_cost, base_pay, meals_pay, overtime_pay, tax_no

							objBuilder.Append "SELECT pmgt.cost_center, emmt.emp_company, SUM(pmgt.pmg_give_total) AS tot_cost, "
							objBuilder.Append "	SUM(pmg_base_pay) AS base_pay, SUM(pmg_meals_pay) AS meals_pay, "
							objBuilder.Append "	SUM(pmg_overtime_pay) AS overtime_pay, SUM(pmg_tax_no) AS tax_no "
							objBuilder.Append "FROM pay_month_give AS pmgt "
							objBuilder.Append "LEFT OUTER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
							objBuilder.Append "	AND emmt.emp_month = '"&cost_month&"' "
							objBuilder.Append "WHERE pmgt.pmg_id = '1' AND pmgt.pmg_yymm = '"&cost_month&"' "
							objBuilder.Append "	AND pmgt.cost_center <> '손익제외' "

							If sales_saupbu = "전사공통비" Or sales_saupbu = "부문공통비" Then
								objBuilder.Append "AND pmgt.cost_center ='"&sales_saupbu&"' "
							Else
								objBuilder.Append "AND (pmgt.cost_center ='직접비' OR pmgt.cost_center ='상주직접비') "
								'SQL = SQL & "AND pmgt.mg_saupbu ='"&sales_saupbu&"' "
								If sales_saupbu = "기타사업부" Then
									objBuilder.Append "	AND pmgt.mg_saupbu IN ('"&sales_saupbu&"', '') "
								Else
									objBuilder.Append "	AND pmgt.mg_saupbu = '"&sales_saupbu&"' "
								End If
							End If
							objBuilder.Append "GROUP BY pmgt.pmg_id, pmgt.cost_center, emmt.emp_company "
							objBuilder.Append "ORDER BY emmt.emp_company "

							Set rsInsure = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							If Not rsInsure.EOF Then
								arrInsure = rsInsure.getRows()
							End If
							rsInsure.Close() : Set rsInsure = Nothing

							If IsArray(arrInsure) Then
								For i = LBound(arrInsure) To UBound(arrInsure, 2)
									cost_center = arrInsure(0, i)
									emp_company = arrInsure(1, i)
									tot_cost = arrInsure(2, i)
									base_pay = arrInsure(3, i)
									meals_pay = arrInsure(4, i)
									overtime_pay = arrInsure(5, i)
									tax_no = arrInsure(6, i)

									'insure_tot = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * insure_tot_per / 100)
									insure_tot = CLng((CLng(tot_cost)) * insure_tot_per / 100)
									'income_tax = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * income_tax_per / 100)
									income_tax = CLng((CLng(tot_cost)) * income_tax_per / 100)
									annual_pay = CLng((CLng(base_pay) + CLng(meals_pay) + CLng(overtime_pay)) * annual_pay_per / 100)
									retire_pay = CLng((CLng(base_pay) + CLng(meals_pay) + CLng(overtime_pay)) * retire_pay_per / 100)

									num = num + 1
						%>
							<tr>
								<td class="first"><%=num%></td>
								<td>인건비</td>
								<td>4대보험료</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>일반</td>
								<td><%=emp_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=FormatNumber(insure_tot, 0)%></td>
							  	<td class="right"><%=FormatNumber(insure_tot, 0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<%num = num + 1%>
							<tr>
								<td class="first"><%=num%></td>
								<td>인건비</td>
								<td>소득세종업원분</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>일반</td>
								<td><%=emp_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=FormatNumber(income_tax, 0)%></td>
							  	<td class="right"><%=FormatNumber(income_tax, 0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<%num = num + 1%>
							<tr>
								<td class="first"><%=num%></td>
								<td>인건비</td>
								<td>연차수당</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>일반</td>
								<td><%=emp_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=FormatNumber(annual_pay, 0)%></td>
							  	<td class="right"><%=FormatNumber(annual_pay, 0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<%num = num + 1%>
							<tr>
								<td class="first"><%=num%></td>
								<td>인건비</td>
								<td>퇴직충당금</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>일반</td>
								<td><%=emp_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=FormatNumber(retire_pay, 0)%></td>
							  	<td class="right"><%=FormatNumber(retire_pay, 0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
						<%
								Next
							End If
						End If

						Dim rsComCost, tot_part_cost, rsAsSum, as_set_sum, set_time_sum, total_time_sum
						Dim dist_part, dist_cost, rsAsTot, arrAsTot, as_company, as_cost

						'인건비 > 설치공사
						'전체 부문공통비
						objBuilder.Append "SELECT SUM(cost_amt_"& Mid(cost_month, 5, 7) &") AS tot_cost "
						objBuilder.Append "FROM company_cost "
						objBuilder.Append "WHERE cost_year ='"& Mid(cost_month, 1, 4) &"' "
						objBuilder.Append "	AND cost_center = '부문공통비' "

						Set rsComCost = DbConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If IsNull(rsComCost("tot_cost")) Then
							tot_part_cost = 0
						Else
							tot_part_cost = CLng(rsComCost("tot_cost"))
						End If
						rsComCost.Close() : Set rsComCost = Nothing

						'AS 현황 집계
						objBuilder.Append "SELECT SUM(as_set) AS 'as_set_sum', SUM(set_time) AS 'set_time_sum', SUM(total_time) AS 'total_time_sum' "
						objBuilder.Append "FROM as_acpt_status "
						objBuilder.Append "WHERE as_month = '"&cost_month&"' "

						Set rsAsSum = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						as_set_sum = CLng(rsAsSum("as_set_sum"))	'설치공사 총 건수
						set_time_sum = CLng(rsAsSum("set_time_sum"))	'설치공사 총 시간
						total_time_sum = CLng(rsAsSum("total_time_sum")) '총 시간

						rsAsSum.Close() : Set rsAsSum = Nothing

						'설치공사 비율 = 설치공사 총 시간 / 총 시간 * 100
						dist_part = FormatNumber(set_time_sum / total_time_sum * 100, 1)

						'설치공사 비중  = 총 부문공통비 * 설치공사 비율(%)
						dist_cost = CDbl(FormatNumber(tot_part_cost * dist_part / 100, 0))

						'AS 현황 > 설치/공사 조회
						objBuilder.Append "SELECT as_company, "

						'부문공통비일 경우 마이너스 처리
						If sales_saupbu = "부문공통비" Then
							objBuilder.Append "	("&dist_cost&" / "&as_set_sum&" * as_set *-1) AS 'as_cost' "
						Else
							objBuilder.Append "	("&dist_cost&" / "&as_set_sum&" * as_set) AS 'as_cost' "
						End If

						objBuilder.Append "FROM as_acpt_status AS aast "
						objBuilder.Append "INNER JOIN trade AS trat ON aast.as_company = trat.trade_name AND trade_id = '매출' "
						objBuilder.Append "WHERE as_month ='"&cost_month&"' AND as_set > 0 "

						Select Case sales_saupbu
							Case "부문공통비"
								objBuilder.Append ""
							Case "전사공통비"
								objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
							Case Else
								objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
						End Select

						Set rsAsTot = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsAsTot.EOF Then
							arrAsTot = rsAsTot.getRows()
						End If
						rsAsTot.Close() : Set rsAsTot = Nothing

						If IsArray(arrAsTot) Then
							For i = LBound(arrAsTot) To UBound(arrAsTot, 2)
								as_company = arrAsTot(0, i)
								as_cost = arrAsTot(1, i)

								num = num + 1
						%>
						<tr>
							<td class="first"><%=num%></td>
							<td>인건비</td>
							<td>설치공사</td>
							<td>상주직접비</td>
							<td></td>
							<td>일반</td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td><%=as_company%></td>
							<td></td>
							<td><%=cost_month%></td>
							<td></td>
							<td></td>
							<td class="right">0</td>
							<td class="right"><%=FormatNumber(as_cost, 0)%></td>
							<td class="right">0</td>
							<td></td>
						</tr>
						<%
							Next
						End If

						Dim rsCowork, arrCowork, cw_company, cw_cost

						'인건비 > 협업
						objBuilder.Append "SELECT as_company, SUM(as_give_cowork * 30000 * -1) + SUM(as_get_cowork * 30000) AS 'cowork_cost' "
						objBuilder.Append "FROM as_acpt_status AS aast "
						objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
						objBuilder.Append "WHERE as_month = '"&cost_month&"' AND trdt.saupbu = '"&sales_saupbu&"' "
						objBuilder.Append "	AND ((as_give_cowork > 0 OR as_give_cowork < 0) OR (as_get_cowork > 0 OR as_get_cowork < 0)) "
						objBuilder.Append "GROUP BY as_company "

						Set rsCowork = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsCowork.EOF Then
							arrCowork = rsCowork.getRows()
						End If
						rsCowork.Close() : Set rsCowork = Nothing

						If IsArray(arrCowork) Then
							For i = LBound(arrCowork) To UBound(arrCowork, 2)
								cw_company = arrCowork(0, i)
								cw_cost = arrCowork(1, i)

								num = num + 1
						%>
						<tr>
							<td class="first"><%=num%></td>
							<td>인건비</td>
							<td>협업</td>
							<td>상주직접비</td>
							<td></td>
							<td>일반</td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td></td>
							<td><%=cw_company%></td>
							<td></td>
							<td><%=cost_month%></td>
							<td></td>
							<td></td>
							<td class="right">0</td>
							<td class="right"><%=FormatNumber(cw_cost, 0)%></td>
							<td class="right">0</td>
							<td></td>
						</tr>
						<%
							Next
						End If

						'인건비 > 알바비
						Dim rsAlba, arrAlba, company, cost_company, draft_man, give_date, draft_no
						Dim alba_give_total, draft_tax_id

						objBuilder.Append "SELECT cost_center, mg_saupbu, company, bonbu, saupbu, team, org_name, cost_company, "
						objBuilder.Append "	draft_man, give_date, draft_no, alba_give_total, alba_give_total, draft_tax_id "
						objBuilder.Append "FROM pay_alba_cost "
						objBuilder.Append "WHERE rever_yymm = '"&cost_month&"' "

						If sales_saupbu = "전사공통비" Or sales_saupbu = "부문공통비" Then
							objBuilder.Append "	AND cost_center ='"&sales_saupbu&"' "
						Else
							objBuilder.Append "	AND (cost_center ='직접비' OR cost_center ='상주직접비') "
							'SQL = SQL & "	AND mg_saupbu ='"&sales_saupbu&"' "
							If sales_saupbu = "기타사업부" Then
								objBuilder.Append "	AND mg_saupbu IN ('"&sales_saupbu&"', '') "
							Else
								objBuilder.Append "	AND mg_saupbu = '"&sales_saupbu&"' "
							End If
						End If
						objBuilder.Append "ORDER BY cost_center,give_date,mg_saupbu,org_name, draft_man "

						Set rsAlba = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsAlba.EOF Then
							arrAlba = rsAlba.getRows()
						End If
						rsAlba.Close() : Set rsAlba = Nothing

						If IsArray(arrAlba) Then
							For i = LBound(arrAlba) To UBound(arrAlba, 2)
								cost_center = arrAlba(0, i)
								mg_saupbu = arrAlba(1, i)
								company = arrAlba(2, i)
								bonbu = arrAlba(3, i)
								saupbu = arrAlba(4, i)
								team = arrAlba(5, i)
								org_name = arrAlba(6, i)
								cost_company = arrAlba(7, i)
								draft_man = arrAlba(8, i)
								give_date = arrAlba(9, i)
								draft_no = arrAlba(10, i)
								alba_give_total = arrAlba(11, i)
								draft_tax_id = arrAlba(12, i)

								tax_bill_yn = "일반"
								gubun = "인건비"
								account = "알바비"
								reside_place = ""
								customer = ""
								cost_vat = 0
								num = num + 1
						%>
							<tr>
								<td class="first"><%=num%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=cost_company%></td>
								<td><%=draft_man%></td>
								<td><%=give_date%></td>
								<td><%=draft_no%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=FormatNumber(alba_give_total, 0)%></td>
							  	<td class="right"><%=FormatNumber(alba_give_total, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost_vat, 0)%></td>
								<td><%=draft_tax_id%></td>
							</tr>
						<%
							Next
						End If

						'일반경비 > 세금계산서
						Dim rsTax, arrTax, slip_gubun, emp_name, slip_date, slip_seq, price, cost

						objBuilder.Append "SELECT slip_gubun, account,cost_center, mg_saupbu, emp_company, bonbu, saupbu, team, "
						objBuilder.Append "	org_name, reside_place, company, emp_name, slip_date, slip_seq, customer, price, "
						objBuilder.Append "	cost, cost_vat, slip_memo, tax_bill_yn "
						objBuilder.Append "FROM general_cost "
						objBuilder.Append "WHERE pl_yn = 'Y' AND cancel_yn ='N' "
						objBuilder.Append "	and substring(slip_date,1,7) = '"&slip_month&"' "

						If sales_saupbu = "전사공통비" Or sales_saupbu = "부문공통비" Then
							objBuilder.Append "	AND cost_center ='"&sales_saupbu&"' "
						Else
							objBuilder.Append "	AND (cost_center ='직접비' OR cost_center ='상주직접비') "
							'SQL = SQL & "	AND mg_saupbu ='"&sales_saupbu&"' "
							If sales_saupbu = "기타사업부" Then
								objBuilder.Append "	AND mg_saupbu IN ('"&sales_saupbu&"', '') "
							Else
								objBuilder.Append "	AND mg_saupbu = '"&sales_saupbu&"' "
							End If
						End If
						objBuilder.Append "ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"

						Set rsTax = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsTax.EOF Then
							arrTax = rsTax.getRows()
						End If
						rsTax.Close() : Set rsTax = Nothing

						If IsArray(arrTax) Then
							For i = LBound(arrTax) To UBound(arrTax, 2)
								slip_gubun = arrTax(0, i)
								account = arrTax(1, i)
								cost_center = arrTax(2, i)
								mg_saupbu = arrTax(3, i)
								emp_company = arrTax(4, i)
								bonbu = arrTax(5, i)
								saupbu = arrTax(6, i)
								team = arrTax(7, i)
								org_name = arrTax(8, i)
								reside_place = arrTax(9, i)
								company = arrTax(10, i)
								emp_name = arrTax(11, i)
								slip_date = arrTax(12, i)
								slip_seq = arrTax(13, i)
								customer = arrTax(14, i)
								price = arrTax(15, i)
								cost = arrTax(16, i)
								cost_vat = arrTax(17, i)
								slip_memo = arrTax(18, i)
								tax_bill_yn = arrTax(19, i)

								If tax_bill_yn = "Y" Then
									tax_bill_yn = "세금계산서"
								Else
									tax_bill_yn = "일반"
								End If

								num = num + 1
						%>
							<tr>
								<td class="first"><%=num%></td>
								<td><%=slip_gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=FormatNumber(price, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost_vat, 0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							Next
						End If

						'일반경비 > 교통비
						Dim rsTran, arrTran, car_owner, run_date, run_seq, somopum, oil_price, fare, parking, toll
						Dim run_memo

						objBuilder.Append "SELECT car_owner, cost_center, mg_saupbu, emp_company, bonbu, saupbu, team, org_name, "
						objBuilder.Append "	reside_place, company, user_name, run_date, run_seq, somopum, oil_price, fare, "
						objBuilder.Append "	parking, toll, run_memo "
						objBuilder.Append "FROM transit_cost "
						objBuilder.Append "WHERE cancel_yn ='N' AND SUBSTRING(run_date, 1, 7) = '"&slip_month&"'  "

						If sales_saupbu = "전사공통비" Or sales_saupbu = "부문공통비" Then
							objBuilder.Append "	AND cost_center ='"&sales_saupbu&"' "
						Else
							objBuilder.Append "	AND (cost_center ='직접비' OR cost_center ='상주직접비') "
							'SQL = SQL & "	and mg_saupbu ='"&sales_saupbu&"' "
							If sales_saupbu = "기타사업부" Then
								objBuilder.Append "	AND mg_saupbu IN ('"&sales_saupbu&"', '') "
							Else
								objBuilder.Append "	AND mg_saupbu = '"&sales_saupbu&"' "
							End If
						End If
						objBuilder.Append "ORDER BY cost_center, run_date, mg_saupbu, org_name, user_name "

						Set rsTran = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsTran.EOF Then
							arrTran = rsTran.getRows()
						End If
						rsTran.Close() : Set rsTran = Nothing

						If IsArray(arrTran) Then
							For i = LBound(arrTran) To UBound(arrTran, 2)
								car_owner = arrTran(0, i)
								cost_center = arrTran(1, i)
								mg_saupbu = arrTran(2, i)
								emp_company = arrTran(3, i)
								bonbu = arrTran(4, i)
								saupbu = arrTran(5, i)
								team = arrTran(6, i)
								org_name = arrTran(7, i)
								reside_place = arrTran(8, i)
								company = arrTran(9, i)
								user_name = arrTran(10, i)
								run_date = arrTran(11, i)
								run_seq = arrTran(12, i)
								somopum = arrTran(13, i)
								oil_price = arrTran(14, i)
								fare = arrTran(15, i)
								parking = arrTran(16, i)
								toll = arrTran(17, i)
								run_memo = arrTran(18, i)

								price = somopum + oil_price + fare + parking + toll
								cost = somopum + oil_price + fare + parking + toll

								tax_bill_yn = "일반"
								gubun = "교통비"
								customer = ""
								cost_vat = 0
								num = num + 1
						%>
							<tr>
								<td class="first"><%=num%></td>
								<td><%=gubun%></td>
								<td><%=car_owner%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=user_name%></td>
								<td><%=run_date%></td>
								<td><%=run_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=FormatNumber(price, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost_vat, 0)%></td>
								<td><%=run_memo%></td>
							</tr>
						<%
							Next
						End If

						'일반경비 > 교통비 > 차량수리비
						Dim rsRepair, arrRepair, repair_cost

						objBuilder.Append "SELECT trct.cost_center, trct.mg_saupbu, eomt.org_company, eomt.org_bonbu, "
						objBuilder.Append "	eomt.org_saupbu, eomt.org_team, eomt.org_name, "
						objBuilder.Append "	trct.reside_place, trct.company, trct.user_name, trct.run_date, trct.run_seq, "
						objBuilder.Append "	trct.repair_cost, trct.run_memo "
						objBuilder.Append "FROM transit_cost AS trct "
						objBuilder.Append "INNER JOIN emp_master_month AS emmt ON trct.mg_ce_id = emmt.emp_no "
						objBuilder.Append "	 AND emmt.emp_month = '"&cost_month&"' "
						objBuilder.Append "INNER JOIN emp_org_mst AS eomt ON emmt.emp_org_code = eomt.org_code "
						objBuilder.Append "WHERE cancel_yn ='N' AND repair_cost > 0 "
						objBuilder.Append "	AND SUBSTRING(run_date, 1, 7) = '"&slip_month&"' "

						If sales_saupbu = "전사공통비" Or sales_saupbu = "부문공통비" Then
							objBuilder.Append "	AND trct.cost_center ='"&sales_saupbu&"' "
						Else
							objBuilder.Append "	AND (trct.cost_center ='직접비' OR trct.cost_center ='상주직접비') "
							'SQL = SQL & "	AND emmt.mg_saupbu ='"&sales_saupbu&"' "
							If sales_saupbu = "기타사업부" Then
								objBuilder.Append "	AND trct.mg_saupbu IN ('"&sales_saupbu&"', '') "
							Else
								objBuilder.Append "	AND trct.mg_saupbu = '"&sales_saupbu&"' "
							End If
						End If
						objBuilder.Append "ORDER BY trct.cost_center, run_date, trct.mg_saupbu, trct.org_name, user_name "

						Set rsRepair = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsRepair.EOF Then
							arrRepair = rsRepair.getRows()
						End If
						rsRepair.Close() : Set rsRepair = Nothing

						If IsArray(arrRepair) Then
							For i = LBound(arrRepair) To UBound(arrRepair, 2)
								cost_center = arrRepair(0, i)
								mg_saupbu = arrRepair(1, i)
								org_company = arrRepair(2, i)
								org_bonbu = arrRepair(3, i)
								org_saupbu = arrRepair(4, i)
								org_team = arrRepair(5, i)
								org_name = arrRepair(6, i)
								reside_place = arrRepair(7, i)
								company = arrRepair(8, i)
								user_name = arrRepair(9, i)
								run_date = arrRepair(10, i)
								run_seq = arrRepair(11, i)
								repair_cost = arrRepair(12, i)
								run_memo = arrRepair(13, i)

								tax_bill_yn  = "일반"
								gubun        = "교통비"
								account      = "차량수리비"
								customer     = ""
								cost_vat     = 0

								num = num + 1
						%>
							<tr>
								<td class="first"><%=num%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=org_company%></td>
								<td><%=org_bonbu%></td>
								<td><%=org_saupbu%></td>
								<td><%=org_team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=user_name%></td>
								<td><%=run_date%></td>
								<td><%=run_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=FormatNumber(repair_cost, 0)%></td>
							  	<td class="right"><%=FormatNumber(repair_cost, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost_vat, 0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							Next
						End If

						'법인카드
						Dim rsCard, arrCard, approve_no, account_item

						objBuilder.Append "SELECT cslt.account, cslt.cost_center, cslt.mg_saupbu, cslt.emp_company, cslt.bonbu, cslt.saupbu, cslt.team, "
						objBuilder.Append "	cslt.org_name, cslt.reside_place, cslt.reside_company, cslt.emp_name, cslt.slip_date, cslt.approve_no, "
						objBuilder.Append "	cslt.customer, cslt.price, cslt.cost, cslt.cost_vat, cslt.account_item "
						objBuilder.Append "FROM card_slip AS cslt "
						objBuilder.Append "INNER JOIN emp_master_month AS emmt ON cslt.emp_no = emmt.emp_no "
						objBuilder.Append "	AND emmt.emp_month = '"&cost_month&"' "
						objBuilder.Append "WHERE pl_yn = 'Y' AND (card_type NOT LIKE '%주유%' OR com_drv_yn = 'Y') "
						objBuilder.Append "	AND SUBSTRING(slip_date, 1, 7) = '"&slip_month&"' "

						If sales_saupbu = "전사공통비" Or sales_saupbu = "부문공통비" Then
							objBuilder.Append "	AND cslt.cost_center = '"&sales_saupbu&"' "
						Else
							objBuilder.Append "	and (cslt.cost_center ='직접비' or cslt.cost_center ='상주직접비') "
							'SQL = SQL & "	and cslt.mg_saupbu ='"&sales_saupbu&"' "
							If sales_saupbu = "기타사업부" Then
								objBuilder.Append "	AND cslt.mg_saupbu IN ('"&sales_saupbu&"', '') "
							Else
								objBuilder.Append "	AND cslt.mg_saupbu = '"&sales_saupbu&"' "
							End If
						End If

						objBuilder.Append "ORDER BY emmt.cost_center, slip_date, emmt.mg_saupbu, org_name, emp_name "

						' and mg_saupbu ='"&sales_saupbu&"' and
						'where (pl_yn = 'Y') and (card_type not like '%주유%' or com_drv_yn = 'Y') and

						Set rsCard = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsCard.EOF Then
							arrCard = rsCard.getRows()
						End If
						rsCard.Close() : Set rsCard = Nothing

						If IsArray(arrCard) Then
							For i = LBound(arrCard) To UBound(arrCard, 2)
								account = arrCard(0, i)
								cost_center = arrCard(1, i)
								mg_saupbu = arrCard(2, i)
								emp_company = arrCard(3, i)
								bonbu = arrCard(4, i)
								saupbu = arrCard(5, i)
								team = arrCard(6, i)
								org_name = arrCard(7, i)
								reside_place = arrCard(8, i)
								reside_company = arrCard(9, i)
								emp_name = arrCard(10, i)
								slip_date = arrCard(11, i)
								approve_no = arrCard(12, i)
								customer = arrCard(13, i)
								price = arrCard(14, i)
								cost = arrCard(15, i)
								cost_vat = arrCard(16, i)
								account_item = arrCard(17, i)

								tax_bill_yn   = "일반"
								gubun         = "법인카드"
								num = num + 1
						%>
							<tr>
								<td class="first"><%=num%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=reside_company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=approve_no%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=FormatNumber(price, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost, 0)%></td>
							  	<td class="right"><%=FormatNumber(cost_vat, 0)%></td>
								<td><%=account_item%></td>
							</tr>
						<%
							Next
						End If
						%>
						</tbody>
					</table>
				</div>
			</div>
		</div>
	</body>
</html>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
DBConn.Close : Set DBConn = Nothing
%>