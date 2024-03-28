<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'/************************************************
' * 비용마감>상주비용 마감 처리
' * 2017-09-13 add. 마감 로직 설명
'************************************************
' * 1차 사업부별 mg_saupbu 처리
' * 2차 부서별 mg_saupbu 처리
' * 3차 전사/부문에 따른 mg_saupbu 처리
'************************************************/
'On Error Resume Next

Server.ScriptTimeOut = 1200
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
Dim end_month, end_yn, cost_year, cost_month
Dim from_date, end_date, to_date, start_date
Dim reside_sw, rsCompanyEnd, arrCompanyEnd
Dim rsEmp, rsEmpSales, rsReside, rsResideTrade
Dim org_bonbu, org_code, trade_bonbu

end_month = f_Request("end_month")
end_yn = f_Request("end_yn")

cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)
from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
start_date = DateAdd("m", -1, from_date)

reside_sw = "Y"

'비용마감 체크
objBuilder.Append "CALL USP_COMPANY_END_COST_CNT('"&from_date&"', '"&to_date&"', '"&end_month&"');"
Set rsCompanyEnd = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCompanyEnd.EOF Then
	arrCompanyEnd = rsCompanyEnd.getRows()
End If
rsCompanyEnd.Close() : Set rsCompanyEnd = Nothing

If IsArray(arrCompanyEnd) Then
	reside_sw = arrCompanyEnd(0, 0)
Else
	reside_sw = "Y"
End If

If reside_sw = "N" Then
	Response.Write "전체 비용 마감이 되어 있지 않습니다."
	Response.End
End If

'트랜젝션 시작
DBConn.BeginTrans

' 인사마스터 및 급여DATA에 관리사업부 지정
Dim arrEmp, i, arrReside, emp_reside_company, emp_org_code
%>
<!--#include virtual="/cost_end/inc/inc_company_end_insa.asp" -->
<%
' 알바비용 관리사업부 및 비용유형 지정
Dim rsAlba, rsAlbaOrg, rsAlbaOutCost, rsAlbaOutCostSales, rsAlbaOutCostTrade
Dim rsAlbaCost
Dim cost_center, cost_company, group_name, bill_trade_name, alba_bonbu
%>
<!--#include virtual="/cost_end/inc/inc_company_end_alba.asp" -->
<%
' 일반비용 관리사업부 및 비용유형 지정
Dim rsNoTax, rsNoTaxOrg, rsTax, rsTaxOrg, rsTaxNoMg, rsTaxNoMgOrg
%>
<!--#include virtual="/cost_end/inc/inc_company_end_cost.asp" -->
<%
' 일반비용 관리사업부 지정
Dim rsNoTaxOut, rsNoTaxOutSales, rsNoTaxOutTrade
Dim rsTaxCost, rsCompDeal, rsCompDealTrade
Dim cost_bonbu
%>
<!--#include virtual="/cost_end/inc/inc_company_end_general.asp" -->
<%
' 일반비용 관리사업부와 지정 끝

' 카드사용 관리사업부 및 비용유형 지정
Dim rsCard, rsCardOrg, rsCardCost, rsCardOutCost, rsCardOutCostTrade
Dim deptName, rsCardReside
%>
<!--#include virtual="/cost_end/inc/inc_company_end_card.asp" -->
<%
' 카드사용 관리사업부 및 비용유형 지정 끝

' 차량관리비 비용유형 지정
Dim rsTran, rsTranOrg, rsTranOutCost, rsTranOutCostOrg
Dim rsTranDeptOutCost, rsTranDeptOutCostTrade
Dim rsTranCost, tradeDept
%>
<!--#include virtual="/cost_end/inc/inc_company_end_set_transit.asp" -->
<%
' 비용구분 Marking 종료

'회사 별 비용 마감(4대 보험율 등)
Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
Dim rsInsure, rsPaySum, rsPayTrade, rsPayCompOutCost
Dim insure_tot, income_tax, annual_pay, retire_pay
Dim rsInsureCost, rsIncomeCost, rsAnnualCost, rsRetireCost
Dim sort_seq, cost_detail
%>
<!--#include virtual="/cost_end/inc/inc_company_end_insure.asp" -->
<%
'회사 별 비용 마감(알바비)
Dim rsAlbaTot, rsAlbaTotTrade, rsAlbaCompanyCost
Dim sum_cost
%>
<!--#include virtual="/cost_end/inc/inc_company_end_alba.asp" -->
<%
' 비용 SUM
Dim rsCostSum, rsCostSumTrade, rsCompanyCost
Dim cost_id
%>
<!--#include virtual="/cost_end/inc/inc_company_end_sum_cost.asp" -->
<%
' 비용 SUM 종료

' 카드비용 집계
Dim rsCardMg, rsCardMgTrade, rsCardCompanyCost
%>
<!--#include virtual="/cost_end/inc/inc_company_end_sum_card.asp" -->
<%
' 카드비용 집계 끝

' 차량관리비 집계
Dim rsTranMg, rsTranMgTrade, rsTranCompanyCost
Dim rsRepair, rsRepairTrade, rsRepairCompanyCost
%>
<!--#include virtual="/cost_end/inc/inc_company_end_sum_transit.asp" -->
<%
' 차량관리비 집계 끝

' 사업부별/회사별 손익 자료 생성
Dim rsCostCompany, rsCostProfit, rsCompanyOutCost, rsProfitCostList
%>
<!--#include virtual="/cost_end/inc/inc_company_end_profit.asp" -->
<%

If end_yn = "C" Then
	'sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
	'"' and saupbu = '상주비용'"
	objBuilder.Append "UPDATE cost_end SET end_yn = 'Y', reg_id = '"&user_id&"', reg_name = '"&user_name&"', reg_date = NOW() "
	objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '상주비용'"
Else
	'sql="INSERT INTO cost_end(end_month, saupbu, end_yn, batch_yn, bonbu_yn, ceo_yn, reg_id, reg_name, reg_date)values('"&end_month& _
	'"','상주비용','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
	objBuilder.Append "INSERT INTO cost_end(end_month, saupbu, end_yn, batch_yn, bonbu_yn, ceo_yn, reg_id, reg_name, reg_date)VALUES("
	objBuilder.Append "'"&end_month&"', '상주비용', 'Y', 'N', 'N', 'N', '"&user_id&"', '"&user_name&"', NOW())"
End If
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = emp_msg & "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	end_msg = emp_msg & "마감처리 되었습니다."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('/cost/cost_end_mg.asp');"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>
