<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<script language="javascript" runat="server">
	function decodeUTF8(str){
		return decodeURIComponent(str);
	}

	function encodeUTF8(str) {
		return encodeURIComponent(str);
	}
</script>
<%
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
Dim org_company, end_month, end_yn, cost_year, cost_month
Dim from_date, end_date, to_date, start_date, rs_oil
Dim deptName, emp_msg, end_msg, oLoop, emp_cnt
Dim rsSalesOrg, arrSalesOrg

'param : inc_cost_end_oil.asp
Dim rsTran, oil_unit_id, liter, oil_unit_average, oil_price
Dim arrTran, i, mg_ce_id, run_date, run_seq, far

'param : inc_cost_end_person.asp
Dim rsOrgInfo, rs_gc, rs_ot, rs_tc, rs_ou
Dim rs_cs, rs_card, emp_end
Dim general_cnt, general_cost, general_pre_cnt, general_pre_cost
Dim overtime_cnt, overtime_cost
Dim gas_km, gas_unit, gas_cost, diesel_km
Dim diesel_unit, diesel_cost, gasol_km, gasol_unit
Dim gasol_cost, somopum_cost, fare_cnt, fare_cost
Dim oil_cash_cost, repair_cost, repair_pre_cost, parking_cost
Dim toll_cost, tot_km, tot_cost
Dim juyoo_card_cnt, juyoo_card_cost, juyoo_card_cost_vat, juyoo_card_price
Dim card_cnt, card_cost, card_cost_vat, card_price
Dim cash_tot_cost
Dim rs_car, car_owner, return_cash
Dim rs_person, variation_memo
Dim arrOrgInfo
Dim org_bonbu, org_saupbu, org_team, emp_reside_place, emp_reside_company
Dim emp_end_date, emp_name, emp_job
Dim arrGc, j, c_cnt, cost, pay_yn
Dim arrOt, cancel_yn
Dim arrTc, fare, oil_kind, parking, toll
Dim arrOu

'param : inc_cost_end_sum_insure.asp
Dim rsPay, rs_insure
Dim sort_seq, cost_detail
Dim insure_tot, income_tax, annual_pay, retire_pay
Dim cost_id
Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
Dim arrPay, pmg_id, base_pay, meals_pay, overtime_pay, research_pay, tax_no

'param : inc_cost_end_sum_bonus.asp
Dim rsBonus, arrBonus, rsAlba, arrAlba, company

'param : inc_cost_end_sum_cost.asp
Dim rsGeneral, rs_endGeneral, rsGeneralEnd, arrGeneralEnd, slip_date, slip_seq
Dim arrGeneral, account, rsEctCost, arrEtcCost

'param : inc_cost_end_sum_transit.asp
Dim rsTransit, rsRepair, arrTransit, arrRepair

'param : inc_cost_end_sum_card.asp
Dim rsCardTran, rsCardSlip, arrCardTran, arrCardSlip

org_company	=	Request("org_company")
end_month	=	Request("end_month")
end_yn		=	Request("end_yn")

cost_year 	= Mid(end_month, 1, 4)
cost_month 	= Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
start_date = DateAdd("m", -1, from_date)

'트랜잭션 시작
DBConn.BeginTrans

'유류비 단가 조회
objBuilder.Append "CALL USP_ORG_END_OIL_UNIT_SEL('"&end_month&"');"
Set rs_oil = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rs_oil.EOF Or rs_oil.BOF Then
	DBConn.RollbackTrans
	Response.Write "유류비 단가가 입력되어 있지 않아 마감을 할 수 없습니다."
	Response.End
End If
rs_oil.Close() : Set rs_oil = Nothing

'사업부 별 비용 마감
%>
<!--#include virtual="/cost_end_org.asp" -->
<%
'사업부 외 비용 마감
%>
<!--#include virtual="/cost_end_etc.asp" -->
<%
'상주 비용 마감

'공통비 마감


If emp_cnt = 0 Then
	'emp_msg = "인사마스터 마감이 되지 않았습니다."
	DBConn.RollbackTrans

	Response.Write "인사마스터 마감이 되지 않았습니다."
	Response.End
End If

If Err.Number <> 0 Then
	DBConn.RollbackTrans
	Response.Write "처리중 Error가 발생하였습니다."
Else
	DBConn.CommitTrans
	Response.Write "마감처리 되었습니다."
End If
Response.End

DBConn.Close() : Set DBConn = Nothing
%>