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
On Error Resume Next

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
Dim from_date, end_date, to_date, start_date
Dim rs_oil, deptName, emp_msg, end_msg, i, arrOil

org_company	=	decodeUTF8(f_Request("org_company"))
deptName		=	decodeUTF8(f_Request("saupbu"))	'사업 본부명으로 변경 사용
end_month	=	f_Request("end_month")
end_yn		=	f_Request("end_yn")

cost_year 	= Mid(end_month, 1, 4)
cost_month 	= Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
start_date = DateAdd("m", -1, from_date)

DBConn.BeginTrans

objBuilder.Append "SELECT oil_unit_id "
objBuilder.Append "FROM oil_unit "
objBuilder.Append "WHERE oil_unit_month = '"&end_month&"' "

Set rs_oil = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rs_oil.EOF Then
	arrOil = rs_oil.getRows()
End If
rs_oil.Close() : Set rs_oil = Nothing

'If rs_oil.EOF Or rs_oil.BOF Then
If Not IsArray(arrOil) Then
	Response.Write "유류비 단가가 입력되어 있지 않아 마감을 할 수 없습니다."
	Response.End
Else
	'유류비 단가 및 계산
	Dim rsTran, arrTran, rs_etc, rs_emp
	Dim oil_unit_id, liter, oil_unit_average, oil_price
	Dim mg_ce_id, oil_kind, far, run_date, run_seq, org_team
%>
	<!--#include virtual="/cost/inc/inc_cost_end_oil.asp" -->
<%
	'개인 경비 정산(교통비, 야특근, 카드)
	Dim rsOrgInfo, rs_gc, rs_ot, rs_tc, rs_ou, rs_cs, rs_card
	Dim emp_cnt, emp_end, overtime_cnt, overtime_cost
	Dim general_cnt, general_cost, general_pre_cnt, general_pre_cost
	Dim gas_km, gas_unit, gas_cost, diesel_km
	Dim diesel_unit, diesel_cost, gasol_km, gasol_unit
	Dim gasol_cost, somopum_cost, fare_cnt, fare_cost
	Dim oil_cash_cost, repair_cost, repair_pre_cost, parking_cost
	Dim toll_cost, tot_km, tot_cost
	Dim juyoo_card_cnt, juyoo_card_cost, juyoo_card_cost_vat, juyoo_card_price
	Dim card_cnt, card_cost, card_cost_vat, card_price
	Dim cash_tot_cost, rs_car, car_owner, return_cash, rs_person, variation_memo

	Dim arrOrgInfo, org_bonbu, org_saupbu, emp_reside_place, emp_reside_company
	Dim emp_end_date, emp_name, emp_job

	emp_cnt = 1
%>
	<!--#include virtual="/cost/inc/inc_cost_end_person.asp" -->
<%
	'월별 인사마스터 구성 여부 파악
	If emp_cnt > 0 Then
		'4대보험 및 급여 SUM 처리
		Dim rsPay, rs_insure, rs_payCost, rs_insureCost, rs_incomeCost, rs_annualCost, rs_retireCost
		Dim sort_seq, cost_detail
		Dim insure_tot, income_tax, annual_pay, retire_pay
		Dim cost_id
		Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
		Dim arrPay, pmg_id, base_pay, meals_pay, overtime_pay, research_pay, tax_no
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_insure.asp" -->
<%
		'상여/알바비 SUM 처리
		Dim rsBunus, rs_bonus, rsAlba, rs_alba
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_bonus.asp" -->
<%
		'DB SUM 일반 경비
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_cost.asp" -->
<%
		'DB SUM 교통비
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_transit.asp" -->
<%
		'카드비용 집계
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_card.asp" -->
<%
		'cost_end 테이블의 saupbu 컬럼을 본부명과 매칭 사용[허정호_20210312]
		If end_yn = "C" Then
			objBuilder.Append "UPDATE cost_end SET "
			objBuilder.Append "	end_yn = 'Y', mod_id = '"&user_id&"', mod_name = '"&user_name&"', mod_date = NOW() "
			objBuilder.Append "WHERE end_month = '"&end_month&"' "
			objBuilder.Append "	AND saupbu = '"&deptName&"' "
		Else
			objBuilder.Append "DELETE FROM cost_end "
			objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '"&deptName&"' "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			objBuilder.Append "INSERT INTO cost_end(end_month, saupbu, end_yn, batch_yn, bonbu_yn, ceo_yn, reg_id, reg_name, reg_date)"
			objBuilder.Append "VALUES("
			objBuilder.Append "'"&end_month&"', '"&deptName&"', 'Y', 'N', 'N', 'N', '"&user_id&"', '"&user_name&"', NOW()) "
		End If

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If
	' 월별 인사마스터 구성 여부 파악 END

	If emp_cnt = 0 Then
		emp_msg = "인사마스터 마감이 되지 않았습니다."
	Else
		emp_msg = ""
	End If

	If Err.Number <> 0 Then
		DBConn.RollbackTrans
		end_msg = emp_msg & "처리중 Error가 발생하였습니다."
	Else
		DBConn.CommitTrans
		end_msg = emp_msg & "마감처리 되었습니다."
	End If

	'Response.Write "<script type='text/javascript'>"
	'Response.Write "	alert('"&end_msg&"');"
	'Response.Write "	location.replace('/cost/cost_end_mg.asp');"
	'Response.Write "</script>"
	Response.write end_msg
	Response.End
End If

DBConn.Close() : Set DBConn = Nothing
%>