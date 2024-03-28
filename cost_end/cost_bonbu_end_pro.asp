<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
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
Dim rs_oil, deptName, emp_msg, end_msg

end_month = f_Request("end_month")
end_yn = f_Request("end_yn")

cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)
from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
start_date = DateAdd("m", -1, from_date)

'Ʈ����� ����
DBConn.BeginTrans

objBuilder.Append "CALL USP_ORG_END_OIL_UNIT_SEL('"&end_month&"');"
Set rs_oil = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rs_oil.EOF Or rs_oil.BOF Then
	DBConn.RollbackTrans
	Response.Write "������ �ܰ��� �ԷµǾ� ���� �ʾ� ������ �� �� �����ϴ�."
	Response.End
End If
rs_oil.Close() : Set rs_oil = Nothing

' ������ �ܰ� �� ������ ���
Dim rsTran, oil_unit_id, liter, oil_unit_average, oil_price
Dim arrTran, i, mg_ce_id, run_date, run_seq, far
%>
<!--#include virtual="/cost_end/inc/inc_bonbu_end_oil.asp" -->
<%
' ���κ� ��� ����
Dim rsOrgInfo, rs_gc, rs_ot, rs_tc, rs_ou
Dim rs_cs, rs_card, emp_cnt, emp_end
Dim general_cnt, general_cost, general_pre_cnt, general_pre_cost
Dim overtime_cnt, overtime_cost
Dim gas_km, gas_unit, gas_cost, diesel_km
Dim diesel_unit, diesel_cost, gasol_km, gasol_unit
Dim gasol_cost, somopum_cost, fare_cnt, fare_cost
Dim oil_cash_cost, repair_cost, repair_pre_cost, parking_cost
Dim toll_cost, tot_km, tot_cost
Dim juyoo_card_cnt, juyoo_card_cost, juyoo_card_cost_vat, juyoo_card_price
Dim card_cnt, card_cost, card_cost_vat, card_price
Dim cash_tot_cost, rs_car, car_owner, return_cash
Dim rs_person, variation_memo

Dim arrOrgInfo
Dim org_bonbu, org_saupbu, org_team, emp_reside_place, emp_reside_company
Dim emp_end_date, emp_name, emp_job

Dim arrGc, j, c_cnt, cost, pay_yn
Dim arrOt, cancel_yn, arrTc, fare, oil_kind, parking, toll, arrOu
%>
<!--#include virtual="/cost_end/inc/inc_bonbu_end_person.asp" -->
<%
' ���� �λ縶���� ���� ���� �ľ�
If emp_cnt > 0 Then
	'4�뺸�� �� �޿� SUM ó��
	Dim rsPay, rs_insure, sort_seq, cost_detail
	Dim insure_tot, income_tax, annual_pay, retire_pay, cost_id
	Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
	Dim arrPay, pmg_id, base_pay, meals_pay, overtime_pay, research_pay, tax_no
%>
	<!--#include virtual="/cost_end/inc/inc_bonbu_end_sum_insure.asp" -->
<%
	'��/�˹ٺ� SUM ó��
	Dim rsBonus, rsAlba, arrBonus, arrAlba, company
%>
	<!--#include virtual="/cost_end/inc/inc_bonbu_end_sum_bonus.asp" -->
<%
	'DB SUM �Ϲ� ���
	Dim rsGeneral, rsGeneralEnd, arrGeneralEnd, arrGeneral, rsEctCost, arrEtcCost
	Dim slip_date, slip_seq, account
%>
	<!--#include virtual="/cost_end/inc/inc_bonbu_end_sum_cost.asp" -->
<%
	'DB SUM �����
	Dim rsTransit, arrTransit, rsRepair, arrRepair
%>
	<!--#include virtual="/cost_end/inc/inc_bonbu_end_sum_transit.asp" -->
<%
	'ī���� ����
	Dim rsCardTran, arrCardTran, rsCardSlip, arrCardSlip
%>
	<!--#include virtual="/cost_end/inc/inc_bonbu_end_sum_card.asp" -->
<%
	objBuilder.Append "CALL USP_ORG_END_PROC('"&end_month&"', '����οܳ�����', '"&end_yn&"', '"&user_id&"', '"&user_name&"');"
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If
' ���� �λ縶���� ���� ���� �ľ� END

If emp_cnt = 0 Then
	DBConn.RollbackTrans
	Response.Write "�λ縶���� ������ ���� �ʾҽ��ϴ�."
	Response.End
End If

If Err.number <> 0 Then
	DBConn.RollbackTrans
	Response.Write "ó���� Error�� �߻��Ͽ����ϴ�."
Else
	DBConn.CommitTrans
	Response.Write "���� ó�� �Ǿ����ϴ�."
End If
Response.End
DBConn.Close() : Set DBConn = Nothing
%>


