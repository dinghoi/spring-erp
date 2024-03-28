<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'/************************************************
' * ��븶��>���ֺ�� ���� ó��
' * 2017-09-13 add. ���� ���� ����
'************************************************
' * 1�� ����κ� mg_saupbu ó��
' * 2�� �μ��� mg_saupbu ó��
' * 3�� ����/�ι��� ���� mg_saupbu ó��
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

'��븶�� üũ
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
	Response.Write "��ü ��� ������ �Ǿ� ���� �ʽ��ϴ�."
	Response.End
End If

'Ʈ������ ����
DBConn.BeginTrans

' �λ縶���� �� �޿�DATA�� ��������� ����
Dim arrEmp, i, arrReside, emp_reside_company, emp_org_code
%>
<!--#include virtual="/cost_end/inc/inc_company_end_insa.asp" -->
<%
' �˹ٺ�� ��������� �� ������� ����
Dim rsAlba, rsAlbaOrg, rsAlbaOutCost, rsAlbaOutCostSales, rsAlbaOutCostTrade
Dim rsAlbaCost
Dim cost_center, cost_company, group_name, bill_trade_name, alba_bonbu
%>
<!--#include virtual="/cost_end/inc/inc_company_end_alba.asp" -->
<%
' �Ϲݺ�� ��������� �� ������� ����
Dim rsNoTax, rsNoTaxOrg, rsTax, rsTaxOrg, rsTaxNoMg, rsTaxNoMgOrg
%>
<!--#include virtual="/cost_end/inc/inc_company_end_cost.asp" -->
<%
' �Ϲݺ�� ��������� ����
Dim rsNoTaxOut, rsNoTaxOutSales, rsNoTaxOutTrade
Dim rsTaxCost, rsCompDeal, rsCompDealTrade
Dim cost_bonbu
%>
<!--#include virtual="/cost_end/inc/inc_company_end_general.asp" -->
<%
' �Ϲݺ�� ��������ο� ���� ��

' ī���� ��������� �� ������� ����
Dim rsCard, rsCardOrg, rsCardCost, rsCardOutCost, rsCardOutCostTrade
Dim deptName, rsCardReside
%>
<!--#include virtual="/cost_end/inc/inc_company_end_card.asp" -->
<%
' ī���� ��������� �� ������� ���� ��

' ���������� ������� ����
Dim rsTran, rsTranOrg, rsTranOutCost, rsTranOutCostOrg
Dim rsTranDeptOutCost, rsTranDeptOutCostTrade
Dim rsTranCost, tradeDept
%>
<!--#include virtual="/cost_end/inc/inc_company_end_set_transit.asp" -->
<%
' ��뱸�� Marking ����

'ȸ�� �� ��� ����(4�� ������ ��)
Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
Dim rsInsure, rsPaySum, rsPayTrade, rsPayCompOutCost
Dim insure_tot, income_tax, annual_pay, retire_pay
Dim rsInsureCost, rsIncomeCost, rsAnnualCost, rsRetireCost
Dim sort_seq, cost_detail
%>
<!--#include virtual="/cost_end/inc/inc_company_end_insure.asp" -->
<%
'ȸ�� �� ��� ����(�˹ٺ�)
Dim rsAlbaTot, rsAlbaTotTrade, rsAlbaCompanyCost
Dim sum_cost
%>
<!--#include virtual="/cost_end/inc/inc_company_end_alba.asp" -->
<%
' ��� SUM
Dim rsCostSum, rsCostSumTrade, rsCompanyCost
Dim cost_id
%>
<!--#include virtual="/cost_end/inc/inc_company_end_sum_cost.asp" -->
<%
' ��� SUM ����

' ī���� ����
Dim rsCardMg, rsCardMgTrade, rsCardCompanyCost
%>
<!--#include virtual="/cost_end/inc/inc_company_end_sum_card.asp" -->
<%
' ī���� ���� ��

' ���������� ����
Dim rsTranMg, rsTranMgTrade, rsTranCompanyCost
Dim rsRepair, rsRepairTrade, rsRepairCompanyCost
%>
<!--#include virtual="/cost_end/inc/inc_company_end_sum_transit.asp" -->
<%
' ���������� ���� ��

' ����κ�/ȸ�纰 ���� �ڷ� ����
Dim rsCostCompany, rsCostProfit, rsCompanyOutCost, rsProfitCostList
%>
<!--#include virtual="/cost_end/inc/inc_company_end_profit.asp" -->
<%

If end_yn = "C" Then
	'sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
	'"' and saupbu = '���ֺ��'"
	objBuilder.Append "UPDATE cost_end SET end_yn = 'Y', reg_id = '"&user_id&"', reg_name = '"&user_name&"', reg_date = NOW() "
	objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '���ֺ��'"
Else
	'sql="INSERT INTO cost_end(end_month, saupbu, end_yn, batch_yn, bonbu_yn, ceo_yn, reg_id, reg_name, reg_date)values('"&end_month& _
	'"','���ֺ��','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
	objBuilder.Append "INSERT INTO cost_end(end_month, saupbu, end_yn, batch_yn, bonbu_yn, ceo_yn, reg_id, reg_name, reg_date)VALUES("
	objBuilder.Append "'"&end_month&"', '���ֺ��', 'Y', 'N', 'N', 'N', '"&user_id&"', '"&user_name&"', NOW())"
End If
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Err.number <> 0 Then
	DBConn.RollbackTrans
	end_msg = emp_msg & "ó���� Error�� �߻��Ͽ����ϴ�."
Else
	DBConn.CommitTrans
	end_msg = emp_msg & "����ó�� �Ǿ����ϴ�."
End If

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('"&end_msg&"');"
Response.Write "	location.replace('/cost/cost_end_mg.asp');"
Response.Write "</script>"
Response.End

DBConn.Close() : Set DBConn = Nothing
%>
