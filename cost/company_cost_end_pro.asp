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

'on Error resume next

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
Dim reside_sw
Dim rsTaxBillCount, taxBillTotCnt
Dim rsCostEndNonSideCount, nonSideTotCnt
Dim rsCostEndMonthCount, costEndTotCnt
Dim end_msg, emp_msg
Dim rsAsStatusCount, asStatusCnt

end_month = Request("end_month")
end_yn = Request("end_yn")

cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
start_date = DateAdd("m", -1, from_date)

'org_company = "���̿��������"

reside_sw = "Y"

'���ݰ�꼭 ��� �̵�� ó�� ���� Ȯ��
objBuilder.Append "SELECT COUNT(*) FROM tax_bill "
objBuilder.Append "WHERE bill_id = '1' AND cost_reg_yn = 'N' "
objBuilder.Append "	AND (bill_date >='"&from_date&"' AND bill_date <='"&to_date&"') "

Set rsTaxBillCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

taxBillTotCnt = CInt(rsTaxBillCount(0)) 'Result.RecordCount

rsTaxBillCount.Close() : Set rsTaxBillCount = Nothing

If taxBillTotCnt > 0 Then
	reside_sw = "N"
Else
	reside_sw = "Y"
End If

'AS��Ȳ ���ε� ���� Ȯ��
'objBuilder.Append "SELECT COUNT(*) FROM as_acpt_status "
'objBuilder.Append "WHERE as_month = '"&end_month&"' "

'Set rsAsStatusCount = DBConn.Execute(objBuilder.ToString())
'objBuilder.Clear()

'asStatusCnt = CInt(rsAsStatusCount(0))

'rsAsStatusCount.Close() : Set rsAsStatusCount = Nothing

'If asStatusCnt > 0 Then
'	reside_sw = "N"
'Else
'	reside_sw = "Y"
'End If

'��� ���� �� ���� ��ȸ
objBuilder.Append "SELECT COUNT(*) "
objBuilder.Append "FROM cost_end "
objBuilder.Append "WHERE end_month = '"&end_month&"' "
objBuilder.Append "AND saupbu <> '���ֺ��' "

Set rsCostEndNonSideCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

nonSideTotCnt = CInt(rsCostEndNonSideCount(0)) 'Result.RersCountordCount
rsCostEndNonSideCount.Close() : Set rsCostEndNonSideCount = Nothing

If nonSideTotCnt > 0 Then
	objBuilder.Append "SELECT COUNT(*) "
	objBuilder.Append "FROM cost_end "
	objBuilder.Append "WHERE end_month = '"&end_month&"' "
	objBuilder.Append "	AND (end_yn = 'N' OR end_yn = 'C') "
	objBuilder.Append "	AND saupbu <> '���ֺ��' "
	objBuilder.Append "	AND saupbu <> '�����/��������' "

	Set rsCostEndMonthCount = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	costEndTotCnt = CInt(rsCostEndMonthCount(0)) 'Result.RecordCount

	rsCostEndMonthCount.Close() : Set rsCostEndMonthCount = Nothing

	If costEndTotCnt > 0 Then
		reside_sw = "N"
	Else
		reside_sw = "Y"
	End If
End If

If reside_sw = "N" Then
	'Response.Write "<script type='text/javascript'>"
	'Response.Write "	alert('��ü ��� ������ �Ǿ� ���� �ʽ��ϴ�.');"
	'Response.Write "	location.replace('/cost/cost_end_mg.asp');"
	'Response.Write "</script>"

	Response.Write "��ü ��� ������ �Ǿ� ���� �ʽ��ϴ�."
	Response.End
Else
	'Response.Write "<script type='text/javascript'>"
	'Response.Write "	alert('����ó����!!!');"
	'Response.Write "</script>"

	DBConn.BeginTrans

	' �λ縶���� �� �޿�DATA�� ��������� ����
	Dim rsEmp, rsEmpSales, rsReside, rsResideTrade
	Dim org_bonbu, org_code, trade_bonbu
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_insa.asp" -->
<%
	' �˹ٺ�� ��������� �� ������� ����
	Dim rsAlba, rsAlbaOrg, rsAlbaOutCost, rsAlbaOutCostSales, rsAlbaOutCostTrade
	Dim rsAlbaCost
	Dim cost_center, cost_company, group_name, bill_trade_name, alba_bonbu
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_mg_alba.asp" -->
<%
	' �Ϲݺ�� ��������� �� ������� ����
	Dim rsNoTax, rsNoTaxOrg, rsTax, rsTaxOrg, rsTaxNoMg, rsTaxNoMgOrg
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_mg_cost.asp" -->
<%
	' �Ϲݺ�� ��������� ����
	Dim rsNoTaxOut, rsNoTaxOutSales, rsNoTaxOutTrade
	Dim rsTaxCost, rsCompDeal, rsCompDealTrade
	Dim cost_bonbu
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_set_cost.asp" -->
<%
	' �Ϲݺ�� ��������ο� ���� ��

	' ī���� ��������� �� ������� ����
	Dim rsCard, rsCardOrg, rsCardCost, rsCardOutCost, rsCardOutCostTrade
	Dim deptName, rsCardReside
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_set_card.asp" -->
<%
	' ī���� ��������� �� ������� ���� ��

	' ���������� ������� ����
	Dim rsTran, rsTranOrg, rsTranOutCost, rsTranOutCostOrg
	Dim rsTranDeptOutCost, rsTranDeptOutCostTrade
	Dim rsTranCost, tradeDept
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_set_transit.asp" -->
<%
	' ��뱸�� Marking ����

	'ȸ�� �� ��� ����(4�� ������ ��)
	Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
	Dim rsInsure, rsPaySum, rsPayTrade, rsPayCompOutCost
	Dim insure_tot, income_tax, annual_pay, retire_pay
	Dim rsInsureCost, rsIncomeCost, rsAnnualCost, rsRetireCost
	Dim sort_seq, cost_detail
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_insure.asp" -->
<%
	'ȸ�� �� ��� ����(�˹ٺ�)
	Dim rsAlbaTot, rsAlbaTotTrade, rsAlbaCompanyCost
	Dim sum_cost
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_alba.asp" -->
<%
	' �Ϲ� ��� SUM
	Dim rsCostSum, rsCostSumTrade, rsCompanyCost
	Dim cost_id
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_sum_cost.asp" -->
<%
	' ��� SUM ����

	' ī���� ����
	Dim rsCardMg, rsCardMgTrade, rsCardCompanyCost
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_sum_card.asp" -->
<%
	' ī���� ���� ��

	' ���������� ����
	Dim rsTranMg, rsTranMgTrade, rsTranCompanyCost
	Dim rsRepair, rsRepairTrade, rsRepairCompanyCost
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_sum_transit.asp" -->
<%
	' ���������� ���� ��

	' ����κ�/ȸ�纰 ���� �ڷ� ����
	Dim rsCostCompany, rsCostProfit, rsCompanyOutCost, rsProfitCostList
%>
	<!--#include virtual="/cost/inc/inc_company_cost_end_profit.asp" -->
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
		'Response.Write "ó���� Error�� �߻��Ͽ����ϴ�."
	Else
		DBConn.CommitTrans
		end_msg = emp_msg & "����ó�� �Ǿ����ϴ�."
		'Response.Write "����ó�� �Ǿ����ϴ�."
	End If

	'Response.Write "<script type='text/javascript'>"
	'Response.Write "	alert('"&end_msg&"');"
	'Response.Write "	location.replace('/cost/cost_end_mg.asp');"
	'Response.Write "</script>"

	Response.Write end_msg
	Response.End
End If
DBConn.Close() : Set DBConn = Nothing
%>
