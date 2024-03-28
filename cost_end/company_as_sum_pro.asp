<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<%
'on Error resume next

Server.ScriptTimeOut = 500

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
Dim end_month, end_yn, from_date, end_date, to_date
Dim mm, cost_year, cost_month

Dim saupbu_tab(11,2), i
Dim rs_check, check_sw
Dim end_msg, emp_msg
Dim sales_month

end_month = Request("end_month")
end_yn = Request("end_yn")

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

mm = Mid(end_month, 5, 2)
cost_year = Mid(end_month, 1, 4)
cost_month = Mid(end_month, 5)

'�ű� ��¥ ǥ�� �߰�[����ȣ]
sales_month = cost_year&"-"&cost_month

For i = 1 To 10
	saupbu_tab(i, 1) = ""
	saupbu_tab(i, 2) = 0
Next

'sql = "select * from cost_end where end_month = '"&end_month&"' and (end_yn = 'Y') and (saupbu = '���ֺ��')"
objBuilder.Append "SELECT end_month "
objBuilder.Append "FROM cost_end "
objBuilder.Append "WHERE end_month = '"&end_month&"' "
objBuilder.Append "	AND end_yn = 'Y' "
objBuilder.Append "	AND saupbu = '���ֺ��' "

Set rs_check = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rs_check.EOF Or rs_check.BOF Then
	check_sw = "N"
Else
  	check_sw = "Y"
End If

If check_sw = "N" Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('���ֺ�� ������ �����ϼž� �մϴ� !!');"
	Response.Write "	location.replace('/cost/cost_end_mg.asp');"
	Response.Write "</script>"
	Response.End
Else
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('����ó����!!!');"
	Response.Write "</script>"

	DBConn.BeginTrans

	'AS ��� ���� ����
	Dim rsCostAmtTot, rsAsCnt, rsRemoteCnt
	Dim tot_cost, won_cnt
	Dim won_per, bang_per
	Dim rsRemoteTrade, trade_bonbu, charge_per, cost_amt
	Dim rsNoRemote, rsNoRemoteCnt, rsNoRemoteTrade, rsCompAsEtc
	Dim bang_cnt
%>
	<!--#include virtual="/cost/inc/inc_company_as_acpt.asp" -->
<%
	' �����/���� �� ���� �ڷ� ����, �ι������ ���
	Dim rsProfitDept, rsProfitDeptCost, rsProfitDeptList
	Dim rsAsCompany, rsAsCompanyCost, rsAsCompTrade, rsAsCompanyList
	Dim profit_cost, company_cost, group_name
%>
	<!--#include virtual="/cost/inc/inc_company_as_field_cost.asp" -->
<%
	' ����κ� ���� �ڷ� ����
	Dim rsSalesDept, rsPayCnt, rsCompanyCostTot
	Dim rsSalesCost, rsSalesCompCost, rsCompanyCost, rsCompanyCommCost
	Dim rsMgCost, rsMgCompCost, rsMgCostTrade, rsMgProfit
	Dim rsMgDeptCost, rsMgDeptCompany, rsMgDeptTrade, rsMgDeptProfitList

	Dim tot_person, saupbu_person, tot_cost_amt
	Dim saupbu_sales, saupbu_per, saupbu_cost_amt, k
	Dim cost
%>
	<!--#include virtual="/cost/inc/inc_company_as_profit_loss.asp" -->
<%
	If end_yn = "C" Then
		'sql = "Update cost_end set end_yn='Y',reg_id='"&user_id&"',reg_name='"&user_name&"',reg_date=now() where end_month = '"&end_month& _
		'"' and saupbu = '�����/��������'"
		objBuilder.Append "UPDATE cost_end SET "
		objBuilder.Append "	end_yn = 'Y',"
		objBuilder.Append "	reg_id = '"&user_id&"',"
		objBuilder.Append "	reg_name = '"&user_name&"',"
		objBuilder.Append "	reg_date = NOW()"
		objBuilder.Append "WHERE end_month = '"&end_month&"' AND saupbu = '�����/��������' "
	Else
		'sql="insert into cost_end (end_month,saupbu,end_yn,batch_yn,bonbu_yn,ceo_yn,reg_id,reg_name,reg_date) values ('"&end_month& _
		'"','�����/��������','Y','N','N','N','"&user_id&"','"&user_name&"',now())"
		objBuilder.Append "INSERT INTO cost_end(end_month, saupbu, end_yn,  "
		objBuilder.Append "batch_yn, bonbu_yn, ceo_yn, reg_id, reg_name, reg_date)VALUES("
		objBuilder.Append "'"&end_month&"', '�����/��������', 'Y', "
		objBuilder.Append "'N', 'N', 'N', '"&user_id&"', '"&user_name&"', NOW()) "
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
	Response.Write "	location.replace('cost_end_mg.asp');"
	Response.Write "</script>"
	Response.End

	DBConn.Close() : Set DBConn = Nothing
End If
%>