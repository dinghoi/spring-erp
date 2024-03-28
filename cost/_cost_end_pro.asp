<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
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
Dim from_date, end_date, to_date, start_date
Dim rs_oil
Dim deptName
Dim emp_msg, end_msg

org_company	=	Request("org_company")
deptName		=	Request("saupbu")	'��� ���θ����� ���� ���
end_month	=	Request("end_month")
end_yn		=	Request("end_yn")

cost_year 	= Mid(end_month, 1, 4)
cost_month 	= Mid(end_month, 5)

from_date = Mid(end_month, 1, 4) & "-" & Mid(end_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
start_date = DateAdd("m", -1, from_date)

Response.Write "<script type='text/javascript'>"
Response.Write "	alert('���� ó����!');"
Response.Write "</script>"

DBConn.BeginTrans

objBuilder.Append "SELECT oil_unit_id "
objBuilder.Append "FROM oil_unit "
objBuilder.Append "WHERE oil_unit_month = '"&end_month&"' "

Set rs_oil = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If rs_oil.EOF Or rs_oil.BOF Then
	Response.Write "<script type='text/javascript'>"
	Response.Write "	alert('������ �ܰ��� �ԷµǾ� ���� �ʾ� ������ �� �� �����ϴ�.');"
	Response.Write "	location.replace('/cost/cost_end_mg.asp');"
	Response.Write "</script>"
	Response.End
Else
	'������ �ܰ� �� ���
%>
	<!--#include virtual="/cost/inc/inc_cost_end_oil.asp" -->
<%
	'���� ��� ����(�����, ��Ư��, ī��)
	emp_cnt = 0
%>
	<!--#include virtual="/cost/inc/inc_cost_end_person.asp" -->
<%
	'���� �λ縶���� ���� ���� �ľ�
	If emp_cnt > 0 Then
		'4�뺸�� �� �޿� SUM ó��
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_insure.asp" -->
<%
		'��/�˹ٺ� SUM ó��
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_bonus.asp" -->
<%
		'DB SUM �Ϲ� ���
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_cost.asp" -->
<%
		'DB SUM �����
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_transit.asp" -->
<%
		'ī���� ����
%>
		<!--#include virtual="/cost/inc/inc_cost_end_sum_card.asp" -->
<%
		'cost_end ���̺��� saupbu �÷��� ���θ�� ��Ī ���[����ȣ_20210312]
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
	' ���� �λ縶���� ���� ���� �ľ� END

	If emp_cnt = 0 Then
		emp_msg = "�λ縶���� ������ ���� �ʾҽ��ϴ�."
	Else
		emp_msg = ""
	End If

	If Err.Number <> 0 Then
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
End If
rs_oil.Close() : Set rs_oil = Nothing
DBConn.Close() : Set DBConn = Nothing
%>