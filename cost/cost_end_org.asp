<%
'��ü ���� �� ��� ���� ó�� �߰�(�ϰ� ��������� �߰�)[����ȣ_20211007]
objBuilder.Append "CALL USP_ORG_END_SALES_SEL();"
Set rsSalesOrg = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsSalesOrg.EOF Then
	arrSalesOrg = rsSalesOrg.getRows()
End If
rsSalesOrg.Close() : Set rsSalesOrg = Nothing

If IsArray(arrSalesOrg) Then
	For oLoop = LBound(arrSalesOrg) To UBound(arrSalesOrg, 2)
		deptName = arrSalesOrg(0, oLoop)	'���θ�
		emp_cnt = 0

		'������ �ܰ� �� ���
%>
		<!--#include virtual="/cost/inc/inc_org_end_oil.asp" -->
<%
		'���� ��� ����(�����, ��Ư��, ī��)
%>
		<!--#include virtual="/cost/inc/inc_org_end_person.asp" -->
<%
		'���� �λ縶���� ���� ���� �ľ�
		If emp_cnt > 0 Then
			'4�뺸�� �� �޿� SUM ó��
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_insure.asp" -->
<%
			'��/�˹ٺ� SUM ó��
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_bonus.asp" -->
<%
			'DB SUM �Ϲ� ���
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_general.asp" -->
<%
			'DB SUM �����
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_transit.asp" -->
<%
			'ī���� ����
%>
			<!--#include virtual="/cost/inc/inc_org_end_sum_card.asp" -->
<%
			'cost_end ���̺��� saupbu �÷��� ���θ�� ��Ī ���[����ȣ_20210312]
			objBuilder.Append "CALL USP_ORG_END_PROC('"&end_month&"', '"&deptName&"', '"&end_yn&"', '"&user_id&"', '"&user_name&"');"
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		End If
		' ���� �λ縶���� ���� ���� �ľ� END
	Next
End If
%>