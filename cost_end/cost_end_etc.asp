<%
' ������ �ܰ� �� ������ ���
%>
<!--#include virtual="/cost/inc/inc_bonbu_end_oil.asp" -->
<%
' ���κ� ��� ����
%>
<!--#include virtual="/cost/inc/inc_bonbu_end_person.asp" -->
<%
' ���� �λ縶���� ���� ���� �ľ�
If emp_cnt > 0 Then
	'4�뺸�� �� �޿� SUM ó��
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_insure.asp" -->
<%
	'��/�˹ٺ� SUM ó��
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_bonus.asp" -->
<%
	'DB SUM �Ϲ� ���
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_cost.asp" -->
<%
	'DB SUM �����
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_transit.asp" -->
<%
	'ī���� ����
%>
	<!--#include virtual="/cost/inc/inc_bonbu_end_sum_card.asp" -->
<%
	objBuilder.Append "CALL USP_ORG_END_PROC('"&end_month&"', '����οܳ�����', '"&end_yn&"', '"&user_id&"', '"&user_name&"');"
	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()
End If
' ���� �λ縶���� ���� ���� �ľ� END