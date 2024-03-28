<%
'4�뺸������ ��Ÿ �ΰǺ��� �˻�
objBuilder.Append "CALL USP_ORG_END_INSURE_SEL('"&cost_year&"');"
Set rs_insure = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

insure_tot_per = rs_insure("insure_tot_per")
income_tax_per = rs_insure("income_tax_per")
annual_pay_per = rs_insure("annual_pay_per")
retire_pay_per = rs_insure("retire_pay_per")

rs_insure.Close() : Set rs_insure = Nothing

'���� ��� ���� �ʱ�ȭ
objBuilder.Append "CALL USP_ORG_END_COST_RESET_UP('"&cost_year&"', '"&cost_month&"', '');"
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'�޿� ��ȸ �� ����
objBuilder.Append "CALL USP_ORG_END_PAY_SEL('"&end_month&"', '');"
Set rsPay = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsPay.EOF Then
	arrPay = rsPay.getRows()
End If
rsPay.Close() : Set rsPay = Nothing

If IsArray(arrPay) Then
	For i = LBound(arrPay) To UBound(arrPay, 2)
		org_company = arrPay(0, i)
		org_bonbu = arrPay(1, i)
		org_saupbu = arrPay(2, i)
		org_team = arrPay(3, i)
		org_name = arrPay(4, i)
		pmg_id = arrPay(5, i)
		tot_cost = arrPay(6, i)
		base_pay = arrPay(7, i)
		meals_pay = arrPay(8, i)
		overtime_pay = arrPay(9, i)
		research_pay = arrPay(10, i)
		tax_no = arrPay(11, i)

		sort_seq = 0
		cost_detail = "�޿�"

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'�ΰǺ�', '"&cost_detail&"', '"&tot_cost&"', '"&sort_seq&"', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'4�뺸���
		insure_tot = CLng((CLng(tot_cost)) * insure_tot_per / 100)
		sort_seq = 2

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'�ΰǺ�', '4�뺸��', '"&insure_tot&"', '"&sort_seq&"', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		' �ҵ漼 ��������
		income_tax = clng((clng(tot_cost)) * income_tax_per / 100)
		sort_seq = 3

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'�ΰǺ�', '�ҵ漼��������', '"&income_tax&"', '"&sort_seq&"', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'��������
		annual_pay = CLng((CLng(base_pay) + CLng(meals_pay) + CLng(overtime_pay)) * annual_pay_per / 100)
		sort_seq = 4

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'�ΰǺ�', '��������', '"&annual_pay&"', '"&sort_seq&"', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		' ��������
		retire_pay = CLng((CLng(base_pay) + CLng(meals_pay) + CLng(overtime_pay)) * retire_pay_per / 100)
		sort_seq = 5

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'�ΰǺ�', '��������', '"&retire_pay&"', '"&sort_seq&"', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
%>