<%
'�ŷ�ó �� ���� �ڷ� ����
Dim manage_type

Dim rsCowork, as_give_cowork, as_get_cowork, cowork_give_cost, cowork_get_cost

Dim exceptDate

'202204������ �������� SI1���� ���� �Ｚ������(��) ���� ���� ó��(�繫 ��û)[����ȣ_20220511]
exceptDate = "202204"

'�ŷ�ó ���� �ʱ�ȭ
objBuilder.Append "DELETE FROM company_cost_profit "
objBuilder.Append "WHERE cost_month = '"&end_month&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'���� ����� ��ȸ
objBuilder.Append "SELECT saupbu FROM sales_org "
objBuilder.Append "WHERE sales_year = '"&cost_year&"' "
objBuilder.Append "ORDER BY sort_seq ASC "

Set rsSalesOrg = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsSalesOrg.EOF Then
	arrSalesOrg = rsSalesOrg.getRows()
End If
rsSalesOrg.Close() : Set rsSalesOrg = Nothing

If IsArray(arrSalesOrg) Then
	For i = LBound(arrSalesOrg) To UBound(arrSalesOrg, 2)
		saupbu = arrSalesOrg(0, i)

		'����κ� ���� ��ȸ
		objBuilder.Append "SELECT SUM(cost_amt) AS 'sales_total' "
		objBuilder.Append "FROM saupbu_sales "
		objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&cost_date&"' "
		objBuilder.Append "	AND saupbu = '"&saupbu&"'; "

		Set rsSalesTot = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		sales_total = CDbl(f_toString(rsSalesTot(0), 0))	'����� �� �� ����

		rsSalesTot.Close() : Set rsSalesTot = Nothing

		'���������� Total(���� ����)
		objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS 'company_total' "
		objBuilder.Append "FROM company_cost "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' AND cost_center = '����������' "

		If saupbu <> "��Ÿ�����" Then
			objBuilder.Append "	AND (company <> '' AND company IS NOT NULL AND company <> '����') "
			objBuilder.Append " AND saupbu = '"&saupbu&"'; "
		Else
			objBuilder.Append "	AND saupbu = '' "
		End If

		Set rsCompanyTot = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		company_tot = CDbl(rsCompanyTot(0))	'����� �� ����������(���� ����)

		rsCompanyTot.Close() : Set rsCompanyTot = Nothing

		'������(������ + ����������(����))
		objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS 'comm_cost', "

		'If mm = "06" Or mm = "12" Then
		''	objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") - (SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' AND saupbu = '"&saupbu&"') FROM company_cost  "
		'Else
		''	objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") FROM company_cost  "
		'End If
		'objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '������'  "

		If saupbu = "��Ÿ�����" Then
			objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") FROM company_cost  "
			objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '������'  "
			objBuilder.Append "		AND (saupbu = '' OR saupbu = '"&saupbu&"')) AS 'direct_cost' "
		Else
			'objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") - (SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' AND saupbu = '"&saupbu&"') FROM company_cost  "
			'objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '������'  "
			If mm = "06" Or mm = "12" Then
				objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") - (SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '������' AND cost_id = '�Ϲݰ��' AND cost_detail = '�޿�' AND saupbu = '"&saupbu&"') FROM company_cost  "
			Else
				objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") FROM company_cost  "
			End If
			objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '������'  "

			objBuilder.Append "		AND saupbu = '"&saupbu&"') AS 'direct_cost' "
		End If

		objBuilder.Append "FROM company_cost "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
		objBuilder.Append "	AND (company = '' OR company is null OR company = '����') "
		objBuilder.Append "	AND cost_center = '����������' "
		objBuilder.Append "	AND saupbu = '"&saupbu&"' "

		'If saupbu = "����Ʈ����" then
		'dbconn.rollbacktrans
		'Response.write objBuilder.toString()
		'Response.end
		'end if
		Set rsComm = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		comm_cost = CDbl(f_toString(rsComm("comm_cost"), 0))	'����������(����)
		direct_cost = CDbl(f_toString(rsComm("direct_cost"), 0))	'������

		'������ = ����������(����) + ������(�ΰǺ�+���)
		common_total = comm_cost + direct_cost

		rsComm.Close() : Set rsComm = Nothing

		'��������
		objBuilder.Append "SELECT ROUND((tot_cost_amt * 0.5 / tot_person * saupbu_person) "
		objBuilder.Append "	+ (tot_cost_amt * 0.5 / tot_sale * saupbu_sale), 1) AS tot_amt "
		objBuilder.Append "FROM ( "
		objBuilder.Append "	SELECT mgct.saupbu, mgct.tot_cost_amt, mgct.saupbu_person, mgct.tot_person, "

		objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
		objBuilder.Append "		FROM saupbu_sales "
		objBuilder.Append "		WHERE SUBSTRING(sales_date, 1, 7) = '"&cost_date&"' "
		objBuilder.Append "			AND mgct.saupbu = saupbu "

		If Replace(cost_date, "-", "") >= exceptDate Then
			objBuilder.Append "		AND company <> '�Ｚ������(��)' "
		End If

		objBuilder.Append "		) AS saupbu_sale, "

		objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
		objBuilder.Append "		FROM saupbu_sales "
		objBuilder.Append "		WHERE SUBSTRING(sales_date, 1, 7) = '"&cost_date&"' "
		objBuilder.Append "			AND saupbu <> '��Ÿ�����' "

		If Replace(cost_date, "-", "") >= exceptDate Then
			objBuilder.Append "		AND company <> '�Ｚ������(��)' "
		End If

		objBuilder.Append "		) AS tot_sale "

		objBuilder.Append "	FROM management_cost AS mgct "
		objBuilder.Append "	WHERE cost_month = '"&end_month&"' "
		objBuilder.Append "		AND saupbu = '"&saupbu&"' "
		objBuilder.Append "	GROUP BY saupbu "
		objBuilder.Append ") r1 "

		Set rsManage = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsManage.EOF Or rsManage.BOF Then
			manage_tot = 0
		Else
			manage_tot = CDbl(f_toString(rsManage(0), 0))	'�μ��� ��������
		End If
		rsManage.Close() : Set rsManage = Nothing

		'if saupbu = "����Ʈ����" then
		''	dbconn.rollbacktrans
		''	response.write manage_tot
		''	response.end
		'end if

		'�ι������(���)
		objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
		objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
		objBuilder.Append "	AND cost_detail = '��ġ����')) AS 'part_tot_cost', "
		objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&cost_year&mm&"') AS 'as_tot_cnt' "
		objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '�ι������' "

		Set rsPart = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'�ι������(���)
		as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS �� �Ǽ�

		rsPart.Close() : Set rsPart = Nothing

		'����� �� AS �� �Ǽ� ��ȸ
		objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
		objBuilder.Append "FROM as_acpt_status AS aast "
		objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
		objBuilder.Append "	AND trdt.trade_id = '����' "
		objBuilder.Append "WHERE as_month = '"&cost_year&mm&"' "
		If saupbu = "��Ÿ�����" Then
			objBuilder.Append "AND trdt.saupbu = '' "
		Else
			objBuilder.Append "	AND trdt.saupbu = '"&saupbu&"' "
		End If

		Set rsSaupbuPart = DBConn.Execute(objBuilder.ToString())

		part_cnt = CInt(f_toString(rsSaupbuPart(0), 0))	'����� AS �� �Ǽ�

		objBuilder.Clear()
		rsSaupbuPart.Close() : Set rsSaupbuPart = Nothing

		'����κ� ��� �κа����
		If part_cnt > 0 Then
			part_tot = part_tot_cost / as_tot_cnt * part_cnt
		Else
			part_tot = 0
		End If

		'�ŷ�ó�� ��� ��Ȳ
		objBuilder.Append "CALL USP_SALES_COMPANY_PROFIT_SEL('"&saupbu&"', '"&cost_year&"', '"&MID(from_date, 1, 7)&"', '"&mm&"');"

		Set rsCompCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If Not rsCompCost.EOF Then
			arrCompCost = rsCompCost.getRows()
		End If
		rsCompCost.Close() : Set rsCompCost = Nothing

		'����Ʈ�� �������� �Ⱥ� ����
		' 1. �����ο� �ΰǺ�
		' 2. �����
		' 3. �ŷ�ó�� ����
		If manage_tot > 0 Then
			'���������� ���� �Ⱥ�
			If company_tot > 0 Then
				'manage_cost = manage_tot * company_cost / company_tot
				manage_type = "company"
			Else
				'���� ���� �Ⱥ�
				If sales_total > 0 Then
					'manage_cost =  manage_tot * sales_cost / sales_total	'����Ʈ�� ��������(���� ����)
					manage_type = "sales"
				Else
					'�ŷ�ó �� ��� �հ� �Ⱥ�
					'manage_cost = manage_tot * (company_cost + common_cost) / (common_total + company_tot)
					manage_type = "common"
				End If
			End If
		Else
			manage_type = "none"
		End If

		If IsArray(arrCompCost) Then
			'����Ʈ �� �б� ó��
			For j = LBound(arrCompCost) To UBound(arrCompCost, 2)
				company = arrCompCost(0, j)	'�ŷ�ó��
				sales_cost = CDbl(arrCompCost(1, j))	'�ŷ�ó�� ����
				company_cost = CDbl(arrCompCost(2, j))	'����������(�ΰǺ�+�Ϲݰ��)
				pay_cost = CDbl(arrCompCost(3, j))	'����������(�ΰǺ�)
				general_cost = CDbl(arrCompCost(4, j))	'����������(�Ϲݰ��)
				as_cnt = CInt(arrCompCost(5, j))	'����Ʈ�� AS �Ǽ�

				'����ΰ����� = �ŷ�ó�� ���������� / ����� �� ����������(���� ����) * ������
				If company_tot > 0 Or company_tot < 0 Then
					common_cost = company_cost / company_tot * common_total
				Else
					common_cost = 0
				End If

				If as_cnt > 0 Then
					part_cost = part_tot / part_cnt * as_cnt	'����Ʈ�� �κа����(AS�Ǽ� ����)
				Else
					part_cost = 0
				End If

				'����Ʈ�� �������� ���
				Select Case manage_type
					Case "company"
						'���������� ���� �Ⱥ�
						manage_cost = manage_tot * company_cost / company_tot
					Case "sales"
						'���� ���� �Ⱥ�
						manage_cost =  manage_tot * sales_cost / sales_total	'����Ʈ�� ��������(���� ����)
					Case "common"
						'�ŷ�ó �� ��� �հ� �Ⱥ�
						manage_cost = manage_tot * (company_cost + common_cost) / (common_total + company_tot)
					Case Else
						'manage_cost = 0
						manage_cost = manage_tot * company_cost / company_tot
				End Select

				'���� �Ǽ�
				objBuilder.Append "SELECT aast.as_give_cowork, aast.as_get_cowork FROM as_acpt_status AS aast "
				objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
				objBuilder.Append "WHERE aast.as_month = '"&end_month&"' AND aast.as_company = '"&company& "' "
				If saupbu = "��Ÿ�����" Then
					objBuilder.Append "	AND trdt.saupbu = '' "
				Else
					objBuilder.Append "	AND trdt.saupbu = '"&saupbu&"' "
				End If

				Set rsCowork = DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()

				If rsCowork.EOF Or rsCowork.BOF Then
					as_give_cowork = 0
					as_get_cowork = 0
				Else
					as_give_cowork = CDbl(rsCowork("as_give_cowork"))
					as_get_cowork = CDbl(rsCowork("as_get_cowork"))
				End If

				rsCowork.Close() : Set rsCowork = Nothing

				'���� ���� ���(���������� * ���������Ǽ� / ����Ʈ�� �ѰǼ�)
				'cowork_give_cost = company_cost * as_give_cowork / as_cnt * -1
				cowork_give_cost = 30000 * as_give_cowork * -1

				'���� ���� ���(���������� * ���������Ǽ� / ����Ʈ�� �ѰǼ�)
				'cowork_get_cost = company_cost * as_get_cowork / as_cnt
				cowork_get_cost = 30000 * as_get_cowork

				'pay_cost = pay_cost + cowork_give_cost + cowork_get_cost

				'���� ���
				profit_cost = sales_cost - (pay_cost + general_cost + common_cost + part_cost + manage_cost)
				'profit_cost = sales_cost - (pay_cost + general_cost + common_cost + part_cost + manage_cost + cowork_give_cost + cowork_get_cost)

				objBuilder.Append "INSERT INTO company_cost_profit(cost_month, company_name, saupbu, sales_cost, pay_cost, "
				objBuilder.Append "general_cost, common_cost, part_cost, manage_cost, profit_cost, "
				objBuilder.Append "reg_date, reg_id, cowork_give_cost, cowork_get_cost)VALUES("
				objBuilder.Append "'"&end_month&"', '"&company&"', '"&saupbu&"', '"&sales_cost&"', '"&pay_cost&"', "
				objBuilder.Append "'"&general_cost&"', '"&common_cost&"', '"&part_cost&"', '"&manage_cost&"', '"&profit_cost&"', "
				objBuilder.Append "NOW(), '"&emp_no&"', '"&cowork_give_cost&"', '"&cowork_get_cost&"');"

				DBConn.Execute(objBuilder.ToString())
				objBuilder.Clear()
			Next

		End If
	Next
End If


%>