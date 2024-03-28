<%
' ȸ�纰 ��� ������ ���� ������ Clear
'sql = "update company_cost set cost_amt_"&cost_month&"='0' where cost_year ='"&cost_year&"'"
objBuilder.Append "UPDATE company_cost SET "
objBuilder.Append "	cost_amt_"&cost_month&" = '0' "
objBuilder.Append "WHERE cost_year ='"&cost_year&"' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 4�뺸������ ��Ÿ �ΰǺ��� �˻�
'sql = "select * from insure_per where insure_year = '"&cost_year&"'"
objBuilder.Append "SELECT insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per "
objBuilder.Append "FROM insure_per WHERE insure_year = '"&cost_year&"' "

Set rsInsure = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

insure_tot_per = rsInsure("insure_tot_per")
income_tax_per = rsInsure("income_tax_per")
annual_pay_per = rsInsure("annual_pay_per")
retire_pay_per = rsInsure("retire_pay_per")

rsInsure.Close() : Set rsInsure = Nothing

' �޿� SUM
' 1. ������ �ΰǺ�
'sql = "select mg_saupbu,cost_center,pmg_reside_company,pmg_id,sum(pmg_give_total) as tot_cost,sum(pmg_base_pay) as base_pay,sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay,sum(pmg_tax_no) as tax_no from pay_month_give where (pmg_yymm ='"&end_month&"') and (cost_center <> '��������') group by mg_saupbu,cost_center,pmg_reside_company,pmg_id"

objBuilder.Append "SELECT mg_saupbu, cost_center, pmg_reside_company, pmg_id, "
objBuilder.Append "	SUM(pmg_give_total) as tot_cost, SUM(pmg_base_pay) as base_pay, "
objBuilder.Append "	SUM(pmg_meals_pay) AS meals_pay, SUM(pmg_overtime_pay) AS overtime_pay, "
objBuilder.Append "	SUM(pmg_tax_no) AS tax_no "
objBuilder.Append "FROM pay_month_give "
objBuilder.Append "WHERE pmg_yymm = '"&end_month&"' "
objBuilder.Append "	AND cost_center <> '��������' "
objBuilder.Append "GROUP BY mg_saupbu, cost_center, pmg_reside_company, pmg_id "

Set rsPaySum = Server.CreateObject("ADODB.RecordSet")
rsPaySum.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsPaySum.EOF
	If rsPaySum("pmg_id") = "1" Or rsPaySum("pmg_id") = "2" Then
		If rsPaySum("pmg_id") = "1" Then
			sort_seq = 0
			cost_detail = "�޿�"
		ElseIf rsPaySum("pmg_id") = "2" Then
			sort_seq = 2
			cost_detail = "��"
		ElseIf rsPaySum("pmg_id") = "4" Then
			sort_seq = 3
			cost_detail = "��������"
		Else
			sort_seq = 9
			cost_detail = "��Ÿ"
		End If

		group_name = ""
		bill_trade_name = ""

		If rsPaySum("cost_center") = "����������" Then
			'sql = "select * from trade where trade_name = '"&rs("pmg_reside_company")&"'"
			objBuilder.Append "SELECT group_name, bill_trade_name "
			objBuilder.Append "FROM trade "
			objBuilder.Append "WHERE trade_name = '"&rsPaySum("pmg_reside_company")&"' "

			Set rsPayTrade = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsPayTrade.EOF Or rsPayTrade.BOF Then
				group_name = "Error"
				bill_trade_name = "Error"
			Else
				group_name = rsPayTrade("group_name")
				bill_trade_name = rsPayTrade("bill_trade_name")
			End If
			rsPayTrade.Close()
		End If

		'sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='"&cost_detail&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
		objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
		objBuilder.Append "FROM company_cost "
		objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
		objBuilder.Append "	AND company = '"&rsPaySum("pmg_reside_company")&"' "
		objBuilder.Append "	AND cost_id ='�ΰǺ�' "
		objBuilder.Append "	AND cost_detail = '"&cost_detail&"' "
		objBuilder.Append "	AND bill_trade_name = '"&bill_trade_name&"' "
		objBuilder.Append "	AND group_name = '"&group_name&"' "
		objBuilder.Append "	AND saupbu = '"&rsPaySum("mg_saupbu")&"' "
		objBuilder.Append "	AND cost_center = '"&rsPaySum("cost_center")&"' "

		Set rsPayCompOutCost = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsPayCompOutCost.EOF Or rsPayCompOutCost.BOF Then
			'sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','�ΰǺ�','"&cost_detail&"','"&rs("mg_saupbu")&"',"&rs("tot_cost")&","&sort_seq&")"
			objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company, "
			objBuilder.Append "bill_trade_name, group_name, cost_id,"
			objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&", "
			objBuilder.Append "sort_seq)VALUES("
			objBuilder.Append "'"&cost_year&"', '"&rsPaySum("cost_center")&"', '"&rsPaySum("pmg_reside_company")&"',"
			objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '�ΰǺ�', "
			objBuilder.Append "'"&cost_detail&"', '"&rsPaySum("mg_saupbu")&"', "&rsPaySum("tot_cost")&", "
			objBuilder.Append sort_seq&")"
		Else
			'sql = "update company_cost set cost_amt_"&cost_month&"="&rs("tot_cost")&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='"&cost_detail&"' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
			objBuilder.Append "UPDATE company_cost SET "
			objBuilder.Append "	cost_amt_"&cost_month&" = "&rsPaySum("tot_cost")&", "
			objBuilder.Append "	sort_seq = "&sort_seq&" "
			objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
			objBuilder.Append "	AND company = '"&rsPaySum("pmg_reside_company")&"' "
			objBuilder.Append "	AND cost_id = '�ΰǺ�' "
			objBuilder.Append "	AND cost_detail = '"&cost_detail&"' "
			objBuilder.Append "	AND bill_trade_name = '"&bill_trade_name&"' "
			objBuilder.Append "	AND group_name = '"&group_name&"' "
			objBuilder.Append "	AND saupbu = '"&rsPaySum("mg_saupbu")&"' "
			objBuilder.Append "	AND cost_center = '"&rsPaySum("cost_center")&"' "
		End If
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
		rsPayCompOutCost.Close()

		If rsPaySum("pmg_id") = "1" Then
			' 4�뺸�� ���� ����
			insure_tot = CLng((CLng(rsPaySum("tot_cost"))) * insure_tot_per / 100)
			sort_seq = 2

			'sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='4�뺸��' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
			objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
			objBuilder.Append "FROM company_cost "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND company ='"&rsPaySum("pmg_reside_company")&"' "
			objBuilder.Append "	AND cost_id ='�ΰǺ�' "
			objBuilder.Append "	AND cost_detail ='4�뺸��' "
			objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
			objBuilder.Append "	AND group_name ='"&group_name&"' "
			objBuilder.Append "	AND saupbu ='"&rsPaySum("mg_saupbu")&"' "
			objBuilder.Append "	AND cost_center ='"&rsPaySum("cost_center")&"' "

			Set rsInsureCost = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsInsureCost.EOF Or rsInsureCost.BOF Then
				'sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','�ΰǺ�','4�뺸��','"&rs("mg_saupbu")&"',"&insure_tot&","&sort_seq&")"
				objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company, "
				objBuilder.Append "bill_trade_name, group_name, cost_id, "
				objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&", "
				objBuilder.Append "sort_seq)VALUES("
				objBuilder.Append "'"&cost_year&"','"&rsPaySum("cost_center")&"','"&rsPaySum("pmg_reside_company")&"',"
				objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '�ΰǺ�',"
				objBuilder.Append "'4�뺸��', '"&rsPaySum("mg_saupbu")&"', "&insure_tot&","
				objBuilder.Append sort_seq&")"
			Else
				'sql = "update company_cost set cost_amt_"&cost_month&"="&insure_tot&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='4�뺸��' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
				objBuilder.Append "UPDATE company_cost SET "
				objBuilder.Append "	cost_amt_"&cost_month&" = "&insure_tot&", "
				objBuilder.Append "	sort_seq = "&sort_seq&" "
				objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND company ='"&rsPaySum("pmg_reside_company")&"' "
				objBuilder.Append "	AND cost_id ='�ΰǺ�' "
				objBuilder.Append "	AND cost_detail ='4�뺸��' "
				objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
				objBuilder.Append "	AND group_name ='"&group_name&"' "
				objBuilder.Append "	AND saupbu ='"&rsPaySum("mg_saupbu")&"' "
				objBuilder.Append "	AND cost_center ='"&rsPaySum("cost_center")&"' "
			End If
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
			rsInsureCost.Close()

			' �ҵ漼 ��������
			income_tax = CLng((CLng(rsPaySum("tot_cost"))) * income_tax_per / 100)
			sort_seq = 3

			'sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='�ҵ漼��������' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
			objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
			objBuilder.Append "FROM company_cost "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND company ='"&rsPaySum("pmg_reside_company")&"' "
			objBuilder.Append "	AND cost_id ='�ΰǺ�' "
			objBuilder.Append "	AND cost_detail ='�ҵ漼��������' "
			objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
			objBuilder.Append "	AND group_name ='"&group_name&"' "
			objBuilder.Append "	AND saupbu ='"&rsPaySum("mg_saupbu")&"' "
			objBuilder.Append "	AND cost_center ='"&rsPaySum("cost_center")&"' "

			set rsIncomeCost = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsIncomeCost.EOF Or rsIncomeCost.BOF Then
				'sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','�ΰǺ�','�ҵ漼��������','"&rs("mg_saupbu")&"',"&income_tax&","&sort_seq&")"
				objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company,"
				objBuilder.Append "bill_trade_name, group_name, cost_id, "
				objBuilder.Append "cost_detail, saupbu,cost_amt_"&cost_month&", "
				objBuilder.Append "sort_seq)VALUES("
				objBuilder.Append "'"&cost_year&"', '"&rsPaySum("cost_center")&"', '"&rsPaySum("pmg_reside_company")&"', "
				objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '�ΰǺ�',"
				objBuilder.Append "'�ҵ漼��������', '"&rsPaySum("mg_saupbu")&"', "&income_tax&","
				objBuilder.Append sort_seq&")"
			Else
				'sql = "update company_cost set cost_amt_"&cost_month&"="&income_tax&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='�ҵ漼��������' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
				objBuilder.Append "UPDATE company_cost SET "
				objBuilder.Append "	cost_amt_"&cost_month&" = "&income_tax&","
				objBuilder.Append "	sort_seq = "&sort_seq&" "
				objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
				objBuilder.Append "	AND company = '"&rsPaySum("pmg_reside_company")&"' "
				objBuilder.Append "	AND cost_id = '�ΰǺ�' "
				objBuilder.Append "	AND cost_detail = '�ҵ漼��������' "
				objBuilder.Append "	AND bill_trade_name = '"&bill_trade_name&"' "
				objBuilder.Append "	AND group_name = '"&group_name&"' "
				objBuilder.Append "	AND saupbu = '"&rsPaySum("mg_saupbu")&"' "
				objBuilder.Append "	AND cost_center = '"&rsPaySum("cost_center")&"' "
			End If
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
			rsIncomeCost.Close()

			' ��������
			annual_pay = CLng((CLng(rsPaySum("base_pay")) + CLng(rsPaySum("meals_pay")) + CLng(rsPaySum("overtime_pay"))) * annual_pay_per / 100)
			sort_seq = 4

			'sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
			objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
			objBuilder.Append "FROM company_cost "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND company ='"&rsPaySum("pmg_reside_company")&"' "
			objBuilder.Append "	AND cost_id ='�ΰǺ�' "
			objBuilder.Append "	AND cost_detail ='��������' "
			objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
			objBuilder.Append "	AND group_name ='"&group_name&"' "
			objBuilder.Append "	AND saupbu ='"&rsPaySum("mg_saupbu")&"' "
			objBuilder.Append "	AND cost_center ='"&rsPaySum("cost_center")&"' "

			Set rsAnnualCost = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsAnnualCost.eof Or rsAnnualCost.bof Then
				'sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','�ΰǺ�','��������','"&rs("mg_saupbu")&"',"&annual_pay&","&sort_seq&")"
				objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company, "
				objBuilder.Append "bill_trade_name, group_name, cost_id, "
				objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&", "
				objBuilder.Append "sort_seq)VALUES("
				objBuilder.Append "'"&cost_year&"', '"&rsPaySum("cost_center")&"', '"&rsPaySum("pmg_reside_company")&"', "
				objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '�ΰǺ�', "
				objBuilder.Append "'��������', '"&rsPaySum("mg_saupbu")&"', "&annual_pay&", "
				objBuilder.Append sort_seq&")"
			Else
				'sql = "update company_cost set cost_amt_"&cost_month&"="&annual_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
				objBuilder.Append "UPDATE company_cost SET "
				objBuilder.Append "	cost_amt_"&cost_month&" = "&annual_pay&", "
				objBuilder.Append "	sort_seq = "&sort_seq&" "
				objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
				objBuilder.Append "	AND company = '"&rsPaySum("pmg_reside_company")&"' "
				objBuilder.Append "	AND cost_id = '�ΰǺ�' "
				objBuilder.Append "	AND cost_detail = '��������' "
				objBuilder.Append "	AND bill_trade_name = '"&bill_trade_name&"' "
				objBuilder.Append "	AND group_name = '"&group_name&"' "
				objBuilder.Append "	AND saupbu = '"&rsPaySum("mg_saupbu")&"' "
				objBuilder.Append "	AND cost_center = '"&rsPaySum("cost_center")&"' "
			End If
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
			rsAnnualCost.Close()

			' ��������
			retire_pay = CLng((CLng(rsPaySum("base_pay")) + CLng(rsPaySum("meals_pay")) + CLng(rsPaySum("overtime_pay"))) * retire_pay_per / 100)
			sort_seq = 5

			'sql = "select cost_amt_"&cost_month&" as cost from company_cost where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
			objBuilder.Append "SELECT cost_amt_"&cost_month&" AS cost "
			objBuilder.Append "FROM company_cost "
			objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
			objBuilder.Append "	AND company ='"&rsPaySum("pmg_reside_company")&"' "
			objBuilder.Append "	AND cost_id ='�ΰǺ�' "
			objBuilder.Append "	AND cost_detail ='��������' "
			objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
			objBuilder.Append "	AND group_name ='"&group_name&"' "
			objBuilder.Append "	AND saupbu ='"&rsPaySum("mg_saupbu")&"' "
			objBuilder.Append "	AND cost_center ='"&rsPaySum("cost_center")&"' "

			Set rsRetireCost = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsRetireCost.EOF Or rsRetireCost.BOF Then
				'sql = "insert into company_cost (cost_year,cost_center,company,bill_trade_name,group_name,cost_id,cost_detail,saupbu,cost_amt_"&cost_month&",sort_seq) values ('"&cost_year&"','"&rs("cost_center")&"','"&rs("pmg_reside_company")&"','"&bill_trade_name&"','"&group_name&"','�ΰǺ�','��������','"&rs("mg_saupbu")&"',"&retire_pay&","&sort_seq&")"
				objBuilder.Append "INSERT INTO company_cost(cost_year, cost_center, company,"
				objBuilder.Append "bill_trade_name, group_name, cost_id,"
				objBuilder.Append "cost_detail, saupbu, cost_amt_"&cost_month&","
				objBuilder.Append "sort_seq)VALUES("
				objBuilder.Append "'"&cost_year&"', '"&rsPaySum("cost_center")&"', '"&rsPaySum("pmg_reside_company")&"',"
				objBuilder.Append "'"&bill_trade_name&"', '"&group_name&"', '�ΰǺ�',"
				objBuilder.Append "'��������', '"&rsPaySum("mg_saupbu")&"', "&retire_pay&","
				objBuilder.Append sort_seq&")"
			Else
				'sql = "update company_cost set cost_amt_"&cost_month&"="&retire_pay&",sort_seq="&sort_seq&" where cost_year ='"&cost_year&"' and company ='"&rs("pmg_reside_company")&"' and cost_id ='�ΰǺ�' and cost_detail ='��������' and bill_trade_name ='"&bill_trade_name&"' and group_name ='"&group_name&"' and saupbu ='"&rs("mg_saupbu")&"' and cost_center ='"&rs("cost_center")&"'"
				objBuilder.Append "UPDATE company_cost SET "
				objBuilder.Append "	cost_amt_"&cost_month&"="&retire_pay&", "
				objBuilder.Append "	sort_seq="&sort_seq&" "
				objBuilder.Append "WHERE cost_year ='"&cost_year&"' "
				objBuilder.Append "	AND company ='"&rsPaySum("pmg_reside_company")&"' "
				objBuilder.Append "	AND cost_id ='�ΰǺ�' "
				objBuilder.Append "	AND cost_detail ='��������' "
				objBuilder.Append "	AND bill_trade_name ='"&bill_trade_name&"' "
				objBuilder.Append "	AND group_name ='"&group_name&"' "
				objBuilder.Append "	AND saupbu ='"&rsPaySum("mg_saupbu")&"' "
				objBuilder.Append "	AND cost_center ='"&rsPaySum("cost_center")&"' "
			End If
			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
			rsRetireCost.Close()

		End If
	End If

	rsPaySum.MoveNext()
Loop
rsPaySum.Close() : Set rsPaySum = Nothing
%>