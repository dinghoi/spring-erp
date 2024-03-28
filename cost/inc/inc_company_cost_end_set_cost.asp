<%
objBuilder.Append "SELECT bonbu, company "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' AND tax_bill_yn = 'N' AND cost_center = '상주직접비' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "GROUP BY bonbu, company "

Set rsNoTaxOut = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsNoTaxOut.EOF
	cost_bonbu = rsNoTaxOut("bonbu")

	objBuilder.Append "SELECT saupbu "
	objBuilder.Append "FROM sales_org "
	objBuilder.Append "WHERE saupbu = '"&cost_bonbu&"' "
	objBuilder.Append "	AND sales_year='"&cost_year&"' "

	Set rsNoTaxOutSales = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'영업사업부가 없는 경우
	If rsNoTaxOutSales.EOF Or rsNoTaxOutSales.BOF Then
		If rsNoTaxOut("company") = "" Or IsNull(rsNoTaxOut("company")) Then
			cost_bonbu = ""
		Else
			objBuilder.Append "SELECT saupbu "
			objBuilder.Append "FROM trade "
			objBuilder.Append "WHERE trade_name = '"&rsNoTaxOut("company")&"' "

			Set rsNoTaxOutTrade = DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()

			If rsNoTaxOutTrade.EOF Or rsNoTaxOutTrade.BOF Then
				cost_bonbu = "Error"
			Else
				cost_bonbu = rsNoTaxOutTrade("saupbu")
			End If
			rsNoTaxOutTrade.Close()
		End If
	End If
	rsNoTaxOutSales.Close()

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	mg_saupbu = '"&cost_bonbu&"' "
	objBuilder.Append "WHERE pl_yn = 'Y' "
	objBuilder.Append "	AND tax_bill_yn = 'N' "
	objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
	objBuilder.Append "	AND bonbu = '"&rsNoTaxOut("bonbu")&"' "
	objBuilder.Append "	AND company = '"&rsNoTaxOut("company")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsNoTaxOut.MoveNext()
Loop
rsNoTaxOut.Close() : Set rsNoTaxOut = Nothing

'상주직접비 관리사업부 없는 경우 발생으로 코드 추가
Dim rsNoTaxOut_Re

objBuilder.Append "SELECT glct.slip_date, glct.slip_seq, glct.mg_saupbu, emmt.mg_saupbu AS m_saupbu "
objBuilder.Append "FROM general_cost AS glct "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON glct.emp_no = emmt.emp_no AND emmt.emp_month = '"&end_month&"' "
objBuilder.Append "WHERE glct.pl_yn = 'Y' AND glct.tax_bill_yn = 'N' AND glct.cost_center = '상주직접비' "
objBuilder.Append "	AND (glct.slip_date >='"&from_date&"' AND glct.slip_date <='"&to_date&"') "
objBuilder.Append "	AND glct.mg_saupbu = '' AND glct.company = '공통' "

Set rsNoTaxOut_Re = Server.CreateObject("ADODB.RecordSet")
rsNoTaxOut_Re.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsNoTaxOut_Re.EOF
	cost_bonbu = rsNoTaxOut_Re("m_saupbu")

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	mg_saupbu = '"&cost_bonbu&"' "
	objBuilder.Append "WHERE slip_date = '"&rsNoTaxOut_Re("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsNoTaxOut_Re("slip_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsNoTaxOut_Re.MoveNext()
Loop
rsNoTaxOut_Re.Close() : Set rsNoTaxOut_Re = Nothing

' 비용 직접비 관리사업부 지정
'sql = "select saupbu from general_cost where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by saupbu"
'objBuilder.Append "SELECT bonbu, org_name, mg_saupbu "
objBuilder.Append "SELECT slip_date, slip_seq, bonbu, mg_saupbu "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND cost_center = '직접비' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
'o'bjBuilder.Append "GROUP BY bonbu, org_name "

Set rsTaxCost = Server.CreateObject("ADODB.RecordSet")
rsTaxCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTaxCost.EOF
	'sql = "update general_cost set mg_saupbu = '"&rs("saupbu")&"' where (pl_yn = 'Y') and (cost_center = '직접비') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (saupbu = '"&rs("saupbu")&"') "

	objBuilder.Append "UPDATE general_cost SET "

	'objBuilder.Append "	mg_saupbu = '"&rsTaxCost("bonbu")&"' "
	If rsTaxCost("mg_saupbu") <> "" Then
		objBuilder.Append "	mg_saupbu = '"&rsTaxCost("mg_saupbu")&"' "
	Else
		objBuilder.Append "	mg_saupbu = '"&rsTaxCost("bonbu")&"' "
	End If
	objBuilder.Append "WHERE pl_yn = 'Y' "
	objBuilder.Append "	AND cost_center = '직접비' "
	'objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
	'objBuilder.Append "	and bonbu = '"&rsTaxCost("bonbu")&"' "
	objBuilder.Append "	AND slip_date = '"&rsTaxCost("slip_date")&"' AND slip_seq = '"&rsTaxCost("slip_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsTaxCost.MoveNext()
Loop
rsTaxCost.Close() : Set rsTaxCost = Nothing

' 회사간거래 체크
'회사간거래 -> 기타사업부
'sql = "select customer_no, slip_date, slip_seq from general_cost where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and tax_bill_yn = 'Y'"
objBuilder.Append "SELECT customer_no, slip_date, slip_seq "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND tax_bill_yn = 'Y' "
objBuilder.Append "	AND mg_saupbu = '기타사업부' "

Set rsCompDeal = Server.CreateObject("ADODB.RecordSet")
rsCompDeal.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsCompDeal.EOF
	'sql = "select trade_id from trade where trade_no = '"&rsGeneralTaxList("customer_no")&"'"
	'objBuilder.Append "SELECT trade_id "
	'objBuilder.Append "FROM trade "
	'objBuilder.Append "WHERE trade_no = '"&rsCompDeal("customer_no")&"' "

	'Set rsCompDealTrade = DBConn.Execute(objBuilder.ToString())
	'objBuilder.Clear()

	'If rsCompDealTrade.EOF Or rsCompDealTrade.BOF Then
	'	cost_center = ""
	'Else
		'If rsCompDealTrade("trade_id") = "계열사" Then
			'sql = "update general_cost set cost_center = '회사간거래' where slip_date ='"&rsGeneralTaxList("slip_date")&"' and slip_seq = '"&rsGeneralTaxList("slip_seq")&"'"
			objBuilder.Append "UPDATE general_cost SET "
			objBuilder.Append "	cost_center = '직접비' "
			objBuilder.Append "WHERE slip_date ='"&rsCompDeal("slip_date")&"' "
			objBuilder.Append "	AND slip_seq = '"&rsCompDeal("slip_seq")&"' "

			DBConn.Execute(objBuilder.ToString())
			objBuilder.Clear()
		'End If
	'End If
	'rsCompDealTrade.Close()

	rsCompDeal.MoveNext()
Loop
rsCompDeal.Close() : Set rsCompDeal = Nothing
%>