<%
' 회사 차량 운행 주유카드 셋팅
objBuilder.Append "CALL USP_ORG_END_TRAN_COMP_SEL('', '"&from_date&"', '"&to_date&"');"
Set rsCardTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCardTran.EOF Then
	arrCardTran = rsCardTran.getRows()
End If
rsCardTran.Close() : Set rsCardTran = Nothing

If IsArray(arrCardTran) Then
	For i = LBound(arrCardTran) To UBound(arrCardTran, 2)
		mg_ce_id = arrCardTran(0, i)

		objBuilder.Append "CALL USP_ORG_END_TRAN_CARD_UP('"&mg_ce_id&"', '"&from_date&"', '"&to_date&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

' 카드비용 집계
objBuilder.Append "CALL USP_ORG_END_CARD_COST_SEL('', '"&from_date&"', '"&to_date&"');"
Set rsCardSlip = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsCardSlip.EOF Then
	arrCardSlip = rsCardSlip.getRows()
End If
rsCardSlip.Close() : Set rsCardSlip = Nothing

If IsArray(arrCardSlip) Then
	For i = LBound(arrCardSlip) To UBound(arrCardSlip, 2)
		emp_company = arrCardSlip(0, i)
		bonbu = arrCardSlip(1, i)
		saupbu = arrCardSlip(2, i)
		team = arrCardSlip(3, i)
		org_name = arrCardSlip(4, i)
		account = arrCardSlip(5, i)
		cost = arrCardSlip(6, i)

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&emp_company&"', '"&bonbu&"', "
		objBuilder.Append "'"&saupbu&"', '"&team&"', '"&org_name&"', "
		objBuilder.Append "'법인카드', '"&account&"', '"&cost&"', '0', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
%>