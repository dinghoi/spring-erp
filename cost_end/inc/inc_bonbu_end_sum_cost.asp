<%
'야특근 마감 처리
objBuilder.Append "CALL USP_ORG_END_OVERTIME_END_UP('"&from_date&"', '"&to_date&"', '');"
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'일반 경비 마감 처리
objBuilder.Append "CALL USP_ORG_END_GENERAL_SEL('"&end_month&"', '', '"&from_date&"', '"&to_date&"');"
Set rsGeneralEnd = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsGeneralEnd.EOF Then
	arrGeneralEnd = rsGeneralEnd.getRows()
End If
rsGeneralEnd.Close() : Set rsGeneralEnd = Nothing

If IsArray(arrGeneralEnd) Then
	For i = LBound(arrGeneralEnd) To UBound(arrGeneralEnd, 2)
		slip_date = arrGeneralEnd(0, i)
		slip_seq = arrGeneralEnd(1, i)

		objBuilder.Append "CALL USP_ORG_END_GENERAL_END_UP('"&slip_date&"', '"&slip_seq&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'DB SUM 처리(비용)
objBuilder.Append "CALL USP_ORG_END_GENERAL_ORG_COST_SEL('', '"&from_date&"', '"&to_date&"');"
Set rsGeneral = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsGeneral.EOF Then
	arrGeneral = rsGeneral.getRows()
End If
rsGeneral.Close() : Set rsGeneral = Nothing

If IsArray(arrGeneral) Then
	For i = LBound(arrGeneral) To UBound(arrGeneral, 2)
		emp_company = arrGeneral(0, i)
		bonbu = arrGeneral(1, i)
		saupbu = arrGeneral(2, i)
		team = arrGeneral(3, i)
		org_name = arrGeneral(4, i)
		account = arrGeneral(5, i)
		cost = arrGeneral(6 , i)

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&emp_company&"', '"&bonbu&"', "
		objBuilder.Append "'"&saupbu&"', '"&team&"', '"&org_name&"', "
		objBuilder.Append "'일반경비', '"&account&"', '"&cost&"', '0', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'DB SUM 처리(비용 외)
objBuilder.Append "CALL USP_ORG_END_GENERAL_ETC_SEL('', '"&from_date&"', '"&to_date&"');"
Set rsEctCost = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsEctCost.EOF Then
	arrEtcCost = rsEctCost.getRows()
End If
rsEctCost.Close() : Set rsEctCost = Nothing

If IsArray(arrEtcCost) Then
	For i = LBound(arrEtcCost) To UBound(arrEtcCost, 2)
		emp_company = arrEtcCost(0, i)
		bonbu = arrEtcCost(1, i)
		saupbu = arrEtcCost(2, i)
		team = arrEtcCost(3, i)
		org_name = arrEtcCost(4, i)
		cost_id = arrEtcCost(5, i)
		account = arrEtcCost(6, i)
		cost = arrEtcCost(7, i)

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'"&cost_id&"', '"&account&"', '"&cost&"', '0', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
%>