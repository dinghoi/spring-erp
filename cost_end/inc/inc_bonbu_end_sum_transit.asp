<%
'교통비 마감
objBuilder.Append "CALL USP_ORG_END_TRANSIT_END_UP('"&from_date&"', '"&to_date&"', '');"
DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'교통비 정보 조회
objBuilder.Append "CALL USP_ORG_END_TRANSIT_SEL('', '"&from_date&"', '"&to_date&"');"
Set rsTransit = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsTransit.EOF Then
	arrTransit = rsTransit.getRows()
End If
rsTransit.Close() : Set rsTransit = Nothing

If IsArray(arrTransit) Then
	For i = LBound(arrTransit) To UBound(arrTransit, 2)
		emp_company = arrTransit(0, i)
		bonbu = arrTransit(1, i)
		saupbu = arrTransit(2, i)
		team = arrTransit(3, i)
		org_name = arrTransit(4, i)
		car_owner = arrTransit(5, i)
		cost = arrTransit(6, i)

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&emp_company&"', '"&bonbu&"', "
		objBuilder.Append "'"&saupbu&"', '"&team&"', '"&org_name&"', "
		objBuilder.Append "'교통비', '"&car_owner&"', '"&cost&"', '0', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'DB SUM 교통비(차량수리비)
objBuilder.Append "CALL USP_ORG_END_TRAN_REPAIR_SEL('', '"&from_date&"', '"&to_date&"');"
Set rsRepair = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsRepair.EOF Then
	arrRepair = rsRepair.getRows()
End If
rsRepair.Close() : Set rsRepair = Nothing

If IsArray(arrRepair) Then
	For i = LBound(arrRepair) To UBound(arrRepair, 2)
		emp_company = arrRepair(0, i)
		bonbu = arrRepair(1, i)
		saupbu = arrRepair(2, i)
		team = arrRepair(3, i)
		org_name = arrRepair(4, i)
		cost = arrRepair(5, i)

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&emp_company&"', '"&bonbu&"', "
		objBuilder.Append "'"&saupbu&"', '"&team&"', '"&org_name&"', "
		objBuilder.Append "'교통비', '차량수리비', '"&cost&"', '0', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
%>