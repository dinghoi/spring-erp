<%
'=======================
'상여 SUM
'=======================
objBuilder.Append "CALL USP_ORG_END_BONUS_SEL('"&end_month&"', '"&deptName&"');"
Set rsBonus = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsBonus.EOF Then
	arrBonus = rsBonus.getRows()
End If
rsBonus.Close() : Set rsBonus = Nothing

If IsArray(arrBonus) Then
	For i = LBound(arrBonus) To UBound(arrBonus, 2)
		org_company = arrBonus(0, i)
		org_bonbu = arrBonus(1, i)
		org_saupbu = arrBonus(2, i)
		org_team = arrBonus(3, i)
		org_name = arrBonus(4, i)
		pmg_id = arrBonus(5, i)
		cost = arrBonus(6, i)

		sort_seq = 1
		cost_detail = "상여"

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'인건비', '"&cost_detail&"', '"&cost&"', '"&sort_seq&"', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'알바비
objBuilder.Append "CALL USP_ORG_END_ALBA_SEL('"&end_month&"', '');"
Set rsAlba = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsAlba.EOF Then
	arrAlba = rsAlba.getRows()
End If
rsAlba.Close() : Set rsAlba = Nothing

If IsArray(arrAlba) Then
	For i = LBound(arrAlba) To UBound(arrAlba, 2)
		company = arrAlba(0, i)
		bonbu = arrAlba(1, i)
		saupbu = arrAlba(2, i)
		team = arrAlba(3, i)
		org_name = arrAlba(4, i)
		cost = arrAlba(5, i)

		sort_seq = 8

		objBuilder.Append "CALL USP_ORG_END_COST_ID_IN_UP('"&cost_year&"', '"&org_company&"', '"&org_bonbu&"', "
		objBuilder.Append "'"&org_saupbu&"', '"&org_team&"', '"&org_name&"', "
		objBuilder.Append "'인건비', '알바비', '"&cost&"', '"&sort_seq&"', '"&cost_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
%>