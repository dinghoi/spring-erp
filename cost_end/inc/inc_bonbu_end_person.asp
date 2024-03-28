<%
'=======================
'개인 비용 정산
'=======================
'전사 직원 정보 조회
objBuilder.Append "CALL USP_ORG_END_PERSON_ORG_SEL('"&end_month&"', '"&start_date&"', '"&from_date&"', '');"
Set rsOrgInfo = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If NOT rsOrgInfo.EOF Then
	arrOrgInfo = rsOrgInfo.getRows()
End If
rsOrgInfo.Close() : Set rsOrgInfo = Nothing

emp_cnt = 0

If isArray(arrOrgInfo) Then
	For i = LBound(arrOrgInfo) To UBound(arrOrgInfo, 2)
		org_company = arrOrgInfo(0, i)
		org_bonbu = arrOrgInfo(1, i)
		org_saupbu = arrOrgInfo(2, i)
		org_team = arrOrgInfo(3, i)
		org_name = arrOrgInfo(4, i)
		emp_reside_place = arrOrgInfo(5, i)
		emp_reside_company = arrOrgInfo(6, i)
		emp_no = arrOrgInfo(7, i)
		emp_end_date = arrOrgInfo(8, i)
		emp_name = arrOrgInfo(9, i)
		emp_job = arrOrgInfo(10, i)
		oil_unit_id = arrOrgInfo(11, i)
		liter = arrOrgInfo(12, i)

		emp_cnt = emp_cnt + 1

		'비용 별 조직 정보 업데이트
		objBuilder.Append "CALL USP_ORG_END_COST_ORG_UP('"&from_date&"', '"&to_date&"', '"&emp_no&"', "
		objBuilder.Append "'"&org_company&"', '"&org_bonbu&"', '"&org_saupbu&"', "
		objBuilder.Append "'"&org_team&"', '"&org_name&"', '"&emp_reside_place&"', '"&emp_reside_company&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'일반 경비
		objBuilder.Append "CALL USP_ORG_END_GENERAL_COST_SEL('"&emp_no&"', '"&from_date&"', '"&to_date&"');"
		Set rs_gc = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		general_cnt = 0
		general_cost = 0
		general_pre_cnt = 0
		general_pre_cost = 0

		If Not rs_gc.EOF Then
			Do Until rs_gc.EOF
				If rs_gc("pay_yn") = "N" Then
					general_cnt  = general_cnt + CInt(rs_gc("c_cnt"))
					general_cost = general_cost + CDbl(rs_gc("cost"))
				Else
					general_pre_cnt  = general_pre_cnt + CInt(rs_gc("c_cnt"))
					general_pre_cost = general_pre_cost + CDbl(rs_gc("cost"))
				End If

				rs_gc.MoveNext()
			Loop
		End If
		rs_gc.Close()

		'야특근 비용
		objBuilder.Append "CALL USP_ORG_END_OVERTIME_COST_SEL('"&emp_no&"', '"&from_date&"', '"&to_date&"');"
		Set rs_ot = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		overtime_cnt = 0
		overtime_cost = 0

		If Not rs_ot.EOF Then
			Do Until rs_ot.EOF
				overtime_cnt  = overtime_cnt + CInt(rs_ot("c_cnt"))
				overtime_cost = overtime_cost + CDbl(rs_ot("cost"))

				rs_ot.MoveNext()
			Loop
		End If
		rs_ot.Close()

		'교통비 조회
		objBuilder.Append "CALL USP_ORG_END_TRANSIT_COST_SEL('"&emp_no&"', '"&from_date&"', '"&to_date&"');"
		Set rs_tc = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		gas_km = 0
		gas_unit = 0
		gas_cost = 0
		diesel_km = 0
		diesel_unit = 0
		diesel_cost = 0
		gasol_km = 0
		gasol_unit = 0
		gasol_cost = 0
		somopum_cost = 0
		fare_cnt = 0
		fare_cost = 0
		oil_cash_cost = 0
		repair_cost = 0
		repair_pre_cost = 0
		parking_cost = 0
		toll_cost = 0

		If Not rs_tc.EOF Then
			Do Until rs_tc.EOF
				If rs_tc("car_owner") = "대중교통" Then
					fare_cnt = fare_cnt + 1
					fare_cost = fare_cost + rs_tc("fare")
				End If

				If rs_tc("car_owner") = "개인" Then
					If rs_tc("oil_kind") = "휘발유" Then
						gasol_km = gasol_km + rs_tc("far")
					End If

					If rs_tc("oil_kind") = "디젤" Then
						diesel_km = diesel_km + rs_tc("far")
					End If

					If rs_tc("oil_kind") = "가스" Then
						gas_km = gas_km + rs_tc("far")
					End If
				End If

				If rs_tc("car_owner") = "회사" Then
					oil_cash_cost = oil_cash_cost + rs_tc("oil_price")
					repair_cost = repair_cost + rs_tc("repair_cost")
				End If

				parking_cost = parking_cost + rs_tc("parking")
				toll_cost = toll_cost + rs_tc("toll")

				rs_tc.MoveNext()
			Loop
		End If
		rs_tc.Close()

		'유류비 구분
		objBuilder.Append "CALL USP_ORG_END_OIL_COST_UNIT_SEL('"&end_month&"', '"&oil_unit_id&"');"
		Set rs_ou = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If Not rs_ou.EOF Then
			Do Until rs_ou.EOF
				If rs_ou("oil_kind") = "휘발유" Then
					gasol_unit = rs_ou("oil_unit_average")
				ElseIf rs_ou("oil_kind") = "가스" Then
					gas_unit = rs_ou("oil_unit_average")
				Else
					diesel_unit = rs_ou("oil_unit_average")
				End If

				rs_ou.MoveNext()
			Loop
		End If
		rs_ou.Close()

		liter = 10
		tot_km = gas_km + diesel_km + gasol_km
		somopum_cost = tot_km * 25

		gas_cost = Round(gas_km * gas_unit / 7)
		diesel_cost = Round(diesel_km * diesel_unit / liter)
		gasol_cost = Round(gasol_km * gasol_unit / liter)
		tot_cost = gas_cost + diesel_cost + gasol_cost

		'주유 카드 사용
		objBuilder.Append "CALL USP_ORG_END_CARD_OIL_SEL('"&emp_no&"', '"&from_date&"', '"&to_date&"');"
		Set rs_cs = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		juyoo_card_cnt = 0
		juyoo_card_cost = 0
		juyoo_card_cost_vat = 0
		juyoo_card_price = 0


		If CInt(rs_cs("c_cnt")) <>  0 Then
			juyoo_card_cnt = juyoo_card_cnt + CInt(rs_cs("c_cnt"))
			juyoo_card_cost = juyoo_card_cost + CDbl(rs_cs("cost"))
			juyoo_card_cost_vat = juyoo_card_cost_vat + CDbl(rs_cs("cost_vat"))
		End If
		rs_cs.close()

		juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat

		'카드 사용
		objBuilder.Append "CALL USP_ORG_END_CARD_SEL('"&emp_no&"', '"&from_date&"', '"&to_date&"');"
		Set rs_card = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		card_cnt = 0
		card_cost = 0
		card_cost_vat = 0
		card_price = 0

		If (CInt(rs_card("c_cnt")) <>  0) Then
			card_cnt = card_cnt + CInt(rs_card("c_cnt"))
			card_cost = card_cost + CDbl(rs_card("cost"))
			card_cost_vat = card_cost_vat + CDbl(rs_card("cost_vat"))
		End If
		rs_card.Close()

		card_price = card_cost + card_cost_vat
		cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost

		'차량 정보 조회
		objBuilder.Append "CALL USP_ORG_END_CAR_SEL("&emp_no&");"
		Set rs_car = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_car.EOF Then
			car_owner = "없음"
		Else
			car_owner = rs_car("car_owner")
		End If
		rs_car.Close()

		If car_owner = "개인" Then
			return_cash = cash_tot_cost - juyoo_card_price
		Else
			return_cash = cash_tot_cost
		End If

		objBuilder.Append "CALL USP_ORG_END_PERSON_MEMO_SEL('"&emp_no&"', '"&end_month&"');"
		Set rs_person = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_person.EOF Then
			variation_memo = ""
		Else
			variation_memo = rs_person("variation_memo")
		End If
		rs_person.Close()

		'퇴사 여부 체크
		If (emp_end_date = "1900-01-01" Or IsNull(emp_end_date) Or emp_end_date >= from_date) Then
			emp_end = "근무"
		Else
			emp_end = "퇴사"
		End If

		objBuilder.Append "CALL USP_ORG_END_PERSON_COST_DEL('"&emp_no&"', '"&end_month&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		objBuilder.Append "CALL USP_ORG_END_PERSON_COST_IN("
		objBuilder.Append "'"&end_month&"', '"&emp_no&"', '"&emp_name&"', "
		objBuilder.Append "'"&emp_job&"', '"&emp_end&"', '"&car_owner&"', "
		objBuilder.Append "'"&org_company&"', '"&org_bonbu&"', '"&org_saupbu&"', "
		objBuilder.Append "'"&org_team&"', '"&org_name&"', '"&emp_reside_place&"', "
		objBuilder.Append "'"&emp_reside_company&"', "&general_cnt&", "&general_cost&", "
		objBuilder.Append "'"&general_pre_cnt&"', '"&general_pre_cost&"', '"&overtime_cnt&"', "
		objBuilder.Append "'"&overtime_cost&"', '"&gas_km&"', '"&gas_unit&"', "
		objBuilder.Append "'"&gas_cost&"', '"&diesel_km&"', '"&diesel_unit&"', "
		objBuilder.Append "'"&diesel_cost&"', '"&gasol_km&"', '"&gasol_unit&"', "
		objBuilder.Append "'"&gasol_cost&"', '"&tot_km&"', '"&tot_cost&"', "
		objBuilder.Append "'"&somopum_cost&"', '"&fare_cnt&"', '"&fare_cost&"', "
		objBuilder.Append "'"&oil_cash_cost&"', '"&repair_cost&"', '"&repair_pre_cost&"', "
		objBuilder.Append "'"&parking_cost&"', '"&toll_cost&"', '"&juyoo_card_cnt&"', "
		objBuilder.Append "'"&juyoo_card_cost&"', '"&juyoo_card_cost_vat&"', '"&card_cnt&"', "
		objBuilder.Append "'"&card_cost&"', '"&card_cost_vat&"', '"&return_cash&"', "
		objBuilder.Append "'"&variation_memo&"'); "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

Set rs_gc = Nothing
Set rs_ot = Nothing
Set rs_tc = Nothing
Set rs_ou = Nothing
Set rs_cs = Nothing
Set rs_card = Nothing
Set rs_car = Nothing
Set rs_person = Nothing
%>