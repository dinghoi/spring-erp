<%
'=======================
'개인 경비 정산
'=======================
'전사 직원 정보 조회
objBuilder.Append "SELECT eomt.org_company, eomt.org_bonbu, eomt.org_saupbu, eomt.org_team, "
objBuilder.Append "	eomt.org_name, emmt.emp_reside_place, emmt.emp_reside_company, "
objBuilder.Append "	emmt.emp_no, emmt.emp_end_date, emmt.emp_name, emmt.emp_job "
objBuilder.Append "FROM emp_master_month AS emmt "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month = '"&end_month&"'"
objBuilder.Append "WHERE emmt.emp_month = '"&end_month&"'"
objBuilder.Append "	AND eomt.org_bonbu =  '"&deptName&"' "
objBuilder.Append "	AND (emmt.emp_end_date = '0000-00-00' OR ISNULL(emmt.emp_end_date) "
objBuilder.Append "		OR emmt.emp_end_date >= '"&start_date&"') "

Set rsOrgInfo = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsOrgInfo.EOF Then
	arrOrgInfo = rsOrgInfo.getRows()
End If
rsOrgInfo.Close() : Set rsOrgInfo = Nothing

If IsArray(arrOrgInfo) Then
	emp_cnt = 0

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

		emp_cnt = emp_cnt + 1

		'교통비 조직 정보 업데이트
		objBuilder.Append "UPDATE transit_cost SET "
		objBuilder.Append "	emp_company ='"&org_company&"', "
		objBuilder.Append "	bonbu ='"&org_bonbu&"', "
		objBuilder.Append "	saupbu ='"&org_saupbu&"', "
		objBuilder.Append "	team ='"&org_team&"', "
		objBuilder.Append "	org_name ='"&org_name&"', "
		objBuilder.Append "	reside_place ='"&emp_reside_place&"' "
		objBuilder.Append "WHERE (run_date >='"&from_date&"' AND run_date <='"&to_date&"') "
		objBuilder.Append "	AND mg_ce_id='"&emp_no&"' "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'야특근 조직 정보 업데이트
		objBuilder.Append "UPDATE overtime SET "
		objBuilder.Append "	emp_company = '"&org_company&"', "
		objBuilder.Append "	bonbu = '"&org_bonbu&"', "
		objBuilder.Append "	saupbu ='"&org_saupbu&"', "
		objBuilder.Append "	team = '"&org_team&"', "
		objBuilder.Append "	org_name = '"&org_name&"',"
		objBuilder.Append "	reside_place = '"&emp_reside_place&"' "
		objBuilder.Append "WHERE (work_date >='"&from_date&"' AND work_date <='"&to_date&"') "
		objBuilder.Append "	AND mg_ce_id = '"&emp_no&"' "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'카드전표 조직 정보 업데이트
		objBuilder.Append "UPDATE card_slip SET "
		objBuilder.Append "	emp_company = '"&org_company&"', "
		objBuilder.Append "	bonbu = '"&org_bonbu&"', "
		objBuilder.Append "	saupbu ='"&org_saupbu&"', "
		objBuilder.Append "	team = '"&org_team&"', "
		objBuilder.Append "	org_name = '"&org_name&"', "
		objBuilder.Append "	reside_place = '"&emp_reside_place&"', "
		objBuilder.Append "	reside_company = '"&emp_reside_company&"' "
		objBuilder.Append "WHERE (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
		objBuilder.Append "	AND emp_no = '"&emp_no&"' "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		'퇴사 여부 체크
		If (emp_end_date = "1900-01-01" Or IsNull(emp_end_date) Or emp_end_date >= from_date) Then
			emp_end = "근무"
		Else
			emp_end = "퇴사"
		End If

		'일반 경비
		general_cnt = 0
		general_cost = 0
		general_pre_cnt = 0
		general_pre_cost = 0

		objBuilder.Append "SELECT pay_yn, COUNT(slip_seq) AS c_cnt, SUM(cost) AS cost "
		objBuilder.Append "FROM general_cost "
		objBuilder.Append "WHERE emp_no='"&emp_no&"' "
		objBuilder.Append "	AND slip_gubun = '비용' "
		objBuilder.Append "	AND (tax_bill_yn = 'N' OR ISNULL(tax_bill_yn)) "
		objBuilder.Append "	AND cancel_yn = 'N' "
		objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
		objBuilder.Append "GROUP BY pay_yn "

		Set rs_gc = DBConn.Execute(objBuilder.ToString())
		'rs_gc.Open objBuilder.ToString(), DBConn, 1
		objBuilder.Clear()

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
		rs_gc.Close() : Set rs_gc = Nothing

		'야특근 비용
		overtime_cnt = 0
		overtime_cost = 0

		objBuilder.Append "SELECT cancel_yn, COUNT(work_date) AS c_cnt, SUM(overtime_amt) AS cost "
		objBuilder.Append "FROM overtime "
		objBuilder.Append "WHERE mg_ce_id = '"&emp_no&"' "
		objBuilder.Append "	AND (work_date >='"&from_date&"' AND work_date <='"&to_date&"') "
		objBuilder.Append "	AND cancel_yn = 'N' "
		objBuilder.Append "GROUP  BY cancel_yn "

		Set rs_ot = DBConn.Execute(objBuilder.ToString())
		'rs_ot.Open objBuilder.ToString(), DBConn, 1
		objBuilder.Clear()

		Do Until rs_ot.EOF
			overtime_cnt  = overtime_cnt + CInt(rs_ot("c_cnt"))
			overtime_cost = overtime_cost + CDbl(rs_ot("cost"))

			rs_ot.MoveNext()
		Loop
		rs_ot.Close() : Set rs_ot = Nothing

		'교통비
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

		objBuilder.Append "SELECT car_owner, fare, far, oil_kind, oil_price, repair_cost, parking, toll "
		objBuilder.Append "FROM transit_cost "
		objBuilder.Append "WHERE mg_ce_id='"&emp_no&"' "
		objBuilder.Append "	AND (run_date >='"&from_date&"' AND run_date <='"&to_date&"') "
		objBuilder.Append "	AND cancel_yn = 'N' "

		'rs_tc.Open objBuilder.ToString(), DBConn, 1
		Set rs_tc = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

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
		rs_tc.Close() : Set rs_tc = Nothing

		'본사팀 구분
		'If rsOrgInfo("org_team") = "본사팀" Or rsOrgInfo("org_team") = "Repair팀" Or rsOrgInfo("org_team") = "SM1팀" Or rsOrgInfo("org_team") = "SM2팀" Then
		'	oil_unit_id = "1"
		'Else
		'	oil_unit_id = "2"
		'End If
		Select Case org_team
			Case "본사팀", "Repair팀"
				oil_unit_id = "1"
			Case Else
				oil_unit_id = "2"
		End Select

		objBuilder.Append "SELECT oil_unit_average, oil_kind "
		objBuilder.Append "FROM oil_unit "
		objBuilder.Append "WHERE oil_unit_month = '"&end_month&"' "
		objBuilder.Append "	AND oil_unit_id = '"&oil_unit_id&"' "

		'rs_ou.Open objBuilder.ToString(), DBConn, 1
		Set rs_ou = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

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
		rs_ou.Close() : Set rs_ou = Nothing

		If emp_reside_company = "한화화약" Then
			liter = 8
		Else
			liter = 10
		End If

		tot_km = gas_km + diesel_km + gasol_km
		somopum_cost = tot_km * 25

		gas_cost = Round(gas_km * gas_unit / 7)
		diesel_cost = Round(diesel_km * diesel_unit / liter)
		gasol_cost = Round(gasol_km * gasol_unit / liter)
		tot_cost = gas_cost + diesel_cost + gasol_cost

		'주유 카드 사용
		juyoo_card_cnt = 0
		juyoo_card_cost = 0
		juyoo_card_cost_vat = 0
		juyoo_card_price = 0

		objBuilder.Append "SELECT COUNT(*) AS c_cnt, SUM(cost) AS cost, SUM(cost_vat) AS cost_vat "
		objBuilder.Append "FROM card_slip "
		objBuilder.Append "WHERE emp_no ='"&emp_no&"' "
		objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
		objBuilder.Append "	AND card_type LIKE '%주유%' "

		Set rs_cs = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If CInt(rs_cs("c_cnt")) <>  0 Then
			juyoo_card_cnt = juyoo_card_cnt + CInt(rs_cs("c_cnt"))
			juyoo_card_cost = juyoo_card_cost + CDbl(rs_cs("cost"))
			juyoo_card_cost_vat = juyoo_card_cost_vat + CDbl(rs_cs("cost_vat"))
		End If

		rs_cs.close() : Set rs_cs = Nothing

		juyoo_card_price = juyoo_card_cost + juyoo_card_cost_vat

		'카드 사용
		card_cnt = 0
		card_cost = 0
		card_cost_vat = 0
		card_price = 0

		objBuilder.Append "SELECT COUNT(*) AS c_cnt , SUM(cost) AS cost , SUM(cost_vat) AS cost_vat "
		objBuilder.Append "FROM card_slip "
		objBuilder.Append "WHERE emp_no ='"&emp_no&"' "
		objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
		objBuilder.Append "	AND card_type NOT LIKE '%주유%' "

		Set rs_card = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If (CInt(rs_card("c_cnt")) <>  0) Then
			card_cnt = card_cnt + CInt(rs_card("c_cnt"))
			card_cost = card_cost + CDbl(rs_card("cost"))
			card_cost_vat = card_cost_vat + CDbl(rs_card("cost_vat"))
		End If

		rs_card.Close() : Set rs_card = Nothing

		card_price = card_cost + card_cost_vat
		cash_tot_cost = general_cost + gas_cost + diesel_cost + gasol_cost + somopum_cost + fare_cost + oil_cash_cost + toll_cost + parking_cost

		'차량 정보 조회
		objBuilder.Append "SELECT car_owner "
		objBuilder.Append "FROM car_info "
		objBuilder.Append "WHERE owner_emp_no ='"&emp_no&"' "

		Set rs_car = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_car.EOF Then
			car_owner = "없음"
		Else
			car_owner = rs_car("car_owner")
		End If

		rs_car.Close() : Set rs_car = Nothing

		If car_owner = "개인" Then
			return_cash = cash_tot_cost - juyoo_card_price
		Else
			return_cash = cash_tot_cost
		End If

		objBuilder.Append "SELECT variation_memo "
		objBuilder.Append "FROM person_cost "
		objBuilder.Append "WHERE cost_month = '"&end_month&"' "
		objBuilder.Append "	AND emp_no ='"&emp_no&"' "

		Set rs_person = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rs_person.EOF Then
			variation_memo = ""
		Else
			variation_memo = rs_person("variation_memo")
		End If

		rs_person.Close() : Set rs_person = Nothing

		'초기화
		objBuilder.Append "DELETE FROM person_cost "
		objBuilder.Append "WHERE cost_month ='"&end_month&"' "
		objBuilder.Append "	AND emp_no ='"&emp_no&"' "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		objBuilder.Append "INSERT INTO person_cost VALUES("
		objBuilder.Append "'"&end_month&"', '"&emp_no&"', '"&emp_name&"', "
		objBuilder.Append "'"&emp_job&"', '"&emp_end&"', '"&car_owner&"', "
		objBuilder.Append "'"&org_company&"', '"&org_bonbu&"', '"&org_saupbu&"', "
		objBuilder.Append "'"&org_team&"', '"&org_name&"', '"&emp_reside_place&"', "
		objBuilder.Append "'"&emp_reside_company&"', "&general_cnt&", "&general_cost&", "
		objBuilder.Append general_pre_cnt&", "&general_pre_cost&", "&overtime_cnt&", "
		objBuilder.Append overtime_cost&", "&gas_km&", "&gas_unit&", "
		objBuilder.Append gas_cost&", "&diesel_km&", "&diesel_unit&", "
		objBuilder.Append diesel_cost&", "&gasol_km&", "&gasol_unit&", "
		objBuilder.Append gasol_cost&", "&tot_km&", "&tot_cost&", "
		objBuilder.Append somopum_cost&", "&fare_cnt&", "&fare_cost&", "
		objBuilder.Append oil_cash_cost&", "&repair_cost&", "&repair_pre_cost&", "
		objBuilder.Append parking_cost&", "&toll_cost&", "&juyoo_card_cnt&", "
		objBuilder.Append juyoo_card_cost&", "&juyoo_card_cost_vat&", "&card_cnt&", "
		objBuilder.Append card_cost&", "&card_cost_vat&", "&return_cash&", "
		objBuilder.Append "'"&variation_memo&"', NOW(), 0) "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
%>