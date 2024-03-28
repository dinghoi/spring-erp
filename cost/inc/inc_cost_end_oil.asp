<%
'========================
'유류비 단가 및 계산(개인)
'========================
objBuilder.Append "SELECT trct.mg_ce_id, trct.oil_kind, trct.far, trct.run_date, trct.run_seq, "
objBuilder.Append "	eomt.org_team "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON trct.mg_ce_id = emmt.emp_no "
objBuilder.Append "	AND emp_month = '"&end_month&"' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month = '"&end_month&"' "
objBuilder.Append "	AND (ISNULL(eomt.org_end_date) OR eomt.org_end_date = '0000-00-00') "
objBuilder.Append "WHERE (trct.run_date >='"&from_date&"' AND trct.run_date <='"&to_date&"') "
objBuilder.Append "	AND trct.car_owner ='개인' AND trct.far > 0 AND eomt.org_bonbu = '"&deptName&"' "

Set rsTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsTran.EOF Then
	arrTran = rsTran.getRows()
End If
rsTran.Close() : Set rsTran = Nothing

If IsArray(arrTran) Then
	For i = LBound(arrTran) To UBound(arrTran, 2)
		mg_ce_id = arrTran(0, i)
		oil_kind = arrTran(1, i)
		far = arrTran(2, i)
		run_date = arrTran(3, i)
		run_seq = arrTran(4, i)
		org_team = arrTran(5, i)

		'유류비 단가 지정
		'If (rsTran("org_team") = "본사팀" Or rsTran("org_team") = "SM1팀" Or rsTran("org_team") = "Repair팀" Or rsTran("org_team") = "SM2팀") Then
		'	oil_unit_id = "1"
		'Else
		'	oil_unit_id = "2"
		'End If

		'Select Case rsTran("org_team")
		Select Case org_team
			Case "본사팀", "Repair팀"
				oil_unit_id = "1"
			Case Else
				oil_unit_id = "2"
		End Select

		objBuilder.Append "SELECT emp_reside_company "
		objBuilder.Append "FROM emp_master_month "
		objBuilder.Append "WHERE emp_month = '"&end_month&"' "
		objBuilder.Append "	AND emp_no = '"&mg_ce_id&"' "

		Set rs_emp = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If Not rs_emp.EOF Then
			If rs_emp("emp_reside_company") = "한화화약" Then
				liter = 8
			Else
				liter = 10
			End If
		Else
			liter = 10
		End If
		rs_emp.Close()

		If oil_kind = "가스" Then
			liter = 7
		End If

		objBuilder.Append "SELECT oil_unit_average "
		objBuilder.Append "FROM oil_unit "
		objBuilder.Append "WHERE oil_unit_month = '"&end_month&"' "
		objBuilder.Append "	AND oil_unit_id = '"&oil_unit_id&"' "
		objBuilder.Append "	AND oil_kind = '"&oil_kind&"' "

		Set rs_etc = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		oil_unit_average = rs_etc("oil_unit_average")
		rs_etc.Close()

		oil_price = Round(Int(far) * oil_unit_average / liter)

		objBuilder.Append "UPDATE transit_cost SET "
		objBuilder.Append "	oil_unit = "&oil_unit_average&", "
		objBuilder.Append "	oil_price = "&oil_price&" "
		objBuilder.Append "WHERE mg_ce_id = '"&mg_ce_id&"' "
		objBuilder.Append "	AND run_date = '"&run_date&"' "
		objBuilder.Append "	AND run_seq = '"&run_seq&"' "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
Set rs_emp = Nothing
Set rs_etc = Nothing
%>