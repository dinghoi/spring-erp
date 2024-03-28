<%
'========================
'??¡¤?¨¬? ¢¥?¡Æ¢® ©ö¡¿ ¡Æ???(¡Æ©ø??)
'========================

Dim rsTran, rs_etc, rs_emp
Dim oil_unit_id, liter, oil_unit_average, oil_price

objBuilder.Append "SELECT trct.mg_ce_id, trct.oil_kind, trct.far, trct.run_date, trct.run_seq, "
objBuilder.Append "	eomt.org_team "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON trct.mg_ce_id = emmt.emp_no "
objBuilder.Append "	AND emp_month = '"&end_month&"' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month = '"&end_month&"' "
objBuilder.Append "	AND (ISNULL(eomt.org_end_date) OR eomt.org_end_date = '0000-00-00') "
objBuilder.Append "WHERE (trct.run_date >='"&from_date&"' AND trct.run_date <='"&to_date&"') "
objBuilder.Append "	AND trct.car_owner ='¡Æ©ø??' AND trct.far > 0 AND eomt.org_bonbu = '"&deptName&"' "

Set rsTran = Server.CreateObject("ADODB.RecordSet")
rsTran.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTran.EOF
	'??¡¤?¨¬? ¢¥?¡Æ¢® ???¢´
	'If (rsTran("org_team") = "¨¬???¨¡?" Or rsTran("org_team") = "SM1¨¡?" Or rsTran("org_team") = "Repair¨¡?" Or rsTran("org_team") = "SM2¨¡?") Then
	'	oil_unit_id = "1"
	'Else
	'	oil_unit_id = "2"
	'End If

	Select Case rsTran("org_team")
		Case "¨¬???¨¡?", "Repair¨¡?"
			oil_unit_id = "1"
		Case Else
			oil_unit_id = "2"
	End Select

	objBuilder.Append "SELECT emp_reside_company "
	objBuilder.Append "FROM emp_master_month "
	objBuilder.Append "WHERE emp_month = '"&end_month&"' "
	objBuilder.Append "	AND emp_no = '"&rsTran("mg_ce_id")&"' "
	
	Set rs_emp = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not rs_emp.EOF Then
		If rs_emp("emp_reside_company") = "???¡©?¡©¨ú?" Then
			liter = 8
		Else
			liter = 10
		End If
	Else
		liter = 10
	End If
	rs_emp.Close()

	If rsTran("oil_kind") = "¡Æ¢®¨ö¨¬" Then
		liter = 7
	End If

	objBuilder.Append "SELECT oil_unit_average "
	objBuilder.Append "FROM oil_unit "
	objBuilder.Append "WHERE oil_unit_month = '"&end_month&"' "
	objBuilder.Append "	AND oil_unit_id = '"&oil_unit_id&"' "
	objBuilder.Append "	AND oil_kind = '"&rsTran("oil_kind")&"' "

	Set rs_etc = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	oil_unit_average = rs_etc("oil_unit_average")

	rs_etc.Close()

	oil_price = Round(Int(rsTran("far")) * oil_unit_average / liter)

	objBuilder.Append "UPDATE transit_cost SET "
	objBuilder.Append "	oil_unit = "&oil_unit_average&", "
	objBuilder.Append "	oil_price = "&oil_price&" "
	objBuilder.Append "WHERE mg_ce_id = '"&rsTran("mg_ce_id")&"' "
	objBuilder.Append "	AND run_date = '"&rsTran("run_date")&"' "
	objBuilder.Append "	AND run_seq = '"&rsTran("run_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsTran.MoveNext()
Loop

Set rs_emp = Nothing
Set rs_etc = Nothing
rsTran.Close() : Set rsTran = Nothing
%>