<%
'===================
'유류비 단가 및 계산(개인)
'===================

Dim rsTran, rs_etc, rs_emp
Dim oil_unit_id, liter, oil_unit_average, oil_price

objBuilder.Append "SELECT trct.mg_ce_id, trct.oil_kind, trct.far, trct.run_date, trct.run_seq, "
objBuilder.Append "	eomt.org_team "
objBuilder.Append "FROM transit_cost AS trct "
objBuilder.Append "INNER JOIN emp_master_month AS emmt ON trct.mg_ce_id = emmt.emp_no "
objBuilder.Append "	AND emmt.emp_month = '"&end_month&"' "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month = '"&end_month&"' "
objBuilder.Append "WHERE (trct.run_date >='"&from_date&"' AND trct.run_date <='"&to_date&"') "
objBuilder.Append "	AND trct.car_owner ='개인' AND trct.far > 0 AND eomt.org_bonbu = '' "

Set rsTran = Server.CreateObject("ADODB.RecordSet")
rsTran.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTran.EOF
	'유뷰비 단가 지정
	oil_unit_id = "2"
	liter = 10

	If rsTran("oil_kind") = "가스" Then
		liter = 7
	End If

	objBuilder.Append "SELECT oil_unit_average FROM oil_unit "
	objBuilder.Append "WHERE oil_unit_month = '"&end_month&"' "
	objBuilder.Append "	AND oil_unit_id = '"&oil_unit_id&"' "
	objBuilder.Append "	AND oil_kind = '"&rsTran("oil_kind")&"' "

	Set rs_etc = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	oil_unit_average = rs_etc("oil_unit_average")

	rs_etc.Close()

	oil_price = Round(Int(rsTran("far")) * oil_unit_average / liter)

	objBuilder.Append "UPDATE transit_cost SET "
	objBuilder.Append "	oil_unit = "&oil_unit_average&", oil_price = "&oil_price&" "
	objBuilder.Append "WHERE mg_ce_id = '"&rsTran("mg_ce_id")&"' "
	objBuilder.Append "	AND run_date = '"&rsTran("run_date")&"' AND run_seq = '"&rsTran("run_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsTran.MoveNext()
Loop
Set rs_emp = Nothing
Set rs_etc = Nothing
rsTran.Close() : Set rsTran = Nothing
%>