<%
'===================
'유류비 단가 및 계산
'===================
objBuilder.Append "CALL USP_ORG_END_OIL_SEL('"&end_month&"', '"&from_date&"', '"&to_date&"', ''); "
Set rsTran = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If NOT rsTran.EOF Then
	arrTran = rsTran.getRows()
End If
rsTran.Close() : Set rsTran = Nothing

If isArray(arrTran) Then
	For i = LBound(arrTran) To UBound(arrTran, 2)
		mg_ce_id = arrTran(0, i)
		run_date = arrTran(1, i)
		run_seq = arrTran(2, i)
		far = arrTran(3, i)
		liter = arrTran(4, i)

		oil_unit_average = arrTran(5, i)
		oil_price = Round(Int(far) * oil_unit_average / liter)

		objBuilder.Append "CALL USP_ORG_END_OIL_UP('"&mg_ce_id&"', '"&run_date&"', '"&run_seq&"', "
		objBuilder.Append "'"&oil_unit_average&"', '"&oil_price&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If
%>