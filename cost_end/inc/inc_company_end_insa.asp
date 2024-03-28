<%
'===============================
'인사마스터 및 급여DATA에 관리사업부 지정
'===============================
'인사 정보 조회(손익제외)
objBuilder.Append "CALL USP_COMPANY_END_ORG_SEL('"&end_month&"');"
Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If Not rsEmp.EOF Then
	arrEmp = rsEmp.getRows()
End If
rsEmp.Close() : Set rsEmp = Nothing

If IsArray(arrEmp) Then
	For i = LBound(arrEmp) To UBound(arrEmp, 2)
		emp_no = arrEmp(0, i)
		org_bonbu = arrEmp(1, i)
		org_code = arrEmp(2, i)

		'직원 별 관리사업부 지정
		objBuilder.Append "CALL USP_COMPANY_END_ORG_UP('"&org_bonbu&"', '"&cost_year&"', '"&end_month&"', '"&emp_no&"');"
		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	Next
End If

'관리사업부 미지정된 상주 정보 조회
'sql = "select emp_reside_company from emp_master_month where (emp_month ='"&end_month&"') and (mg_saupbu = '') and (emp_reside_company <> '') and (cost_center <> '손익제외') /* group by emp_reside_company */ "
objBuilder.Append "CALL USP_COMPANY_END_RESIDE_SEL('"&end_month&"');"
Set rsReside = DBConn.Execute(objBuilder.ToSTring())
objBuilder.Clear()

If Not rsReside.EOF Then
	arrReside = rsReside.getRows()
End If
rsReside.Close() : Set rsReside = Nothing

If IsArray(arrReside) Then
	For i = LBound(arrReside) To UBound(arrReside ,2)
		emp_no = arrReside(0, i)
		emp_reside_company = arrReside(1, i)
		emp_org_code = arrReside(2, i)

		'sql = "SELECT saupbu FROM trade WHERE trade_name = '"&rs("emp_reside_company")&"'"

		objBuilder.Append "CALL USP_COMPANY_END_RESIDE_UP('"&end_month&"', '"&emp_org_code&"', '"&emp_no&"');"
		DBConn.Execute(objBuilder.ToSTring())
		objBuilder.Clear()
	Next
End If
%>