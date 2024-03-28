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

'조직 별 상주회사 정보 조회
'sql = "select emp_reside_company from emp_master_month where (emp_month ='"&end_month&"') and (mg_saupbu = '') and (emp_reside_company <> '') and (cost_center <> '손익제외') /* group by emp_reside_company */ "
objBuilder.Append "SELECT emp_reside_company "
objBuilder.Append "FROM emp_master_month "
objBuilder.Append "WHERE emp_month = '"&end_month&"' "
objBuilder.Append "	AND mg_saupbu = '' "
objBuilder.Append "	AND emp_reside_company <> '' "
objBuilder.Append "	AND cost_center <> '손익제외' "
objBuilder.Append "	AND emp_pay_id <> '2' "

Set rsReside = Server.CreateObject("ADODB.RecordSet")
rsReside.Open objBuilder.ToSTring(), DBConn, 1
objBuilder.Clear()

Do Until rsReside.EOF
	'sql = "SELECT saupbu FROM trade WHERE trade_name = '"&rs("emp_reside_company")&"'"
	objBuilder.Append "SELECT saupbu "
	objBuilder.Append "FROM trade "
	objBuilder.Append "WHERE trade_name = '"&rsReside("emp_reside_company")&"' "

	Set rsResideTrade = DBConn.Execute(objBuilder.ToSTring())
	objBuilder.Clear()

	If rsResideTrade.EOF Or rsResideTrade.BOF Then
		'saupbu = "Error"
		trade_bonbu = "Error"
	Else
		'saupbu = rs_trade("saupbu")
		trade_bonbu = rsResideTrade("saupbu")
	End If
	rsResideTrade.Close()

	'sql = "update emp_master_month set mg_saupbu = '"&saupbu&"' where emp_month ='"&end_month&"' and mg_saupbu = '' and emp_reside_company = '"&rs("emp_reside_company")&"'"
	objBuilder.Append "UPDATE emp_master_month SET "
	objBuilder.Append "	mg_saupbu = '"&trade_bonbu&"' "
	objBuilder.Append "WHERE emp_month ='"&end_month&"' "
	objBuilder.Append "	AND mg_saupbu = '' "
	objBuilder.Append "	AND emp_reside_company = '"&rsReside("emp_reside_company")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'sql = "update pay_month_give set mg_saupbu = '"&saupbu&"' where pmg_yymm ='"&end_month&"' and mg_saupbu = '' and pmg_reside_company = '"&rs("emp_reside_company")&"'"
	objBuilder.Append "UPDATE pay_month_give SET "
	objBuilder.Append "	mg_saupbu = '"&trade_bonbu&"' "
	objBuilder.Append "WHERE pmg_yymm ='"&end_month&"' "
	objBuilder.Append "	AND mg_saupbu = '' "
	objBuilder.Append "	AND pmg_reside_company = '"&rsReside("emp_reside_company")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsReside.MoveNext()
Loop
rsReside.Close() : Set rsReside = Nothing
%>