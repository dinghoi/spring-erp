<%
'### 인사마스터 및 급여DATA에 관리사업부 지정 ###

'영업 사업부 관리 사업부 지정
'sql = "select emp_saupbu from emp_master_month where (emp_month ='"&end_month&"') and (cost_center <> '손익제외') /* group by emp_saupbu */ "
objBuilder.Append "SELECT emmt.emp_no, eomt.org_bonbu, eomt.org_code "
objBuilder.Append "FROM emp_master_month AS emmt "
objBuilder.Append "INNER JOIN emp_org_mst_month AS eomt ON emmt.emp_org_code = eomt.org_code "
objBuilder.Append "	AND eomt.org_month ='"&end_month&"' "
objBuilder.Append "WHERE emmt.emp_month ='"&end_month&"' "
objBuilder.Append "	AND emmt.cost_center <> '손익제외' "

Set rsEmp = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

Do Until rsEmp.EOF
	'saupbu = rs("emp_saupbu")
	org_bonbu = rsEmp("org_bonbu")
	org_code = rsEmp("org_code")
	emp_no = rsEmp("emp_no")

	'sql = "select sort_seq from sales_org where saupbu = '"&saupbu&"' and sales_year='" & cost_year & "' "
	objBuilder.Append "SELECT sort_seq "
	objBuilder.Append "FROM sales_org "
	objBuilder.Append "WHERE saupbu = '"&org_bonbu&"' "
	objBuilder.Append "	AND sales_year='" & cost_year & "' "

	Set rsEmpSales = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsEmpSales.EOF Or rsEmpSales.BOF Then
		'saupbu = ""
		org_bonbu = ""
	End If
	rsEmpSales.Close()

	'sql = "update emp_master_month set mg_saupbu = '"&saupbu&"' where emp_month ='"&end_month&"' and emp_saupbu = '"&rs("emp_saupbu")&"'"
	objBuilder.Append "UPDATE emp_master_month SET "
	objBuilder.Append "	mg_saupbu = '"&org_bonbu&"' "
	objBuilder.Append "WHERE emp_month ='"&end_month&"' "
	objBuilder.Append "	AND emp_no = '"&emp_no&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'sql = "update pay_month_give set mg_saupbu = '"&saupbu&"' where pmg_yymm ='"&end_month&"' and pmg_saupbu = '"&rs("emp_saupbu")&"'"
	objBuilder.Append "UPDATE pay_month_give SET "
	objBuilder.Append "	mg_saupbu = '"&org_bonbu&"' "
	objBuilder.Append "WHERE pmg_yymm = '"&end_month&"' "
	objBuilder.Append "	AND pmg_emp_no = '"&emp_no&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsEmp.MoveNext()
Loop
Set rsEmpSales = Nothing
rsEmp.Close() : Set rsEmp = Nothing

'상주 인원 관리사업부 지정
'sql = "select emp_reside_company from emp_master_month where (emp_month ='"&end_month&"') and (mg_saupbu = '') and (emp_reside_company <> '') and (cost_center <> '손익제외') /* group by emp_reside_company */ "

'상주 회사 > 조직 코드로 변경[허정호_20220213]
objBuilder.Append "SELECT emp_reside_company "
objBuilder.Append "FROM emp_master_month "
objBuilder.Append "WHERE emp_month = '"&end_month&"' "
objBuilder.Append "	AND mg_saupbu = '' "
objBuilder.Append "	AND emp_reside_company <> '' "
objBuilder.Append "	AND cost_center <> '손익제외' "
objBuilder.Append "	AND emp_pay_id <> '2' "

Set rsReside = DBConn.Execute(objBuilder.ToSTring())
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
Set rsResideTrade = Nothing
rsReside.Close() : Set rsReside = Nothing
%>