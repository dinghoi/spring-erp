<%
' 초기값 Clear
'sql = "update general_cost set mg_saupbu = '', cost_center = '' where (tax_bill_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	mg_saupbu = '', "
objBuilder.Append "	cost_center = '' "
objBuilder.Append "WHERE (tax_bill_yn = 'N') "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 세금계산서는 입력시 관리사업부 지정하게 변경
'sql = "update general_cost set cost_center = '' where (tax_bill_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	cost_center = '' "
objBuilder.Append "WHERE tax_bill_yn = 'Y' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"')"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 비용유형 셋팅
'sql = "update general_cost set cost_center = '상주직접비' where (pl_yn = 'Y') and (company <> '공통' and company <> '전사' and company <> '부문' and company <> '기타' and company <> '본사' and company <> '케이원정보통신' and company <> '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	cost_center = '상주직접비' "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"')"
objBuilder.Append "	AND company NOT IN ('공통', '전사', '부문', '기타', '본사', '케이원정보통신', '케이원', '') "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' 공통비 비용유형 세팅(비용)
'sql = "select emp_company,org_name from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '공통' or company = '전사' or company = '부문' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,org_name"

objBuilder.Append "SELECT slip_date, slip_seq, org_name, company "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' AND tax_bill_yn = 'N' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company IN ('공통', '전사', '부문', '기타', '본사', '케이원정보통신', '케이원', '') "

Set rsNoTax = Server.CreateObject("ADODB.RecordSet")
rsNoTax.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsNoTax.EOF
	'sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("emp_company")&"' and org_name = '"&rs("org_name")&"'"

	objBuilder.Append "SELECT org_cost_center "
	objBuilder.Append "FROM emp_org_mst_month "
	objBuilder.Append "WHERE org_month = '"&end_month&"' "
	objBuilder.Append "	AND org_name = '"&rsNoTax("org_name")&"' "

	Set rsNoTaxOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsNoTaxOrg.EOF Or rsNoTaxOrg.BOF Then
		cost_center = "전사공통비"
		cost_company = ""
		group_name = ""
		bill_trade_name = ""
	Else
		cost_center = rsNoTaxOrg("org_cost_center")
		cost_company = ""
		group_name = ""
		bill_trade_name = ""
	End If
	rsNoTaxOrg.Close()

	'sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = '"&cost_center&"' "
	objBuilder.Append "WHERE slip_date = '"&rsNoTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsNoTax("slip_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'//2017-06-19 공통비(전사/부문) 로직 추가
	'sql = "update general_cost set cost_center = (case when company='전사' then '전사공통비' when company='부문' then '부문공통비' end) where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '전사' or company = '부문') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = (CASE WHEN company='전사' THEN '전사공통비' WHEN company='부문' THEN '부문공통비' END) "
	objBuilder.Append "WHERE slip_date = '"&rsNoTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsNoTax("slip_seq")&"' "
	objBuilder.Append "	AND company IN ('전사', '부문'); "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsNoTax.MoveNext()
Loop
rsNoTax.Close() : Set rsNoTax = Nothing

' 공통비 비용 유형세팅 ( 세금계산서 )
' 관리사업부 있는경우
'sql = "select emp_company, mg_saupbu from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '공통' or company = '전사' or company = '부문' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,mg_saupbu"

objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	mg_saupbu = bonbu "
objBuilder.Append "WHERE pl_yn = 'Y' AND tax_bill_yn = 'Y' AND mg_saupbu NOT IN ('', bonbu) "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company = '공통' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "SELECT slip_date, slip_seq, org_name, mg_saupbu, company "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' AND tax_bill_yn = 'Y' AND mg_saupbu <> '' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company IN ('공통', '전사', '부문', '기타', '본사', '케이원정보통신', '케이원', '') "

Set rsTax = Server.CreateObject("ADODB.RecordSet")
rsTax.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTax.EOF
	'sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("emp_company")&"' and org_name = '"&rs("mg_saupbu")&"'"
	objBuilder.Append "SELECT org_cost_center "
	objBuilder.Append "FROM emp_org_mst_month "
	objBuilder.Append "WHERE org_month = '"&end_month&"' "

	'If rsTax("company") = "공통" Then
	'	objBuilder.Append "	AND org_name = '"&rsTax("org_name")&"' "
	'Else
		objBuilder.Append "	AND org_name = '"&rsTax("mg_saupbu")&"' "
	'End If


	Set rsTaxOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsTaxOrg.EOF Or rsTaxOrg.BOF Then
		cost_center = "전사공통비"
		cost_company = ""
		group_name = ""
		bill_trade_name = ""
	Else
		cost_center = rsTaxOrg("org_cost_center")
		cost_company = ""
		group_name = ""
		bill_trade_name = ""
	End If
	rsTaxOrg.Close()

	'sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (mg_saupbu = '"&rs("mg_saupbu")&"')"

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = '"&cost_center&"' "
	objBuilder.Append "WHERE slip_date = '"&rsTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTax("slip_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'//2017-06-19 공통비(전사/부문) 로직 추가
	'sql = "update general_cost set cost_center = (case when company='전사' then '전사공통비' when company='부문' then '부문공통비' end) where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '전사' or company = '부문') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (mg_saupbu = '"&rs("mg_saupbu")&"')"

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = (case when company='전사' then '전사공통비' when company='부문' then '부문공통비' end) "
	objBuilder.Append "WHERE slip_date = '"&rsTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTax("slip_seq")&"' "
	objBuilder.Append "	AND company IN ('전사', '부문') "

	DBConn.Execute(objBuilder.tostring())
	objBuilder.Clear()

	rsTax.MoveNext()
Loop
rsTax.Close() : Set rsTax = Nothing

' 관리사업부가 없는경우
'sql = "select emp_company,org_name from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '공통' or company = '전사' or company = '부문' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,org_name"

objBuilder.Append "SELECT slip_date, slip_seq, org_name "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND tax_bill_yn = 'Y' "
objBuilder.Append "	AND mg_saupbu = '' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company IN ('공통', '전사', '부문', '기타', '본사', '케이원정보통신', '케이원', '') "

Set rsTaxNoMg = Server.CreateObject("ADODB.RecordSet")
rsTaxNoMg.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTaxNoMg.EOF
	'sql = "select * from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("emp_company")&"' and org_name = '"&rs("org_name")&"'"

	objBuilder.Append "SELECT org_cost_center "
	objBuilder.Append "FROM emp_org_mst_month "
	objBuilder.Append "WHERE org_month = '"&end_month&"' "
	objBuilder.Append "	AND org_name = '"&rsTaxNoMg("org_name")&"' "

	Set rsTaxNoMgOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsTaxNoMgOrg.EOF Or rsTaxNoMgOrg.BOF Then
		cost_center = "전사공통비"
		cost_company = ""
		group_name = ""
		bill_trade_name = ""
	Else
		cost_center = rsTaxNoMgOrg("org_cost_center")
		cost_company = ""
		group_name = ""
		bill_trade_name = ""
	End If
	rsTaxNoMgOrg.Close()

	'sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '공통' or company = '기타' or company = '본사' or company = '케이원정보통신' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = '"&cost_center&"' "
	objBuilder.Append "WHERE slip_date = '"&rsTaxNoMg("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTaxNoMg("slip_seq")&"' "

	DBConn.Execute(objBuilder.tostring())
	objBuilder.Clear()

	'//2017-06-19 공통비(전사/부문) 로직 추가
	'sql = "update general_cost set cost_center = (case when company='전사' then '전사공통비' when company='부문' then '부문공통비' end) where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '전사' or company = '부문') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = (CASE WHEN company='전사' THEN '전사공통비' WHEN company='부문' THEN '부문공통비' END) "
	objBuilder.Append "WHERE slip_date = '"&rsTaxNoMg("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTaxNoMg("slip_seq")&"' "
	objBuilder.Append "	AND company IN ('전사', '부문') "

	DBConn.Execute(objBuilder.tostring())
	objBuilder.Clear()

	rsTaxNoMg.MoveNext()
Loop

rsTaxNoMg.Close() : Set rsTaxNoMg = Nothing

%>