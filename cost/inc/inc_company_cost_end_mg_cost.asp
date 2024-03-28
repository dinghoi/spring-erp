<%
' �ʱⰪ Clear
'sql = "update general_cost set mg_saupbu = '', cost_center = '' where (tax_bill_yn = 'N') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') "
objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	mg_saupbu = '', "
objBuilder.Append "	cost_center = '' "
objBuilder.Append "WHERE (tax_bill_yn = 'N') "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' ���ݰ�꼭�� �Է½� ��������� �����ϰ� ����
'sql = "update general_cost set cost_center = '' where (tax_bill_yn = 'Y') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	cost_center = '' "
objBuilder.Append "WHERE tax_bill_yn = 'Y' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"')"

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' ������� ����
'sql = "update general_cost set cost_center = '����������' where (pl_yn = 'Y') and (company <> '����' and company <> '����' and company <> '�ι�' and company <> '��Ÿ' and company <> '����' and company <> '���̿��������' and company <> '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"')"
objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	cost_center = '����������' "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"')"
objBuilder.Append "	AND company NOT IN ('����', '����', '�ι�', '��Ÿ', '����', '���̿��������', '���̿�', '') "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

' ����� ������� ����(���)
'sql = "select emp_company,org_name from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '����' or company = '����' or company = '�ι�' or company = '��Ÿ' or company = '����' or company = '���̿��������' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,org_name"

objBuilder.Append "SELECT slip_date, slip_seq, org_name, company "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' AND tax_bill_yn = 'N' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company IN ('����', '����', '�ι�', '��Ÿ', '����', '���̿��������', '���̿�', '') "

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
		cost_center = "��������"
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

	'sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '����' or company = '��Ÿ' or company = '����' or company = '���̿��������' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = '"&cost_center&"' "
	objBuilder.Append "WHERE slip_date = '"&rsNoTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsNoTax("slip_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'//2017-06-19 �����(����/�ι�) ���� �߰�
	'sql = "update general_cost set cost_center = (case when company='����' then '��������' when company='�ι�' then '�ι������' end) where (pl_yn = 'Y') and (tax_bill_yn = 'N') and (company = '����' or company = '�ι�') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = (CASE WHEN company='����' THEN '��������' WHEN company='�ι�' THEN '�ι������' END) "
	objBuilder.Append "WHERE slip_date = '"&rsNoTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsNoTax("slip_seq")&"' "
	objBuilder.Append "	AND company IN ('����', '�ι�'); "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsNoTax.MoveNext()
Loop
rsNoTax.Close() : Set rsNoTax = Nothing

' ����� ��� �������� ( ���ݰ�꼭 )
' ��������� �ִ°��
'sql = "select emp_company, mg_saupbu from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '����' or company = '����' or company = '�ι�' or company = '��Ÿ' or company = '����' or company = '���̿��������' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,mg_saupbu"

objBuilder.Append "UPDATE general_cost SET "
objBuilder.Append "	mg_saupbu = bonbu "
objBuilder.Append "WHERE pl_yn = 'Y' AND tax_bill_yn = 'Y' AND mg_saupbu NOT IN ('', bonbu) "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company = '����' "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

objBuilder.Append "SELECT slip_date, slip_seq, org_name, mg_saupbu, company "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' AND tax_bill_yn = 'Y' AND mg_saupbu <> '' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company IN ('����', '����', '�ι�', '��Ÿ', '����', '���̿��������', '���̿�', '') "

Set rsTax = Server.CreateObject("ADODB.RecordSet")
rsTax.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTax.EOF
	'sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_company = '"&rs("emp_company")&"' and org_name = '"&rs("mg_saupbu")&"'"
	objBuilder.Append "SELECT org_cost_center "
	objBuilder.Append "FROM emp_org_mst_month "
	objBuilder.Append "WHERE org_month = '"&end_month&"' "

	'If rsTax("company") = "����" Then
	'	objBuilder.Append "	AND org_name = '"&rsTax("org_name")&"' "
	'Else
		objBuilder.Append "	AND org_name = '"&rsTax("mg_saupbu")&"' "
	'End If


	Set rsTaxOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If rsTaxOrg.EOF Or rsTaxOrg.BOF Then
		cost_center = "��������"
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

	'sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '����' or company = '��Ÿ' or company = '����' or company = '���̿��������' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (mg_saupbu = '"&rs("mg_saupbu")&"')"

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = '"&cost_center&"' "
	objBuilder.Append "WHERE slip_date = '"&rsTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTax("slip_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'//2017-06-19 �����(����/�ι�) ���� �߰�
	'sql = "update general_cost set cost_center = (case when company='����' then '��������' when company='�ι�' then '�ι������' end) where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu <> '') and (company = '����' or company = '�ι�') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (mg_saupbu = '"&rs("mg_saupbu")&"')"

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = (case when company='����' then '��������' when company='�ι�' then '�ι������' end) "
	objBuilder.Append "WHERE slip_date = '"&rsTax("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTax("slip_seq")&"' "
	objBuilder.Append "	AND company IN ('����', '�ι�') "

	DBConn.Execute(objBuilder.tostring())
	objBuilder.Clear()

	rsTax.MoveNext()
Loop
rsTax.Close() : Set rsTax = Nothing

' ��������ΰ� ���°��
'sql = "select emp_company,org_name from general_cost where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '����' or company = '����' or company = '�ι�' or company = '��Ÿ' or company = '����' or company = '���̿��������' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') group by emp_company,org_name"

objBuilder.Append "SELECT slip_date, slip_seq, org_name "
objBuilder.Append "FROM general_cost "
objBuilder.Append "WHERE pl_yn = 'Y' "
objBuilder.Append "	AND tax_bill_yn = 'Y' "
objBuilder.Append "	AND mg_saupbu = '' "
objBuilder.Append "	AND (slip_date >='"&from_date&"' AND slip_date <='"&to_date&"') "
objBuilder.Append "	AND company IN ('����', '����', '�ι�', '��Ÿ', '����', '���̿��������', '���̿�', '') "

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
		cost_center = "��������"
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

	'sql = "update general_cost set cost_center = '"&cost_center&"' where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '����' or company = '��Ÿ' or company = '����' or company = '���̿��������' or company = '') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"

	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = '"&cost_center&"' "
	objBuilder.Append "WHERE slip_date = '"&rsTaxNoMg("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTaxNoMg("slip_seq")&"' "

	DBConn.Execute(objBuilder.tostring())
	objBuilder.Clear()

	'//2017-06-19 �����(����/�ι�) ���� �߰�
	'sql = "update general_cost set cost_center = (case when company='����' then '��������' when company='�ι�' then '�ι������' end) where (pl_yn = 'Y') and (tax_bill_yn = 'Y') and (mg_saupbu = '') and (company = '����' or company = '�ι�') and (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (emp_company = '"&rs("emp_company")&"') and (org_name = '"&rs("org_name")&"')"
	objBuilder.Append "UPDATE general_cost SET "
	objBuilder.Append "	cost_center = (CASE WHEN company='����' THEN '��������' WHEN company='�ι�' THEN '�ι������' END) "
	objBuilder.Append "WHERE slip_date = '"&rsTaxNoMg("slip_date")&"' "
	objBuilder.Append "	AND slip_seq = '"&rsTaxNoMg("slip_seq")&"' "
	objBuilder.Append "	AND company IN ('����', '�ι�') "

	DBConn.Execute(objBuilder.tostring())
	objBuilder.Clear()

	rsTaxNoMg.MoveNext()
Loop

rsTaxNoMg.Close() : Set rsTaxNoMg = Nothing

%>