<%
'sql = "update transit_cost set cost_center = '����������' where (company <> '����' and company <> '��Ÿ' and company <> '���̿��������' and company <> '') and (run_date >='"&from_date&"' and run_date <='"&to_date&"')"
objBuilder.Append "UPDATE transit_cost SET "
objBuilder.Append "	cost_center = '����������' "
objBuilder.Append "WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
objBuilder.Append "	AND company NOT IN ('����', '��Ÿ', '���̿��������', '���̿�', '') "

DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

'sql = "select org_name from transit_cost where (company = '����' or company = '��Ÿ' or company = '���̿��������' or company = '' OR company = '���̿�') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by org_name"
objBuilder.Append "SELECT org_name "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
objBuilder.Append "	AND company IN ('����', '��Ÿ', '���̿��������', '', '���̿�') "
objBuilder.Append "GROUP BY org_name "

Set rsTran = Server.CreateObject("ADODB.RecordSet")
rsTran.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTran.EOF
	'sql = "select org_cost_center from emp_org_mst_month where org_month = '"&end_month&"' and org_name = '"&rs("org_name")&"' group by org_name"
	objBuilder.Append "SELECT org_cost_center "
	objBuilder.Append "FROM emp_org_mst_month "
	objBuilder.Append "WHERE org_month = '"&end_month&"' "
	objBuilder.Append "	AND org_name = '"&rsTran("org_name")&"' "
	objBuilder.Append "GROUP BY org_name "

	Set rsTranOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not(rsTranOrg.BOF Or rsTranOrg.EOF) Then
		'sql = "update transit_cost set cost_center = '"&rs_org("org_cost_center")&"' where (company = '����' or company = '��Ÿ' or company = '���̿��������' or company = '' OR company = '���̿�') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and org_name = '"&rs("org_name")&"'"
		objBuilder.Append "UPDATE transit_cost SET "
		objBuilder.Append "	cost_center = '"&rsTranOrg("org_cost_center")&"' "
		objBuilder.Append "WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
		objBuilder.Append "	AND org_name = '"&rsTran("org_name")&"' "
		objBuilder.Append "	AND company IN ('����', '��Ÿ', '���̿��������', '', '���̿�') "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If
	rsTranOrg.Close()

	rsTran.MoveNext()
Loop
rsTran.Close() : Set rsTran = Nothing

' �����̸鼭 ��Ÿ�� �Է½� ��Ÿ�� ���� ȸ��� ����
'sql = "select org_name from transit_cost where (company = '����' or company = '��Ÿ' or company = '���̿��������' or company = '' OR company = '���̿�') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cost_center = '����������') group by org_name"
objBuilder.Append "SELECT org_name "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
objBuilder.Append "	AND cost_center = '����������' "
objBuilder.Append "	AND company IN ('����', '��Ÿ', '���̿��������', '', '���̿�') "
objBuilder.Append "GROUP BY org_name "

Set rsTranOutCost = Server.CreateObject("ADODB.RecordSet")
rsTranOutCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTranOutCost.EOF
	'sql = "select org_reside_company from emp_org_mst_month where org_month = '"&end_month&"' and org_name = '"&rs("org_name")&"' group by org_name"
	objBuilder.Append "SELECT org_reside_company "
	objBuilder.Append "FROM emp_org_mst_month "
	objBuilder.Append "WHERE org_month = '"&end_month&"' "
	objBuilder.Append "	AND org_name = '"&rsTranOutCost("org_name")&"' "
	objBuilder.Append "GROUP BY org_name "

	set rsTranOutCostOrg = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If Not(rsTranOutCostOrg.bof Or rsTranOutCostOrg.eof) Then
		'sql = "update transit_cost set company = '"&rs_org("org_reside_company")&"' where (company = '����' or company = '��Ÿ' or company = '���̿��������' or company = '' OR company = '���̿�') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (cost_center = '����������') and org_name = '"&rs("org_name")&"'"
		objBuilder.Append "UPDATE transit_cost SET "
		objBuilder.Append "	company = '"&rsTranOutCostOrg("org_reside_company")&"' "
		objBuilder.Append "WHERE (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
		objBuilder.Append "	AND cost_center = '����������' "
		objBuilder.Append "	AND org_name = '"&rsTranOutCost("org_name")&"' "
		objBuilder.Append "	AND company IN ('����', '��Ÿ', '���̿��������', '', '���̿�') "

		DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()
	End If
	rsTranOutCostOrg.Close()

	rsTranOutCost.MoveNext()
Loop
rsTranOutCost.Close() : Set rsTranOutCost = Nothing

' ���������� ��������� ����
'�ŷ�ó ���� ����� ��� x, �������� ����η� ��������� ���� ó��[����ȣ_20211006]
'objBuilder.Append "SELECT bonbu, company "
objBuilder.Append "SELECT run_date, mg_ce_id, run_seq, bonbu, company "
objBuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE cost_center = '����������' "
objBuilder.Append "	AND (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
'objBuilder.Append "GROUP BY bonbu, company "

Set rsTranDeptOutCost = Server.CreateObject("ADODB.RecordSet")
rsTranDeptOutCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Dim rsTranDeptOutCostSales

Do Until rsTranDeptOutCost.EOF
	tradeDept = rsTranDeptOutCost("bonbu")

	objBuilder.Append "SELECT saupbu "
	objBuilder.Append "FROM sales_org "
	objBuilder.Append "WHERE saupbu = '"&tradeDept&"' "
	objBuilder.Append "	AND sales_year='"&cost_year&"' "

	Set rsTranDeptOutCostSales = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	'��������ΰ� ���� ���
	If rsTranDeptOutCostSales.EOF Or rsTranDeptOutCostSales.BOF Then
		'objBuilder.Append "SELECT saupbu "
		'objBuilder.Append "FROM trade "
		'objBuilder.Append "WHERE trade_name = '"&rsTranDeptOutCost("company")&"' "
		objBuilder.Append "SELECT emp_bonbu FROM emp_master_month "
		objBuilder.Append "WHERE emp_no = '"&rsTranDeptOutCost("mg_ce_id")&"' AND emp_month = '"&end_month&"' "

		Set rsTranDeptOutCostTrade = DBConn.Execute(objBuilder.ToString())
		objBuilder.Clear()

		If rsTranDeptOutCostTrade.EOF Or rsTranDeptOutCostTrade.BOF then
			tradeDept = "Error"
		Else
			'tradeDept = rsTranDeptOutCostTrade("saupbu")
			tradeDept = rsTranDeptOutCostTrade("emp_bonbu")
		End If
		rsTranDeptOutCostTrade.Close()
	end if
	rsTranDeptOutCostSales.Close()


	objBuilder.Append "UPDATE transit_cost SET "
	objBuilder.Append "	mg_saupbu = '"&tradeDept&"' "
	'objBuilder.Append "WHERE cost_center = '����������' "
	'objBuilder.Append "	AND (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
	'objBuilder.Append "	AND bonbu = '"&rsTranDeptOutCost("bonbu")&"' "
	'objBuilder.Append "	AND company = '"&rsTranDeptOutCost("company")&"' "
	objBuilder.Append "WHERE run_date = '"&rsTranDeptOutCost("run_date")&"' "
	objBuilder.Append "	AND mg_ce_id = '"&rsTranDeptOutCost("mg_ce_id")&"' "
	objBuilder.Append "	AND run_seq = '"&rsTranDeptOutCost("run_seq")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsTranDeptOutCost.MoveNext()
Loop
rsTranDeptOutCost.Close() : Set rsTranDeptOutCost = Nothing

' ���������� ������ ��������� ����
'sql = "select saupbu from transit_cost where (cost_center = '������') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') group by saupbu"
objBuilder.Append "SELECT bonbu "
objbuilder.Append "FROM transit_cost "
objBuilder.Append "WHERE cost_center = '������' "
objBuilder.Append "	AND (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"') "
objBuilder.Append "GROUP BY bonbu "

Set rsTranCost = Server.CreateObject("ADODB.RecordSet")
rsTranCost.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

Do Until rsTranCost.EOF
	'sql = "update transit_cost set mg_saupbu = '"&rs("org_bonbu")&"' where (cost_center = '������') and (run_date >='"&from_date&"' and run_date <='"&to_date&"') and (mg_ce_id = '"&rs("mg_ce_id")&"')"
	objBuilder.Append "UPDATE transit_cost SET "
	objBuilder.Append "	mg_saupbu = '"&rsTranCost("bonbu")&"'"
	objBuilder.Append "WHERE cost_center = '������' "
	objBuilder.Append "	AND (run_date >= '"&from_date&"' AND run_date <= '"&to_date&"')"
	objBuilder.Append "	AND bonbu = '"&rsTranCost("bonbu")&"' "

	DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	rsTranCost.MoveNext()
Loop
rsTranCost.Close() : Set rsTranCost = Nothing
%>