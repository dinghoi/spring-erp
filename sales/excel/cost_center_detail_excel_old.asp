<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--include virtual="/include/db_create.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim cost_month, sales_saupbu, slip_month, title_line, savefilename, i

cost_month = f_Request("cost_month")
sales_saupbu = f_Request("sales_saupbu")

If sales_saupbu = "��Ÿ�����" Then
	sales_saupbu = ""
End If

slip_month = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2)

title_line = cost_month & "�� " & sales_saupbu & " ��뼼�� ����"
savefilename = title_line & ".xls"

Response.Buffer = True
Response.ContentType = "appllication/vnd.ms-excel" '// ������ ����
Response.CacheControl = "public"
Response.AddHeader "Content-Disposition","attachment; filename=" &savefilename

i = 0

Dim insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per
Dim rs_etc, arrEtc

'sql = "select * from insure_per where insure_year = '"&mid(cost_month,1,4)&"'"
objBuilder.Append "SELECT insure_tot_per, income_tax_per, annual_pay_per, retire_pay_per "
objBuilder.Append "FROM insure_per WHERE insure_year = '"&Mid(cost_month, 1, 4)&"' "

Set rs_etc = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

arrEtc = rs_etc.getRows()
rs_etc.close() : Set rs_etc = Nothing

insure_tot_per = arrEtc(0, 0)
income_tax_per = arrEtc(1, 0)
annual_pay_per = arrEtc(2, 0)
retire_pay_per = arrEtc(3, 0)
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��뱸��</th>
								<th scope="col">��������</th>
								<th scope="col">�������</th>
								<th scope="col">��翵�������</th>
								<th scope="col">��꼭����</th>
								<th scope="col">���ȸ��</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��</th>
								<th scope="col">������</th>
								<th scope="col">����ó</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">��������</th>
								<th scope="col">�������</th>
								<th scope="col">���־�ü</th>
								<th scope="col">�հ�</th>
								<th scope="col">���ް���</th>
								<th scope="col">�ΰ���</th>
								<th scope="col">���೻��</th>
							</tr>
						</thead>
						<tbody>
						<%
						Dim tot_tax, rsPay, tax_bill_yn, gubun, account, cost_center, mg_saupbu
						Dim company, emp_name, slip_date, slip_seq, customer, price, cost, cost_vat
						Dim slip_memo, arrPay, j, pmg_id

						tot_tax = 0

						If (saupbu = sales_saupbu And position = "�������") Or (saupbu = sales_saupbu And position = "������") Or sales_grade = "0" Then
							'�޿�
							objBuilder.Append "SELECT pmg_id, cost_center, mg_saupbu, pmg_company, pmg_bonbu, pmg_saupbu, pmg_team, "
							objBuilder.Append "	pmg_org_name, pmg_reside_place, pmg_reside_company, pmg_emp_name, pmg_yymm, pmg_emp_no, "
							objBuilder.Append "	pmg_give_total "
							objBuilder.Append "FROM pay_month_give "
							objBuilder.Append "WHERE pmg_yymm = '"&cost_month&"' AND pmg_id <> '4' AND (cost_center <> '��������') "

							If sales_saupbu = "��������" Or sales_saupbu = "�ι������" Then
								'sql = "select * from pay_month_give where pmg_id <>'4' and cost_center ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') ORDER BY pmg_id, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_name, pmg_reside_place, pmg_reside_company, pmg_emp_name"

								objBuilder.Append "	AND cost_center = '"&sales_saupbu&"' "
								objBuilder.Append "	AND pmg_yymm = '"&cost_month&"' "
							Else
								'sql = "select * from pay_month_give where pmg_id <>'4' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') ORDER BY pmg_id, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_name, pmg_reside_place, pmg_reside_company, pmg_emp_name"
								objBuilder.Append "	AND (cost_center ='������' or cost_center ='����������') "
								objBuilder.Append "	AND mg_saupbu ='"&sales_saupbu&"' "
							End If
							objBuilder.Append "ORDER BY pmg_id, pmg_bonbu, pmg_saupbu, pmg_team, pmg_org_name, pmg_reside_place, pmg_reside_company, pmg_emp_name "

							'Rs.Open Sql, Dbconn, 1
							Set rsPay = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							If Not rsPay.EOF Then
								arrPay = rsPay.getRows()
							End If
							rsPay.Close() : Set rsPay = Nothing

							'Do Until rsPay.EOF
							If IsArray(arrPay) Then
								tax_bill_yn = "�Ϲ�"
								gubun = "�ΰǺ�"
								account = "������"
								customer = ""
								slip_memo = ""

								For j = LBound(arrPay) To UBound(arrPay, 2)
									pmg_id = arrPay(0, j)

									if pmg_id = "1" then
										account = "�޿�"
									end If

									if pmg_id = "2" then
										account = "��"
									end If

									if pmg_id = "3" then
										account = "��õ�μ�Ƽ��"
									end If

									cost_center  = arrPay(1, j)
									mg_saupbu    = arrPay(2, j)
									emp_company  = arrPay(3, j)
									bonbu        = arrPay(4, j)
									saupbu       = arrPay(5, j)
									team         = arrPay(6, j)
									org_name     = arrPay(7, j)
									reside_place = arrPay(8, j)
									company      = arrPay(9, j)
									emp_name     = arrPay(10, j)
									slip_date    = arrPay(11, j)
									slip_seq     = arrPay(12, j)
									price        = arrPay(13, j)
									cost         = arrPay(13, j)
									cost_vat     = 0

									i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price, 0)%></td>
							  	<td class="right"><%=formatnumber(cost, 0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat, 0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
								Next
							End If

							'4�� �����
							Dim rsInsure, insure_tot, income_tax, annual_pay, retire_pay, arrInsure
							Dim pmg_company, tot_cost, base_pay, meals_pay, overtime_pay, tax_no

							objBuilder.Append "SELECT cost_center, pmg_company, sum(pmg_give_total) as tot_cost, sum(pmg_base_pay) as base_pay, "
							objBuilder.Append "	sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay, sum(pmg_tax_no) as tax_no "
							objBuilder.Append "FROM pay_month_give "
							objBuilder.Append "WHERE pmg_id = '1' AND pmg_yymm = '"&cost_month&"' "
							objBuilder.Append "	AND cost_center <> '��������' "

							if sales_saupbu = "��������" or sales_saupbu = "�ι������" Then
								'sql = "select cost_center,pmg_company,sum(pmg_give_total) as tot_cost,sum(pmg_base_pay) as base_pay,sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay,sum(pmg_tax_no) as tax_no from pay_month_give where pmg_id = '1' and cost_center ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') GROUP BY pmg_id, pmg_company ORDER BY pmg_company"
								objBuilder.Append "	AND cost_center ='"&sales_saupbu&"' "
							else
								'sql = "select cost_center,pmg_company,sum(pmg_give_total) as tot_cost,sum(pmg_base_pay) as base_pay,sum(pmg_meals_pay) as meals_pay,sum(pmg_overtime_pay) as overtime_pay,sum(pmg_tax_no) as tax_no from pay_month_give where pmg_id = '1' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and pmg_yymm = '"&cost_month&"' and (cost_center <> '��������') GROUP BY pmg_id, cost_center,pmg_company ORDER BY pmg_company"
								objBuilder.Append "	AND (cost_center ='������' or cost_center ='����������') AND  mg_saupbu ='"&sales_saupbu&"' "
							end If

							objBuilder.Append "GROUP BY pmg_id, cost_center,pmg_company ORDER BY pmg_company "

							'Rs.Open Sql, Dbconn, 1
							Set rsInsure = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							If Not rsInsure.EOF Then
								arrInsure = rsInsure.getRows()
							End If
							rsInsure.close() : Set rsInsure = Nothing

							IF IsArray(arrInsure) Then
								For j = LBound(arrInsure) To UBound(arrInsure, 2)
									cost_center = arrInsure(0, j)
									pmg_company = arrInsure(1, j)
									tot_cost = arrInsure(2, j)
									base_pay = arrInsure(3, j)
									meals_pay = arrInsure(4, j)
									overtime_pay = arrInsure(5, j)
									tax_no = arrInsure(6, j)

									'insure_tot = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * insure_tot_per / 100)
									insure_tot = clng((clng(tot_cost)) * insure_tot_per / 100)
									'income_tax = clng((clng(rs("tot_cost")) - clng(rs("tax_no"))) * income_tax_per / 100)
									income_tax = clng((clng(tot_cost)) * income_tax_per / 100)
									annual_pay = clng((clng(base_pay)+clng(meals_pay)+clng(overtime_pay)) * annual_pay_per / 100)
									retire_pay = clng((clng(base_pay)+clng(meals_pay)+clng(overtime_pay)) * retire_pay_per / 100)

									i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>4�뺸���</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=pmg_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(insure_tot,0)%></td>
							  	<td class="right"><%=formatnumber(insure_tot,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>�ҵ漼��������</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=pmg_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(income_tax,0)%></td>
							  	<td class="right"><%=formatnumber(income_tax,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>��������</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=pmg_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(annual_pay,0)%></td>
							  	<td class="right"><%=formatnumber(annual_pay,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
							<tr>
								<td class="first"><%=i%></td>
								<td>�ΰǺ�</td>
								<td>��������</td>
								<td><%=cost_center%></td>
								<td></td>
								<td>�Ϲ�</td>
								<td><%=pmg_company%></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td></td>
								<td><%=cost_month%></td>
								<td></td>
								<td></td>
							  	<td class="right"><%=formatnumber(retire_pay,0)%></td>
							  	<td class="right"><%=formatnumber(retire_pay,0)%></td>
							  	<td class="right">0</td>
								<td></td>
							</tr>
						<%
								Next
							End If
						End If

						'�˹ٺ�
						Dim rsAlba, arrAlba

						objBuilder.Append "SELECT cost_center, mg_saupbu, company, bonbu, saupbu, team, org_name, "
						objBuilder.Append "	cost_company, draft_man, give_date, draft_no, alba_give_total, draft_tax_id "
						objBuilder.Append "FROM pay_alba_cost "
						objBuilder.Append "WHERE rever_yymm = '"&cost_month&"' "

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							'sql = "select * from pay_alba_cost where cost_center ='"&sales_saupbu&"' and rever_yymm = '"&cost_month&"' ORDER BY cost_center,give_date,mg_saupbu,org_name, draft_man"
							objBuilder.Append "	AND cost_center = '"&sales_saupbu&"' "
						else
							'sql = "select * from pay_alba_cost where mg_saupbu ='"&sales_saupbu&"' and (cost_center ='������' or cost_center ='����������') and rever_yymm = '"&cost_month&"' ORDER BY cost_center,give_date,mg_saupbu,org_name, draft_man"
							objBuilder.Append "	AND mg_saupbu ='"&sales_saupbu&"' and (cost_center ='������' or cost_center ='����������') "
						end If
						objBuilder.Append "ORDER BY cost_center,give_date,mg_saupbu,org_name, draft_man "

						'Rs.Open Sql, Dbconn, 1
						Set rsAlba = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsAlba.EOF Then
							arrAlba = rsAlba.getRows()
						End If

						rsAlba.close() : Set rsAlba = Nothing

						If IsArray(arrAlba) Then
							tax_bill_yn  = "�Ϲ�"
							gubun        = "�ΰǺ�"
							account      = "�˹ٺ�"
							reside_place = ""
							customer     = ""
							cost_vat     = 0

							For j = LBound(arrAlba) To UBound(arrAlba, 2)
								cost_center  = arrAlba(0, j)
								mg_saupbu    = arrAlba(1, j)
								emp_company  = arrAlba(2, j)
								bonbu        = arrAlba(3, j)
								saupbu       = arrAlba(4, j)
								team         = arrAlba(5, j)
								org_name     = arrAlba(6, j)
								company      = arrAlba(7, j)
								emp_name     = arrAlba(8, j)
								slip_date    = arrAlba(9, j)
								slip_seq     = arrAlba(10, j)
								price        = arrAlba(11, j)
								cost         = arrAlba(11, j)
								slip_memo    = arrAlba(12, j)

								i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							Next
						End If

						'�Ϲ� ���
						Dim rsComm, arrComm

						objBuilder.Append "SELECT tax_bill_yn, slip_gubun, account, cost_center, mg_saupbu, emp_company, "
						objBuilder.Append "	bonbu, saupbu, team, org_name, reside_place, company, emp_name, slip_date, "
						objBuilder.Append "	slip_seq, customer, price, cost, cost_vat, slip_memo "
						objBuilder.Append "FROM general_cost "
						objBuilder.Append "WHERE pl_yn = 'Y' AND cancel_yn = 'N' AND SUBSTRING(slip_date, 1, 7) = '"&slip_month&"' "

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							'sql = "select * from general_cost where (pl_yn = 'Y') and cancel_yn ='N' and cost_center ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
							objBuilder.Append "	AND cost_center = '"&sales_saupbu&"' "
						  else
							'sql = "select * from general_cost where (pl_yn = 'Y') and cancel_yn ='N' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
							objBuilder.Append "	AND (cost_center ='������' or cost_center ='����������') AND mg_saupbu = '"&sales_saupbu&"' "
						end If
						objBuilder.Append "ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name "

						'Rs.Open Sql, Dbconn, 1
						Set rsComm = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsComm.EOF Then
							arrComm = rsComm.getRows()
						End If
						rsComm.close() : Set rsComm = Nothing

						If IsArray(arrComm) Then
							For j = LBound(arrComm) To UBound(arrComm, 2)
								tax_bill_yn = arrComm(0, j)

								if tax_bill_yn = "Y" then
									tax_bill_yn = "���ݰ�꼭"
								else
									tax_bill_yn = "�Ϲ�"
								end If

								gubun        = arrComm(1, j)
								account      = arrComm(2, j)
								cost_center  = arrComm(3, j)
								mg_saupbu    = arrComm(4, j)
								emp_company  = arrComm(5, j)
								bonbu        = arrComm(6, j)
								saupbu       = arrComm(7, j)
								team         = arrComm(8, j)
								org_name     = arrComm(9, j)
								reside_place = arrComm(10, j)
								company      = arrComm(11, j)
								emp_name     = arrComm(12, j)
								slip_date    = arrComm(13, j)
								slip_seq     = arrComm(14, j)
								customer     = arrComm(15, j)
								price        = arrComm(16, j)
								cost         = arrComm(17, j)
								cost_vat     = arrComm(18, j)
								slip_memo    = arrComm(19, j)

								i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							Next
						End If

						'�����
						Dim rsTran, arrTran, somopum, oil_price, fare, parking, toll

						objBuilder.Append "SELECT car_owner, cost_center, mg_saupbu, emp_company, bonbu, saupbu, team, "
						objBuilder.Append "	org_name, reside_place, company, user_name, run_date, run_seq, "
						objBuilder.Append "	somopum, oil_price, fare, parking, toll, run_memo "
						objBuilder.Append "FROM transit_cost "
						objBuilder.Append "WHERE cancel_yn ='N' AND SUBSTRING(run_date, 1, 7) = '"&slip_month&"' "

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							'sql = "select * from transit_cost where cancel_yn ='N' and cost_center ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
							objBuilder.Append "	AND cost_center ='"&sales_saupbu&"' "
						else
							'sql = "select * from transit_cost where cancel_yn ='N' and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
							objBuilder.Append "	AND (cost_center ='������' or cost_center ='����������') AND mg_saupbu ='"&sales_saupbu&"' "
						end If
						objBuilder.Append "ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name "

						'Rs.Open Sql, Dbconn, 1
						Set rsTran = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsTran.EOF Then
							arrTran = rsTran.getRows()
						End If
						rsTran.close() : Set rsTran = Nothing

						If IsArray(arrTran) Then
							tax_bill_yn  = "�Ϲ�"
							gubun        = "�����"
							customer     = ""
							cost_vat     = 0

							For j = LBound(arrTran) To UBound(arrTran, 2)
								account      = arrTran(0, j)
								cost_center  = arrTran(1, j)
								mg_saupbu    = arrTran(2, j)
								emp_company  = arrTran(3, j)
								bonbu        = arrTran(4, j)
								saupbu       = arrTran(5, j)
								team         = arrTran(6, j)
								org_name     = arrTran(7, j)
								reside_place = arrTran(8, j)
								company      = arrTran(9, j)
								emp_name     = arrTran(10, j)
								slip_date    = arrTran(11, j)
								slip_seq     = arrTran(12, j)

								somopum = arrTran(13, j)
								oil_price = arrTran(14, j)
								fare = arrTran(15, j)
								parking = arrTran(16, j)
								toll = arrTran(17, j)

								'price        = arrTran(13, j) + arrTran(14, j) + arrTran(15, j) + arrTran(16, j) + arrTran(17, j)
								'cost         = arrTran(13, j) + arrTran(14, j) + arrTran(15, j) + arrTran(16, j) + arrTran(17, j)
								price = somopum + oil_price + fare + parking + toll
								cost = price

								slip_memo    = arrTran(18, j)

								i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							Next
						End If

						'����������
						Dim rsRepair, arrRepair

						objBuilder.Append "SELECT cost_center, mg_saupbu, emp_company, bonbu, saupbu, team, "
						objBuilder.Append "	org_name, reside_place, company, user_name, run_date, run_seq, repair_cost "
						objBuilder.Append "FROM transit_cost "
						objBuilder.Append "WHERE cancel_yn ='N' AND repair_cost > 0 AND SUBSTRING(run_date, 1, 7) = '"&slip_month&"' "


						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							'sql = "select * from transit_cost where cancel_yn ='N' and repair_cost > 0 and cost_center ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
							objBuilder.Append "	AND cost_center ='"&sales_saupbu&"' "
						else
							'sql = "select * from transit_cost where cancel_yn ='N' and (cost_center ='������' or cost_center ='����������') and repair_cost > 0 and mg_saupbu ='"&sales_saupbu&"' and substring(run_date,1,7) = '"&slip_month&"' ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name"
							objBuilder.Append "	AND (cost_center ='������' OR cost_center ='����������') AND mg_saupbu ='"&sales_saupbu&"' "
						end If
						objBuilder.Append "ORDER BY cost_center,run_date,mg_saupbu,org_name, user_name "

						'Rs.Open Sql, Dbconn, 1
						Set rsRepair = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsRepair.EOF Then
							arrRepair = rsRepair.getRows()
						End If
						rsRepair.close() : Set rsRepair = Nothing

						If IsArray(arrRepair) Then
						  	tax_bill_yn  = "�Ϲ�"
							gubun        = "�����"
							account      = "����������"
							customer     = ""
							cost_vat     = 0

							For j = LBound(arrRepair) To UBound(arrRepair, 2)
								cost_center  = arrRepair(0, j)
								mg_saupbu    = arrRepair(1, j)
								emp_company  = arrRepair(2, j)
								bonbu        = arrRepair(3, j)
								saupbu       = arrRepair(4, j)
								team         = arrRepair(5, j)
								org_name     = arrRepair(6, j)
								reside_place = arrRepair(7, j)
								company      = arrRepair(8, j)
								emp_name     = arrRepair(9, j)
								slip_date    = arrRepair(10, j)
								slip_seq     = arrRepair(11, j)
								price        = arrRepair(12, j)
								cost         = arrRepair(12, j)
								slip_memo    = arrRepair(13, j)

								i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							Next
						End If

						'����ī��
						Dim rsCard, arrCard

						objBuilder.Append "SELECT account, cost_center, mg_saupbu, emp_company, bonbu, saupbu, team, "
						objBuilder.Append "	org_name, reside_place, reside_company, emp_name, slip_date, approve_no, "
						objBuilder.Append "	customer, price, cost, cost_vat, account_item "
						objBuilder.Append "FROM card_slip "
						objBuilder.Append "where pl_yn = 'Y' AND (card_type not like '%����%' OR com_drv_yn = 'Y') "
						objBuilder.Append "	AND SUBSTRING(slip_date, 1, 7) = '"&slip_month&"' "

						if sales_saupbu = "��������" or sales_saupbu = "�ι������" then
							'sql = "select * from card_slip where (pl_yn = 'Y') and (card_type not like '%����%' or com_drv_yn = 'Y') and cost_center ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
							objBuilder.Append "	AND cost_center ='"&sales_saupbu&"' "
						else
							'sql = "select * from card_slip where (pl_yn = 'Y') and (card_type not like '%����%' or com_drv_yn = 'Y') and (cost_center ='������' or cost_center ='����������') and mg_saupbu ='"&sales_saupbu&"' and substring(slip_date,1,7) = '"&slip_month&"' ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name"
							objBuilder.Append "	AND (cost_center ='������' OR cost_center ='����������') AND mg_saupbu ='"&sales_saupbu&"' "
						end If
						objBuilder.Append "ORDER BY cost_center,slip_date,mg_saupbu,org_name, emp_name "

						'// and mg_saupbu ='"&sales_saupbu&"' and
						'//where (pl_yn = 'Y') and (card_type not like '%����%' or com_drv_yn = 'Y') and
						'Rs.Open Sql, Dbconn, 1
						Set rsCard = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsCard.EOF Then
							arrCard = rsCard.getRows()
						End If
						rsCard.close() : Set rsCard = Nothing
						DBConn.Close() : Set DBConn = Nothing

						If IsArray(arrCard) Then
						  	tax_bill_yn   = "�Ϲ�"
							gubun         = "����ī��"

							For j = LBound(arrCard) To UBound(arrCard, 2)
								account       = arrCard(0, j)
								cost_center   = arrCard(1, j)
								mg_saupbu     = arrCard(2, j)
								emp_company   = arrCard(3, j)
								bonbu         = arrCard(4, j)
								saupbu        = arrCard(5, j)
								team          = arrCard(6, j)
								org_name      = arrCard(7, j)
								reside_place  = arrCard(8, j)
								company       = arrCard(9, j)
								emp_name      = arrCard(10, j)
								slip_date     = arrCard(11, j)
								slip_seq      = arrCard(12, j)
								customer      = arrCard(13, j)
								price         = arrCard(14, j)
								cost          = arrCard(15, j)
								cost_vat      = arrCard(16, j)
								slip_memo     = arrCard(17, j)

								i = i + 1
						%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=gubun%></td>
								<td><%=account%></td>
								<td><%=cost_center%></td>
								<td><%=mg_saupbu%></td>
								<td><%=tax_bill_yn%></td>
								<td><%=emp_company%></td>
								<td><%=bonbu%></td>
								<td><%=saupbu%></td>
								<td><%=team%></td>
								<td><%=org_name%></td>
								<td><%=reside_place%></td>
								<td><%=company%></td>
								<td><%=emp_name%></td>
								<td><%=slip_date%></td>
								<td><%=slip_seq%></td>
								<td><%=customer%></td>
							  	<td class="right"><%=formatnumber(price,0)%></td>
							  	<td class="right"><%=formatnumber(cost,0)%></td>
							  	<td class="right"><%=formatnumber(cost_vat,0)%></td>
								<td><%=slip_memo%></td>
							</tr>
						<%
							Next
						End If
						%>
						</tbody>
					</table>
				</div>
		</div>
	</div>
	</body>
</html>

