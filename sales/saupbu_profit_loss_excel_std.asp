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
Dim sum_amt(10)
Dim tot_amt(10)
Dim detail_tab(30)
Dim cost_amt(30,10)
Dim cost_tab

Dim cost_year, cost_mm, sales_saupbu, cost_month, title_line, savefilename
Dim before_year, before_mm, before_month, c_month, b_month
Dim condi_sql, i, rsPreCostSum, curr_sales_amt, before_sales_amt
Dim rsCurrCostSum

cost_year = f_Request("cost_year")
cost_mm = f_Request("cost_mm")
sales_saupbu = f_Request("sales_saupbu")
cost_month = cstr(cost_year) & cstr(cost_mm)

title_line = cost_year & "��" & cost_mm & "�� " & sales_saupbu & " ����κ� ���� ��Ȳ"
savefilename = title_line & ".xls"

cost_tab = array("�ΰǺ�","��Ư��","�Ϲݰ��","�����","����ī��","������","���ֺ�","����","���","��ݺ�","�󰢺�")

'���� �ٿ�ε� ����
Call ViewExcelType(savefilename)

If cost_mm = "01" Then
	before_year = CStr(Int(cost_year) - 1)
	before_mm = "12"
Else
	before_year = cost_year
	before_mm = Right("0" & CStr(Int(cost_mm) - 1), 2)
End If

before_month = CStr(before_year) & CStr(before_mm)	'���� �⵵(yyyyMM)
c_month = CStr(cost_year) & "-" & CStr(cost_mm)		'��� �⵵(yyyy-MM)
b_month = CStr(before_year) & "-" & CStr(before_mm)	'���� �⵵(yyyy-MM)

'if sales_saupbu = "��ü" then
'	condi_sql = ""
'  else
'  	condi_sql = " and saupbu ='"&sales_saupbu&"'"
'end if
'if sales_saupbu = "��Ÿ�����" then
'  	condi_sql = " and (saupbu ='' or saupbu = '��Ÿ�����')"
'end if

Select Case sales_saupbu
	Case "��ü"
		condi_sql = ""
	Case "��Ÿ�����"
		condi_sql = " AND (saupbu ='' OR saupbu = '��Ÿ�����') "
	Case "����", "�����׷�"
		condi_sql = " AND saupbu IN ('����', '�����׷�') "
	Case Else
		condi_sql = " AND saupbu ='"&sales_saupbu&"' "
End Select

for i = 0 to 10
	sum_amt(i) = 0
	tot_amt(i) = 0
next

'�����(����)
'sql = "select sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&b_month&"'"&condi_sql
'Set rs_sum = Dbconn.Execute (sql)
'if isnull(rs_sum(0)) then
'	before_sales_amt = 0
'  else
'	before_sales_amt = CCur(rs_sum(0))
'end if
objBuilder.Append "SELECT SUM(cost_amt) AS sales_amt "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(SALES_DATE, 1, 7) = '"&b_month&"'"&condi_sql

Set rsPreCostSum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsPreCostSum(0)) Then
	before_sales_amt = 0
Else
	before_sales_amt = CDbl(rsPreCostSum(0))
End If

rsPreCostSum.Close() : Set rsPreCostSum = Nothing

'�����(���)
'sql = "select sum(cost_amt) as sales_amt from saupbu_sales where substring(sales_date,1,7) = '"&c_month&"'"&condi_sql
'Set rs_sum = Dbconn.Execute (sql)
'if isnull(rs_sum(0)) then
'	curr_sales_amt = 0
'  else
'	curr_sales_amt = CCur(rs_sum(0))
'end if
objBuilder.Append "SELECT SUM(cost_amt) AS sales_amt "
objBuilder.Append "FROM saupbu_sales "
objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&c_month&"'"&condi_sql

Set rsCurrCostSum = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

If IsNull(rsCurrCostSum(0)) Then
	curr_sales_amt = 0
Else
	curr_sales_amt = CDbl(rsCurrCostSum(0))
End If

rsCurrCostSum.Close() : Set rsCurrCostSum = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
	</head>
	<body>
		<div id="wrap">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
                <div class="gView">
					<table border="1" cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="*" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="7%" >
							<col width="6%" >
							<col width="7%" >
							<col width="7%" >
							<col width="8%" >
							<col width="8%" >
							<col width="6%" >
							<col width="1%" >
						</colgroup>
						<thead>
							<tr>
							  <th rowspan="2" class="first" scope="col">����׸�</th>
							  <th rowspan="2" scope="col">���γ���</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� ��&nbsp;(<%=before_year%>��<%=before_mm%>��)</th>
							  <th colspan="5" scope="col" style=" border-bottom:1px solid #e3e3e3;">�� ��&nbsp;(<%=cost_year%>��<%=cost_mm%>��)</th>
							  <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">����</th>
							  <th rowspan="2" scope="col"></th>
						  </tr>
							<tr>
							  <th scope="col" style="border-left:1px solid #e3e3e3;">����������</th>
							  <th scope="col">������</th>
							  <th scope="col">��������</th>
							  <th scope="col">�ι������</th>
							  <th scope="col">��</th>
							  <th scope="col">����������</th>
							  <th scope="col">������</th>
							  <th scope="col">��������</th>
							  <th scope="col">�ι������</th>
							  <th scope="col">��</th>
							  <th scope="col">�ݾ�</th>
							  <th scope="col">��</th>
                          </tr>
						</thead>
						<tbody>
						<tr bgcolor="#FFFFCC">
							  <td colspan="2" class="first" scope="col"><strong>�����</strong></td>
							  <td colspan="5" scope="col" class="right"><%=formatnumber(before_sales_amt,0)%></td>
							  <td colspan="5" scope="col" class="right"><%=formatnumber(curr_sales_amt,0)%></td>
						<%
						   	Dim incr_amt, incr_per

							incr_amt = curr_sales_amt - before_sales_amt

							If before_sales_amt = 0 And curr_sales_amt = 0 Then
								incr_per = 0
							ElseIf before_sales_amt = 0 Then
								incr_per = 100
							Else
								incr_per = incr_amt / before_sales_amt * 100
							End If
						%>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
					<%
					Dim jj, rec_cnt, j
					Dim rsCostDetail, rsCostSum

					for jj = 0 to 10
						rec_cnt = 0

						for i = 1 to 30
							detail_tab(i) = ""
							for j = 1 to 10
								cost_amt(i,j) = 0
								sum_amt(j) = 0
							next
						next

						If cost_tab(jj) = "�ΰǺ�" Then
							'sql = "select cost_detail from saupbu_cost_account where cost_id ='"&cost_tab(jj)&"' order by view_seq"
							'rs.Open sql, Dbconn, 1
							'do until rs.eof
							'	rec_cnt = rec_cnt + 1
							'	detail_tab(rec_cnt) = rs("cost_detail")
							'	rs.movenext()
							'loop
							'rs.close()

							objBuilder.Append "SELECT cost_detail "
							objBuilder.Append "FROM saupbu_cost_account "
							objBuilder.Append "WHERE cost_id = '�ΰǺ�' "

							'objBuilder.Append "	AND cost_detail NOT IN('��ġ����') "	'ǥ���� ��� ��ġ����,���� ����
							'objBuilder.Append "	AND cost_detail NOT IN('��ġ����', '����') "	'ǥ���� ��� ��ġ����,���� ����

							objBuilder.Append "ORDER BY view_seq "

							Set rsCostDetail = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rsCostDetail.EOF
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rsCostDetail("cost_detail")

								rsCostDetail.MoveNext()
							Loop
							rsCostDetail.Close() : Set rsCostDetail = Nothing
						Else
							'sql = "select cost_detail from saupbu_profit_loss where (cost_year ='"&cost_year&"' or cost_year ='"&before_year&"') and cost_id ='"&'cost_tab(jj)&"'"&condi_sql&" group by cost_detail order by cost_detail"
							'rs.Open sql, Dbconn, 1
							'do until rs.eof
							'	rec_cnt = rec_cnt + 1
							'	detail_tab(rec_cnt) = rs("cost_detail")
							'	rs.movenext()
							'loop
							'rs.close()
							objBuilder.Append "SELECT cost_detail "
							objBuilder.Append "FROM saupbu_profit_loss "
							objBuilder.Append "WHERE (cost_year = '"& cost_year &"' OR cost_year = '"& before_year &"') "
							objBuilder.Append "	AND cost_id ='"& cost_tab(jj) &"'"& condi_sql
							objBuilder.Append "GROUP BY cost_detail "
							objBuilder.Append "ORDER BY cost_detail "

							Set rsCostDetail = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rsCostDetail.EOF
								rec_cnt = rec_cnt + 1
								detail_tab(rec_cnt) = rsCostDetail("cost_detail")

								rsCostDetail.MoveNext()
							Loop
							rsCostDetail.Close() : Set rsCostDetail = Nothing
						End If

						If rec_cnt <> 0 Then
							' ���� �ݾ� SUM
							'sql = "SELECT cost_center, cost_detail, SUM(cost_amt_"& before_mm &") AS cost "
							'sql = sql & "FROM saupbu_profit_loss "
							'sql = sql &  "WHERE cost_year = '"& before_year &"' "
							'sql = sql &  "	AND cost_id = '"& cost_tab(jj) &"'"&condi_sql
							'sql = sql &  "	AND cost_center NOT IN ('�ι������', '��������') "
							'sql = sql &  "GROUP BY cost_center, cost_detail "
							'sql = sql &  "ORDER BY cost_center, cost_detail "
							'rs.Open sql, Dbconn, 1
							'do until rs.eof
							'	for i = 1 to 30
							'		if rs("cost_detail") = detail_tab(i) then
							'			select case rs("cost_center")
							'				case "����������"
							'					j = 1
							'				case "������"
							'					j = 2
							'				case "��������"
							'					j = 3
							'				case "�ι������"
							'					j = 4
							'			end select
							'			cost_amt(i,j) = cost_amt(i,j) + ccur(rs("cost"))
							'			cost_amt(i,5) = cost_amt(i,5) + ccur(rs("cost"))
							'			sum_amt(j) = sum_amt(j) + ccur(rs("cost"))
							'			sum_amt(5) = sum_amt(5) + ccur(rs("cost"))
							'			tot_amt(j) = tot_amt(j) + ccur(rs("cost"))
							'			tot_amt(5) = tot_amt(5) + ccur(rs("cost"))
							'			exit for
							'		end if
							'	next
							'	rs.movenext()
							'loop
							'rs.close()
							objBuilder.Append "SELECT cost_center, cost_detail, SUM(cost_amt_"& before_mm &") AS cost "
							objBuilder.Append "FROM saupbu_profit_loss "
							objBuilder.Append "WHERE cost_year = '"& before_year &"' "
							objBuilder.Append "	AND cost_id = '"& cost_tab(jj) &"'"&condi_sql
							objBuilder.Append "	AND cost_center NOT IN ('�ι������', '��������') "
							objBuilder.Append "GROUP BY cost_center, cost_detail "
							objBuilder.Append "ORDER BY cost_center, cost_detail "

							Set rsPreCostSum = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rsPreCostSum.EOF
								For i = 1 To 30
									' �������� ������ detail_tab�� ���ٸ� cost_detail�� ������ �ʴ´�..
									If rsPreCostSum("cost_detail") = detail_tab(i) Then
										Select Case rsPreCostSum("cost_center")
											Case "����������" : j = 1
											Case "������"     : j = 2
											Case "��������" : j = 3
											Case "�ι������" : j = 4
										End Select

										cost_amt(i, j) = cost_amt(i, j) + CDbl(rsPreCostSum("cost"))
										cost_amt(i, 5) = cost_amt(i, 5) + CDbl(rsPreCostSum("cost"))
										sum_amt(j) = sum_amt(j) + CDbl(rsPreCostSum("cost"))
										sum_amt(5) = sum_amt(5) + CDbl(rsPreCostSum("cost"))
										tot_amt(j) = tot_amt(j) + CDbl(rsPreCostSum("cost"))
										tot_amt(5) = tot_amt(5) + CDbl(rsPreCostSum("cost"))

										Exit For
									End If
								Next

								rsPreCostSum.MoveNext()
							Loop
							rsPreCostSum.Close() : Set rsPreCostSum = Nothing

							' ��� �ݾ� SUM
							'sql = "SELECT cost_center, cost_detail, SUM(cost_amt_"&cost_mm&") AS cost "
							'sql = sql & "FROM  SAUPBU_PROFIT_LOSS "
							'sql = sql & "WHERE  cost_year ='"& cost_year &"' "
							'sql = sql & "	AND cost_id ='"& cost_tab(jj) &"' "&condi_sql
							'sql = sql & "	AND cost_center NOT IN ('�ι������', '��������') "
							'sql = sql & "GROUP  BY cost_center, cost_detail "
							'sql = sql & "ORDER  BY cost_center, cost_detail "
							'rs.Open sql, Dbconn, 1
							'do until rs.eof
							'	for i = 1 to 30
							'		if rs("cost_detail") = detail_tab(i) then
							'			select case rs("cost_center")
							'				case "����������"
							'					j = 6
							'				case "������"
							'					j = 7
							'				case "��������"
							'					j = 8
							'				case "�ι������"
							'					j = 9
							'			end select
							'			cost_amt(i,j) = cost_amt(i,j) + ccur(rs("cost"))
							'			cost_amt(i,10) = cost_amt(i,10) + ccur(rs("cost"))
							'			sum_amt(j) = sum_amt(j) + ccur(rs("cost"))
							'			sum_amt(10) = sum_amt(10) + ccur(rs("cost"))
							'			tot_amt(j) = tot_amt(j) + ccur(rs("cost"))
							'			tot_amt(10) = tot_amt(10) + ccur(rs("cost"))
							'			exit for
							'		end if
							'	next
							'	rs.movenext()
							'loop
							'rs.close()
							objBuilder.Append "SELECT cost_center, cost_detail, SUM(cost_amt_"&cost_mm&") AS cost "
							objBuilder.Append "FROM  SAUPBU_PROFIT_LOSS "
							objBuilder.Append "WHERE  cost_year ='"& cost_year &"' "
							objBuilder.Append "	AND cost_id ='"& cost_tab(jj) &"' "&condi_sql
							objBuilder.Append "	AND cost_center NOT IN ('�ι������', '��������') "
							objBuilder.Append "GROUP  BY cost_center, cost_detail "
							objBuilder.Append "ORDER  BY cost_center, cost_detail "

							Set rsCurrCostSum = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()

							Do Until rsCurrCostSum.EOF
								For i = 1 To 30
									' �������� ������ detail_tab�� ���ٸ� cost_detail�� ������ �ʴ´�..
									If rsCurrCostSum("cost_detail") = detail_tab(i) Then
										Select Case rsCurrCostSum("cost_center")
											Case "����������"	: j = 6
											Case "������"	    : j = 7
											Case "��������"	: j = 8
											Case "�ι������"	: j = 9
										End Select

										cost_amt(i, j) = cost_amt(i, j) + CDbl(rsCurrCostSum("cost"))
										cost_amt(i, 10) = cost_amt(i, 10) + CDbl(rsCurrCostSum("cost"))
										sum_amt(j) = sum_amt(j) + CDbl(rsCurrCostSum("cost"))
										sum_amt(10) = sum_amt(10) + CDbl(rsCurrCostSum("cost"))
										tot_amt(j) = tot_amt(j) + CDbl(rsCurrCostSum("cost"))
										tot_amt(10) = tot_amt(10) + CDbl(rsCurrCostSum("cost"))

										Exit For
									End If
								Next

								rsCurrCostSum.MoveNext()
							Loop

							rsCurrCostSum.Close() : Set rsCurrCostSum = Nothing
						%>
							<tr>
							  	<td rowspan="<%=rec_cnt + 1%>" class="first">
							<% if jj = 2 or jj = 3 then	%>
                        	  	<%=cost_tab(jj)%><br>(���ݻ��)
							<%   else	%>
                        	  	<%=cost_tab(jj)%>
                        	<% end if	%>
                              	</td>
								<td class="left"><%=detail_tab(1)%></td>

							<% for j = 1 to 10	%>
								<td class="right"><%=formatnumber(cost_amt(1,j),0)%></td>
							<% next	%>
							<%
						   	incr_amt = cost_amt(1,10) - cost_amt(1,5)
						   	if cost_amt(1,5) = 0 and cost_amt(1,10) = 0 then
						   		incr_per = 0
							  elseif cost_amt(1,5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / cost_amt(1,5) * 100
						   	end if
							%>
								<td class="right"><%=formatnumber(incr_amt,0)%></td>
				        		<td class="right"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right">&nbsp;</td>
							</tr>
							<% for i = 2 to rec_cnt	%>
                        	<tr>
								<td class="left" style=" border-left:1px solid #e3e3e3;"><%=detail_tab(i)%></td>
							<%   for j = 1 to 10	%>
								<td class="right"><%=formatnumber(cost_amt(i,j),0)%></td>
							<%   next	%>
							<%
						    incr_amt = cost_amt(i,10) - cost_amt(i,5)
						    if cost_amt(i,5) = 0 and cost_amt(i,10) = 0 then
						   		incr_per = 0
							  elseif cost_amt(i,5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / cost_amt(i,5) * 100
						    end if
							%>
					     		<td class="right"><%=formatnumber(incr_amt,0)%></td>
								<td class="right"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right">&nbsp;</td>
							</tr>
							<% next	%>
							<tr>
							  <td class="left" style=" border-left:1px solid #e3e3e3;" bgcolor="#EEFFFF">�Ұ�</td>
							<% for j = 1 to 10	%>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(sum_amt(j),0)%></td>
							<% next	%>
							<%
						   	incr_amt = sum_amt(10) - sum_amt(5)
						   	if sum_amt(5) = 0 and sum_amt(10) = 0 then
						   		incr_per = 0
							  elseif sum_amt(5) = 0 then
								incr_per = 100
							  else
						   		incr_per = incr_amt / sum_amt(5) * 100
						   	end if
							%>
					      		<td class="right" bgcolor="#EEFFFF"><%=formatnumber(incr_amt,0)%></td>
								<td class="right" bgcolor="#EEFFFF"><%=formatnumber(incr_per,2)%>%</td>
								<td class="right" bgcolor="#EEFFFF">&nbsp;</td>
							</tr>
					<%
						end if
					next

					%>
						<tr bgcolor="#FFFFCC">
							<td colspan="2" class="first" scope="col"><strong>����հ�</strong></td>
						<% 'for j = 1 to 10	%>
							<!--<td scope="col" class="right"><%'=formatnumber(tot_amt(j),0)%></td>-->
						<% 'next	%>
						<%
							Dim arrManage, arrManageCost, arrComm, arrCommCost
							Dim kk, manage_cost, comm_cost
							Dim tot_amt_before, tot_amt_curr

							'�ι� ��� ����
							arrManage = Array("SI1����", "SI2����", "NI����", "��������")
							arrManageCost = Array("115500000", "50200000", "35300000", "400000")

							'���� ��� ����
							arrComm = Array("SI1����", "SI2����", "��������", "NI����", "ICT����", "����SI����", "����SI����", "����Ʈ����", "DI����ι�")
							arrCommCost = Array("78000000", "83000000", "22000000", "30000000", "19000000", "15000000", "17000000", "5000000", "5000000")

							For j = 1 To 10
								If j = 5 Or j = 10 Then
						%>
							<td class="right" alt="����հ� > ��">
								<strong>
								<%'=formatnumber(tot_amt(j),0)%>
								<%
									If j = 5 Then
										tot_amt_before = tot_amt(j) + manage_cost + comm_cost
										Response.write FormatNumber(tot_amt_before, 0)
									Else
										tot_amt_curr = tot_amt(j) + manage_cost + comm_cost
										Response.write FormatNumber(tot_amt_curr, 0)
									End If

								%>
								</strong>
							</td>
							<%
								Else
							%>
							<td class="right">
							<%'=formatnumber(tot_amt(j),0)%>
							<%
							If cost_month = "202101" Then
								Select Case j
									Case 3, 4 : Response.Write 0
									Case 8 :
										For kk = 0 To 7
											If arrComm(kk) = sales_saupbu Then
												comm_cost = arrCommCost(kk)
											End If
										Next
										Response.Write FormatNumber(comm_cost, 0)
									Case 9 :
										For kk = 0 To 3
											If arrManage(kk) = sales_saupbu Then
												manage_cost = arrManageCost(kk)
											End If
										Next
										Response.Write FormatNumber(manage_cost, 0)
									Case Else
										Response.write FormatNumber(tot_amt(j), 0)
								End Select
							Else
								If j = 3 Or j = 8 Then	'����
									For kk = 0 To 7
										If arrComm(kk) = sales_saupbu Then
											comm_cost = arrCommCost(kk)
										End If
									Next
									Response.Write FormatNumber(comm_cost, 0)
								ElseIf j = 4 Or j = 9 Then	'�ι�
									For kk = 0 To 3
										If arrManage(kk) = sales_saupbu Then
											manage_cost = arrManageCost(kk)
										End If
									Next
									Response.Write FormatNumber(manage_cost, 0)
								Else	'����, ����
									Response.write FormatNumber(tot_amt(j), 0)
								End If
							End If
							%>
							</td>
							<%
								End If
							Next

						   'incr_amt = tot_amt(10) - tot_amt(5)
						   'if tot_amt(5) = 0 and tot_amt(10) = 0 then
							'	incr_per = 0
							'  elseif tot_amt(5) = 0 then
							'	incr_per = 100
							'  else
							'	incr_per = incr_amt / tot_amt(5) * 100
						   'end If

						    incr_amt = tot_amt_curr - tot_amt_before

							If tot_amt_before = 0 And tot_amt_curr = 0 Then
								incr_per = 0
							ElseIf tot_amt_before = 0 Then
								incr_per = 100
							Else
								incr_per = incr_amt / tot_amt_before * 100
							End If
						%>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                        </tr>
						<tr bgcolor="#FFDFDF">
							  <td colspan="2" class="first" scope="col"><strong>����</strong></td>
						<%
							Dim be_profit_loss, curr_profit_loss
						   	'be_profit_loss = before_sales_amt - tot_amt(5)
						   	'curr_profit_loss = curr_sales_amt - tot_amt(10)

							be_profit_loss = before_sales_amt - tot_amt_before
							curr_profit_loss = curr_sales_amt - tot_amt_curr

						   	incr_amt = curr_profit_loss - be_profit_loss

						   	if be_profit_loss = 0 and curr_profit_loss = 0 then
						   		incr_per = 0
							elseif be_profit_loss = 0 then
								incr_per = 100
							else
						   		incr_per = incr_amt / be_profit_loss * 100
						   	end If

							if be_profit_loss < 0 then
								incr_per = incr_per * -1
							end if
						%>
							  <td scope="col" colspan="5" class="right"><%=formatnumber(be_profit_loss,0)%></td>
							  <td scope="col" colspan="5" class="right"><%=formatnumber(curr_profit_loss,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_amt,0)%></td>
							  <td scope="col" class="right"><%=formatnumber(incr_per,2)%>%</td>
 							  <td scope="col" class="right">&nbsp;</td>
                         </tr>
						</tbody>
					</table>
		</div>
	</div>
	</body>
</html>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
DBConn.Close : Set DBConn = Nothing
%>