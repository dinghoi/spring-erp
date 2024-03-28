<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'on Error resume next
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
Dim cost_month, sales_saupbu, before_date
Dim condi_sql, mm, cost_year
Dim from_date, end_date, to_date
Dim rsCompCost, arrCompCost
Dim title_line, i, j
Dim view_yn, cost_date

cost_month = f_Request("cost_month")
sales_saupbu = f_Request("sales_saupbu")

If sales_saupbu = "" Then
	sales_saupbu = "��ü"
End If

'����� ��ü View ����
Select Case emp_no
	Case "102592", "100359"
		view_yn = "Y"
	Case Else
		view_yn = "N"
		sales_saupbu = bonbu
End Select

If cost_month = "" Then
	before_date = DateAdd("m", -1, Now())
	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date),6,2)
	sales_saupbu = "��ü"
End If

from_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))
mm = Mid(cost_month, 5, 2)
cost_year = Mid(cost_month, 1, 4)
cost_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2)

title_line = "�ŷ�ó�� ������Ȳ"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
		<!-- <link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" /> -->
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.cost_month.value == "") {
					alert ("�߻������ �Է��ϼ���.");
					return false;
				}
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}
		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<h3 class="stit">1. õ���� ���� �ŷ�ó ����� ��Ÿ �׸����� ó�� </h3>
				<form action="/sales/company_cost_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>�߻����&nbsp;</strong>(��201401) :
                                	<input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
								</label>

								<label>
								<strong>����� &nbsp;:</strong>
                                <%
								Dim rsOrg, arrOrg, org_saupbu

								objBuilder.Append "SELECT saupbu "
								objBuilder.Append "FROM saupbu_sales "
								objBuilder.Append "WHERE saupbu <> '' AND SUBSTRING(sales_date, 1, 4) = '"&cost_year&"' "

								If view_yn = "N" Then
									objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
								End If

								objBuilder.Append "GROUP BY saupbu "
								objBuilder.Append "ORDER BY saupbu ASC "

								Set rsOrg = DBConn.Execute(objBuilder.ToString())

								If Not rsOrg.EOF Then
									arrOrg = rsOrg.getRows()
								End If
								objBuilder.Clear()
                                rsOrg.Close() : Set rsOrg = Nothing
                                %>
                                <select name="sales_saupbu" id="sales_saupbu" style="width:150px">
                                    <option value="��ü" <%If sales_saupbu = "��ü" then %>selected<% end if %>>��ü</option>
                                    <%
                                    If IsArray(arrOrg) Then
										For i = LBound(arrOrg) To UBound(arrOrg, 2)
											org_saupbu = arrOrg(0, i)
                                    %>
                                        <option value='<%=org_saupbu%>' <%If org_saupbu = sales_saupbu  then %>selected<% end if %>><%=org_saupbu%></option>
                                    <%
                                        Next
                                    End If
                                    %>
                                </select>
								</label>
								<img src="/image/but_ser.jpg" onclick="frmcheck();" style="cursor:pointer;" alt="�˻�">
                            </p>
						</dd>
					</dl>
				</fieldset>
					<table cellpadding="0" cellspacing="0" width="100%">
					<tr>
                    	<td>
      			<DIV id="topLine2" style="width:1200px;overflow:hidden;">
				<div class="gView">
						<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="8%" >
							<col width="12%" >
							<col width="12%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">�����</th>
								<th scope="col">�ŷ�ó ��</th>
								<th scope="col">����</th>
								<th scope="col">����������(�ΰǺ�)</th>
								<th scope="col">����������(�Ϲݰ��)</th>
								<th scope="col">����ΰ����</th>
								<th scope="col">�ι������</th>
								<th scope="col">��������</th>
								<th scope="col">NKP ����</th>
								<th scope="col"></th>
							</tr>
						</thead>
						</table>
                        </DIV>
						</td>
                    </tr>
					<tr>
                    	<td valign="top">
				        <DIV id="mainDisplay2" style="width:1200;height:400px;overflow:scroll" onscroll="scrollAll()">
						<table cellpadding="0" cellspacing="0" class="scrollList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="8%" >
							<col width="12%" >
							<col width="12%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="2%" >
						</colgroup>
						<tbody>
						<%
						Dim rsSalesOrg, arrSalesOrg, row_cnt
						Dim company, sales_cost, pay_cost, general_cost
						Dim rsComm, comm_cost, direct_cost, rsSalesTot, sales_total
						Dim sales_per, common_cost, common_total, profit_cost
						Dim sales_sum, pay_sum, general_sum, common_sum, profit_sum
						Dim rsSalesCost, company_cost

						Dim rsManage, manage_tot, manage_cost, manage_sum
						Dim rsPart, part_tot_cost, as_tot_cnt, as_cnt, as_saupbu_cnt
						Dim rsSaupbuPart, part_cnt, part_tot
						Dim part_cost, part_sum, comm_per

						Dim rsCompanyTot, company_tot

						sales_sum = 0
						pay_sum = 0
						general_sum = 0
						common_sum = 0
						part_sum = 0
						manage_sum = 0
						profit_sum = 0

						'���� ����� ��ȸ
						objBuilder.Append "SELECT saupbu FROM sales_org "
						objBuilder.Append "WHERE sales_year = '"&cost_year&"' "

						If sales_saupbu <> "��ü" Then
							objBuilder.Append "AND saupbu = '"&sales_saupbu&"' "
						End If

						objBuilder.Append "ORDER BY sort_seq ASC "

						Set rsSalesOrg = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsSalesOrg.EOF Then
							arrSalesOrg = rsSalesOrg.getRows()
						End If
						rsSalesOrg.Close() : Set rsSalesOrg = Nothing

						If IsArray(arrSalesOrg) Then
							For i = LBound(arrSalesOrg) To UBound(arrSalesOrg, 2)
								saupbu = arrSalesOrg(0, i)
								'company_cnt = arrSalesOrg(1, i)

								'����κ� ���� ��ȸ
								objBuilder.Append "SELECT SUM(cost_amt) AS 'sales_total' "
								objBuilder.Append "FROM saupbu_sales "
								objBuilder.Append "WHERE SUBSTRING(sales_date, 1, 7) = '"&cost_date&"' "
								objBuilder.Append "	AND saupbu = '"&saupbu&"'; "

								Set rsSalesTot = DBConn.Execute(objBuilder.ToString())

								sales_total = CDbl(f_toString(rsSalesTot(0), 0))	'����� �� �� ����

								objBuilder.Clear()
								rsSalesTot.Close() : Set rsSalesTot = Nothing

								'���������� Total(���� ����)
								objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS 'company_total' "
								objBuilder.Append "FROM company_cost "
								objBuilder.Append "WHERE cost_year = '"&cost_year&"' AND cost_center = '����������' "

								If saupbu <> "��Ÿ�����" Then
									objBuilder.Append "	AND (company <> '' AND company IS NOT NULL AND company <> '����') "
									objBuilder.Append " AND saupbu = '"&saupbu&"'; "
								Else
									objBuilder.Append "	AND saupbu = '' "
								End If

								Set rsCompanyTot = DBConn.Execute(objBuilder.ToString())

								company_tot = CDbl(rsCompanyTot(0))	'����� �� ����������(���� ����)

								objBuilder.Clear()
								rsCompanyTot.Close() : Set rsCompanyTot = Nothing

								'������(������ + ����������(����))
								objBuilder.Append "SELECT SUM(cost_amt_"&mm&") AS 'comm_cost', "
								objBuilder.Append "	(SELECT SUM(cost_amt_"&mm&") FROM company_cost  "
								objBuilder.Append "	WHERE cost_year = '"&cost_year&"' AND cost_center = '������'  "

								If saupbu = "��Ÿ�����" Then
									objBuilder.Append "		AND (saupbu = '' OR saupbu = '"&saupbu&"')) AS 'direct_cost' "
								Else
									objBuilder.Append "		AND saupbu = '"&saupbu&"') AS 'direct_cost' "
								End If

								objBuilder.Append "FROM company_cost "
								objBuilder.Append "WHERE cost_year = '"&cost_year&"' "
								objBuilder.Append "	AND (company = '' OR company is null OR company = '����') "
								objBuilder.Append "	AND cost_center = '����������' "
								objBuilder.Append "	AND saupbu = '"&saupbu&"' "

								Set rsComm = DBConn.Execute(objBuilder.ToString())

								comm_cost = CDbl(f_toString(rsComm("comm_cost"), 0))	'����������(����)
								direct_cost = CDbl(f_toString(rsComm("direct_cost"), 0))	'������

								'������ = ����������(����) + ������(�ΰǺ�+���)
								common_total = comm_cost + direct_cost

								objBuilder.Clear()
								rsComm.Close() : Set rsComm = Nothing

								'��������
								objBuilder.Append "SELECT ROUND((tot_cost_amt * 0.5 / tot_person * saupbu_person) "
								objBuilder.Append "	+ (tot_cost_amt * 0.5 / tot_sale * saupbu_sale), 1) AS tot_amt "
								objBuilder.Append "FROM ( "
								objBuilder.Append "	SELECT mgct.saupbu, mgct.tot_cost_amt, mgct.saupbu_person, mgct.tot_person, "

								objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
								objBuilder.Append "		FROM saupbu_sales "
								objBuilder.Append "		WHERE SUBSTRING(sales_date, 1, 7) = '"&MID(from_date, 1, 7)&"' "
								objBuilder.Append "			AND mgct.saupbu = saupbu) AS saupbu_sale, "

								objBuilder.Append "		(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
								objBuilder.Append "		FROM saupbu_sales "
								objBuilder.Append "		WHERE SUBSTRING(sales_date, 1, 7) = '"&MID(from_date, 1, 7)&"' "
								objBuilder.Append "			AND saupbu <> '��Ÿ�����') AS tot_sale "

								objBuilder.Append "	FROM management_cost AS mgct "
								objBuilder.Append "	WHERE cost_month = '"&Replace(MID(from_date, 1, 7), "-", "")&"' "
								objBuilder.Append "		AND saupbu = '"&saupbu&"' "
								objBuilder.Append "	GROUP BY saupbu "
								objBuilder.Append ") r1 "

								Set rsManage = DBConn.Execute(objBuilder.ToString())

								If rsManage.EOF Or rsManage.BOF Then
									manage_tot = 0
								Else
									manage_tot = CDbl(f_toString(rsManage(0), 0))	'�μ��� ��������
								End If

								objBuilder.Clear()
								rsManage.Close() : Set rsManage = Nothing

								'�ι������(���)
								objBuilder.Append "SELECT (SUM(cost_amt_"&mm&") - "
								objBuilder.Append "(SELECT SUM(cost_amt_"&mm&") FROM company_cost WHERE cost_year ='"&cost_year&"' "
								objBuilder.Append "	AND cost_detail = '��ġ����')) AS 'part_tot_cost', "
								objBuilder.Append "(SELECT SUM(as_total - as_set) FROM as_acpt_status WHERE as_month = '"&cost_year&mm&"') AS 'as_tot_cnt' "
								objBuilder.Append "FROM company_cost WHERE cost_year = '"&cost_year&"' AND cost_center = '�ι������' "

								Set rsPart = DBConn.Execute(objBuilder.ToString())

								part_tot_cost = CDbl(f_toString(rsPart("part_tot_cost"), 0))	'�ι������(���)
								as_tot_cnt = CInt(f_toString(rsPart("as_tot_cnt"), 0))	'AS �� �Ǽ�

								objBuilder.Clear()
								rsPart.Close() : Set rsPart = Nothing

								'����� �� AS �� �Ǽ� ��ȸ
								objBuilder.Append "SELECT SUM(as_total - as_set) AS as_cnt "
								objBuilder.Append "FROM as_acpt_status AS aast "
								objBuilder.Append "INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
								objBuilder.Append "	AND trdt.trade_id = '����' "
								objBuilder.Append "WHERE as_month = '"&cost_year&mm&"' "
								If saupbu = "��Ÿ�����" Then
									objBuilder.Append "AND trdt.saupbu = '' "
								Else
									objBuilder.Append "	AND trdt.saupbu = '"&saupbu&"' "
								End If

								Set rsSaupbuPart = DBConn.Execute(objBuilder.ToString())

								part_cnt = CInt(f_toString(rsSaupbuPart(0), 0))	'����� AS �� �Ǽ�

								objBuilder.Clear()
								rsSaupbuPart.Close() : Set rsSaupbuPart = Nothing

								'����κ� ��� �κа����
								If part_cnt > 0 Then
									part_tot = part_tot_cost / as_tot_cnt * part_cnt
								Else
									part_tot = 0
								End If



								'�ŷ�ó�� ��� ��Ȳ
								objBuilder.Append "CALL USP_SALES_COMPANY_PROFIT_SEL('"&saupbu&"', '"&cost_year&"', '"&MID(from_date, 1, 7)&"', '"&mm&"');"

								Set rsCompCost = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If Not rsCompCost.EOF Then
									arrCompCost = rsCompCost.getRows()
								End If
								rsCompCost.Close() : Set rsCompCost = Nothing

								If IsArray(arrCompCost) Then
									'����Ʈ �� ����
									row_cnt = UBound(arrCompCost, 2) + 1

									'����Ʈ �� �б� ó��
									For j = LBound(arrCompCost) To UBound(arrCompCost, 2)
										company = arrCompCost(0, j)	'�ŷ�ó��
										sales_cost = CDbl(f_toString(arrCompCost(1, j), 0))	'�ŷ�ó�� ����
										company_cost = CDbl(f_toString(arrCompCost(2, j), 0))	'����������(�ΰǺ�+�Ϲݰ��)
										pay_cost = CDbl(f_toString(arrCompCost(3, j), 0))	'����������(�ΰǺ�)
										general_cost = CDbl(f_toString(arrCompCost(4, j), 0))	'����������(�Ϲݰ��)
										as_cnt = CInt(f_toString(arrCompCost(5, j), 0))	'����Ʈ�� AS �Ǽ�

										'����ΰ����� = �ŷ�ó�� ���������� / ����� �� ����������(���� ����) * ������
										If company_tot > 0 Then
											'common_cost = company_cost / company_tot *  common_total	'����ΰ�����

											common_cost = company_cost / company_tot * common_total
										Else
											common_cost = 0
										End If

										If as_cnt > 0 Then
											part_cost = part_tot / part_cnt * as_cnt	'����Ʈ�� �κа����(AS�Ǽ� ����)
										Else
											part_cost = 0
										End If

										manage_cost = sales_cost * manage_tot / sales_total	'����Ʈ�� ��������(���� ����)
										profit_cost = sales_cost - (pay_cost + general_cost + common_cost + part_cost + manage_cost)

										'�Ѱ�
										sales_sum = FormatNumber(sales_sum + sales_cost, 0)
										pay_sum = FormatNumber(pay_sum + pay_cost, 0)
										general_sum = FormatNumber(general_sum + general_cost, 0)
										'per_sum = FormatNumber(per_sum + sales_per, 0)
										common_sum = FormatNumber(common_sum + common_cost, 0)
										part_sum = FormatNumber(part_sum + part_cost, 0)
										manage_sum = FormatNumber(manage_sum + manage_cost, 0)

										profit_sum = FormatNumber(profit_sum + profit_cost, 0)
							%>
							<tr <%If company = "��Ÿ" Then %>bgcolor="#FFFFCC"<%End If %>>
							<%If j = 0 Then %>
								<td class="first" rowspan="<%=CInt(row_cnt)%>" style="background-color:#EEFFFF;font-weight:bold;"><%=saupbu%></td>
							<%End If %>
								<td><%=company%></td>
								<td class="right"><%=FormatNumber(sales_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(pay_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(general_cost, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(common_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(part_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(manage_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(profit_cost, 0)%>&nbsp;</td>
								<td>&nbsp;</td>
							</tr>
						<%
									Next
								End If
							Next
						End If

						DBConn.Close() : Set DBConn = Nothing
						%>
							<tr>
								<td colspan="2" bgcolor="#FFE8E8" class="first" style="font-weight:bold;">�Ѱ�</td>
								<td bgcolor="#FFE8E8" class="right"><%=sales_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=pay_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=general_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=common_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=part_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=manage_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=profit_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8">&nbsp;</td>
							</tr>
						</tbody>
						</table>
                        </DIV>
						</td>
                    </tr>
				</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
					<div class="btnCenter">
						<a href="/sales/part_cost_excel.asp?cost_month=<%=cost_month%>&sales_saupbu=<%=sales_saupbu%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				</table>
			</form>
			<br>
		</div>
	</div>
	</body>
</html>
