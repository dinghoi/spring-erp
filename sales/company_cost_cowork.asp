<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
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

Dim from_month, to_month, min_month, now_month, trade_name

cost_month = f_Request("cost_month")
sales_saupbu = f_Request("sales_saupbu")

from_month = f_Request("from_month")
to_month = f_Request("to_month")

trade_name = f_Request("trade_name")

'If sales_saupbu = "" Then
'	sales_saupbu = "��ü"
'End If

'����� ��ü View ����
Select Case emp_no
	Case "102592", "100359", "100001", "100740"
		view_yn = "Y"
	Case Else
		view_yn = "N"
		sales_saupbu = bonbu
End Select

'If cost_month = "" Then
'	before_date = DateAdd("m", -1, Now())
'	cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date),6,2)
'	sales_saupbu = "��ü"
'End If

'min_month = "201501"
now_month = CStr(Mid(Now(), 1, 4)) & CStr(Mid(Now(), 6, 2))

If from_month = "" Then
	from_month = now_month - 1
End If

If to_month = "" Then
	to_month = now_month
End If

'If sales_saupbu = "" Then
'	sales_saupbu = "��ü"
'End If

cost_year = Mid(to_month, 1, 4)

'from_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2) & "-01"
'end_date = DateValue(from_date)
'end_date = DateAdd("m", 1, from_date)
'to_date = CStr(DateAdd("d", -1, end_date))
'mm = Mid(cost_month, 5, 2)
'cost_year = Mid(cost_month, 1, 4)
'cost_date = Mid(cost_month, 1, 4) & "-" & Mid(cost_month, 5, 2)

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

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				var from_year = $('#from_year').val();
				var to_year = $('#to_year').val();

				if(from_year != to_year){
					alert("�˻� �⵵�� �����ؾ� �մϴ�.");
					return false;
				}

				if (document.frm.from_month.value == "") {
					alert ("���۳���� �Է��ϼ���.");
					return false;
				}
				if (document.frm.to_month.value == "") {
					alert ("�������� �Է��ϼ���.");
					return false;
				}
				return true;
			}

			function scrollAll() {
			//  document.all.leftDisplay2.scrollTop = document.all.mainDisplay2.scrollTop;
			  document.all.topLine2.scrollLeft = document.all.mainDisplay2.scrollLeft;
			}

			//���� �˻�
			function saupbuSearch(){
				console.log($('#sales_saupbu').val());

				$('#trade_name').val('');
				frmcheck();
			}
			/*
			function tradeSearch(){
				console.log($('#trade_name').val());

				frmcheck()
			}*/

			//���� ���� �ٿ�ε�
			function cowork_excel(from_date, to_date, dept){
				var url = '/sales/excel/company_cost_cowork_excel.asp';
				console.log(dept);

				location.href = url+'?from_month='+from_date+'&to_month='+to_date+'&sales_saupbu='+dept;
			}

		</script>
	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3><br/>
				<!--<h3 class="stit">1. õ���� ���� �ŷ�ó ����� ��Ÿ �׸����� ó�� </h3>-->
				<form action="/sales/company_cost_cowork.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
								<label>
								&nbsp;&nbsp;<strong>���۳��&nbsp;</strong>(��201401) :
                                	<input name="from_month" type="text" value="<%=from_month%>" style="width:70px" />
									<input type="hidden" name="from_year" value="<%=Mid(from_month, 1, 4)%>" />
								</label>
								~
								<label>
								&nbsp;&nbsp;<strong>������&nbsp;</strong>(��201501) :
                                	<input name="to_month" type="text" value="<%=to_month%>" style="width:70px" />
									<input type="hidden" name="to_year" value="<%=Mid(to_month, 1, 4)%>" />
								</label>

								<label>
									<strong>����� &nbsp;:</strong>
									<%
									Dim rsOrg, arrOrg, org_saupbu

									objBuilder.Append "SELECT saupbu "
									objBuilder.Append "FROM saupbu_sales "
									objBuilder.Append "WHERE saupbu <> '' AND SUBSTRING(sales_date, 1, 4) = '"&cost_year&"' "

									'�Ҽ� ����� ���� ó��
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
									<select name="sales_saupbu" id="sales_saupbu" style="width:150px" onchange="saupbuSearch();">
										<option value="" <%If sales_saupbu = "" then %>selected<% end if %>>��ü</option>
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

								<label>
									<strong>�ŷ�ó &nbsp;:</strong>
									<%
									Dim rsTrade, arrTrade, tradeName

									objBuilder.Append "SELECT company_name AS 'trade_name' FROM company_cost_profit "
									objBuilder.Append "WHERE (cost_month >= '"&from_month&"' AND cost_month <= '"&to_month&"') "
									objBuilder.Append "	AND saupbu = '"&sales_saupbu&"' "
									objBuilder.Append "	AND (sales_cost <> '0' OR (pay_cost + general_cost + common_cost + part_cost + manage_cost)) <> '0' "

									Set rsTrade = DBConn.Execute(objBuilder.ToString())
									objBuilder.Clear()

									If Not rsTrade.EOF Then
										arrTrade = rsTrade.getRows()
									End If
									rsTrade.Close() : Set rsTrade = Nothing
									%>
									<select name="trade_name" id="trade_name" style="width:150px;" onchange="frmcheck();">
										<option value="" <%If trade_name = "" Then %>selected<%End If %>>��ü</option>
										<%
										If IsArray(arrTrade) Then
											For i = LBound(arrTrade) To UBound(arrTrade, 2)
												tradeName = arrTrade(0, i)
										%>
										<option value='<%=tradeName%>' <%If tradeName = trade_name Then %>selected<%End If %>><%=tradeName%></option>
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
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="2%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">�����</th>
								<th scope="col" rowspan="2" style="border-right:1px solid #cbcbcb">�ŷ�ó ��</th>
								<th scope="col" colspan="2" style="border-bottom:1px solid #cbcbcb">���� ����</th>
								<th scope="col" colspan="2" style="border-bottom:1px solid #cbcbcb">���� ����</th>
								<th scope="col" rowspan="2">�� �Ǽ�</th>
								<th scope="col" rowspan="2">�� ���</th>
								<th scope="col" rowspan="2"></th>
							</tr>
							<tr>
								<th scope="col">�Ǽ�</th>
								<th scope="col">���</th>
								<th scope="col">�Ǽ�</th>
								<th scope="col">���</th>
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
							<col width="*%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="2%" >
						</colgroup>
						<tbody>
						<%
						Dim rsSalesOrg, arrSalesOrg, row_cnt
						Dim rsCowork, arrCowork
						Dim as_company, as_give_cowork, as_get_cowork, cowork_give_cost, cowork_get_cost
						Dim as_total, cost_total, give_sum, get_sum, give_cost_sum, get_cost_sum
						Dim as_sum, cost_sum, std_cost_2021

						'���� ����� ��ȸ
						objBuilder.Append "SELECT saupbu FROM sales_org "
						objBuilder.Append "WHERE sales_year = '"&cost_year&"' "

						If sales_saupbu <> "" Then
							objBuilder.Append "AND saupbu = '"&sales_saupbu&"' "
						End If

						objBuilder.Append "ORDER BY sort_seq ASC "

						Set rsSalesOrg = DBConn.Execute(objBuilder.ToString())
						objBuilder.Clear()

						If Not rsSalesOrg.EOF Then
							arrSalesOrg = rsSalesOrg.getRows()
						End If
						rsSalesOrg.Close() : Set rsSalesOrg = Nothing

						'ǥ�� �ΰǺ�
						std_cost_2021 = 30000

						If IsArray(arrSalesOrg) Then
							For i = LBound(arrSalesOrg) To UBound(arrSalesOrg, 2)
								saupbu = arrSalesOrg(0, i)

								objBuilder.Append "SELECT as_company, as_give_cowork, as_get_cowork, "
								objBuilder.Append "	cowork_give_cost, cowork_get_cost,  "
								objBuilder.Append "	as_total, "
								objBuilder.Append "	(cowork_give_cost + cowork_get_cost) AS 'cost_total' "
								objBuilder.Append "FROM ( "
								objBuilder.Append "	SELECT as_company, as_give_cowork, as_get_cowork, "
								objBuilder.Append "		(as_give_cowork * "&std_cost_2021&" * -1) AS 'cowork_give_cost', "
								objBuilder.Append "		(as_get_cowork * "&std_cost_2021&") AS 'cowork_get_cost', "
								objBuilder.Append "		(as_give_cowork + as_get_cowork) AS 'as_total' "
								objBuilder.Append "	FROM as_acpt_status AS aast "
								objBuilder.Append "	INNER JOIN trade AS trdt ON aast.as_company = trdt.trade_name "
								objBuilder.Append "		AND trdt.trade_id = '����' "
								objBuilder.Append "	WHERE (aast.as_month >= '"&from_month&"' AND aast.as_month <= '"&to_month&"') "
								objBuilder.Append "		AND (as_give_cowork > 0 OR as_get_cowork > 0) "

								If saupbu = "��Ÿ�����" Then
									objBuilder.Append "		AND trdt.saupbu = '' "
								Else
									objBuilder.Append "		AND trdt.saupbu = '"&saupbu&"' "
								End If

								If trade_name <> "" Then
									objBuilder.Append "	AND aast.as_company LIKE '%"&trade_name&"%' "
								End If
								objBuilder.Append ") r1 "

								Set rsCowork = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

								If rsCowork.EOF Or rsCowork.BOF Then
									arrCowork = ""
								Else
									arrCowork = rsCowork.getRows()
								End If
								rsCowork.Close() : Set rsCowork = Nothing

								If IsArray(arrCowork) Then
									'����Ʈ �� ����
									row_cnt = UBound(arrCowork, 2) + 1

									'����Ʈ �� �б� ó��
									For j = LBound(arrCowork) To UBound(arrCowork, 2)
										as_company = arrCowork(0, j)	'�ŷ�ó��
										as_give_cowork = CDbl(f_toString(arrCowork(1, j), 0))	'���� ���� �Ǽ�
										as_get_cowork = CDbl(f_toString(arrCowork(2, j), 0))	'���� ���� �Ǽ�
										cowork_give_cost = CDbl(f_toString(arrCowork(3, j), 0))	'���� ���� ���
										cowork_get_cost = CDbl(f_toString(arrCowork(4, j), 0))	'���� ���� ���
										as_total = CDbl(f_toString(arrCowork(5, j), 0))	'���� �� �Ǽ�
										cost_total = CDbl(f_toString(arrCowork(6, j), 0))	'���� �� ���

										'�Ѱ�
										give_sum = FormatNumber(give_sum + as_give_cowork, 0)
										get_sum = FormatNumber(get_sum + as_get_cowork, 0)
										give_cost_sum = FormatNumber(give_cost_sum + cowork_give_cost, 0)
										get_cost_sum = FormatNumber(get_cost_sum + cowork_get_cost, 0)
										as_sum = FormatNumber(as_sum + as_total, 0)
										cost_sum = FormatNumber(cost_sum + cost_total, 0)
							%>
							<tr>
							<%If j = 0 Then %>
								<td class="first" rowspan="<%=CInt(row_cnt)%>" style="background-color:#EEFFFF;font-weight:bold;"><%=saupbu%></td>
							<%End If %>
								<td style="border-left:1px solid #CBCBCB"><%=as_company%></td>
								<td class="right"><%=FormatNumber(as_give_cowork, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(cowork_give_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(as_get_cowork, 0)%>&nbsp;</td>
                                <td class="right"><%=FormatNumber(cowork_get_cost, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(as_total, 0)%>&nbsp;</td>
								<td class="right"><%=FormatNumber(cost_total, 0)%>&nbsp;</td>
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
								<td bgcolor="#FFE8E8" class="right"><%=give_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=give_cost_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=get_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=get_cost_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=as_sum%>&nbsp;</td>
								<td bgcolor="#FFE8E8" class="right"><%=cost_sum%>&nbsp;</td>
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
						<a href="#" onclick="cowork_excel('<%=from_month%>', '<%=to_month%>', '<%=sales_saupbu%>');" class="btnType04">�����ٿ�ε�</a>
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