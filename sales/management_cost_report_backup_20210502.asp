<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
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
Dim saupbu_tab(10, 2)

Dim i,ck_sw, cost_month, before_date,cost_year, cost_mm
Dim prosCost, privCost, title_line

Dim rsComm

For i = 1 To 10
    saupbu_tab(i,1) = ""
    saupbu_tab(i,2) = 0
Next 

ck_sw = Request("ck_sw")

If ck_sw = "y" Then
    cost_month = Request("cost_month")
    saupbu = Request("saupbu")
Else
    cost_month = Request.form("cost_month")
    saupbu = Request.form("saupbu")
End if

If cost_month = "" Then 
    before_date = DateAdd("m", -1, Now())
    cost_month = Mid(CStr(before_date), 1, 4) & Mid(CStr(before_date), 6, 2)
End If

cost_year = Mid(cost_month, 1, 4)
cost_mm = Mid(cost_month, 5)

'�ش� �⵵ �� ���� ��� ����(����ȣ_20201208)
Select Case Left(cost_month, 4)
	Case "2020"
		prosCost = "0.01179"	'�ش� �⵵ ���� ����
		privCost = "125000"	'�ش� �⵵ �� 1�δ� ���
	Case "2021"
		prosCost = "0.015696"
		privCost = "168269"
	Case Else	'2019�� ���� ���Ǵ� ���� ��(���� �⵵���� �ش簪�� ����)
		prosCost = "0.01388"	'�ش� �⵵ ���� ���� / 100���� ����
		privCost = "133200"	'�ش� �⵵ �� 1�δ� ���
End Select

'sql = "    SELECT a.saupbu          /* ����� �� */    " & chr(13) &_
'        "         , a.saupbu_person   /* ����� �η� */  " & chr(13) &_
'        "         , a.tot_person      /* ���η� */       " & chr(13) &_
'        "         , a.saupbu_per      /* ������ */       " & chr(13) &_
'        "         , a.saupbu_person * "&privCost&" as saupbu_cost_amt /* ��������1 */  " & chr(13) &_

'        "         , (SELECT IFNULL(sum(b.cost_amt), 0) as sales_amt from saupbu_sales b where replace(substring(b.sales_date,1,7),'-','') = '"&cost_month&"' AND a.saupbu = b.saupbu) AS saupbu_sale  " & chr(13) &_

'        "         , IFNULL(a.tot_sale, 0) AS tot_sale        /* �� ���� */      " & chr(13) &_

'        "         ,(SELECT IFNULL(sum(b.cost_amt), 0) as sales_amt from saupbu_sales b where replace(substring(b.sales_date,1,7),'-','') = '"&cost_month&"' AND a.saupbu = b.saupbu)/(SELECT IFNULL(sum(b.cost_amt), 0) as sales_amt FROM saupbu_sales b WHERE replace(substring(b.sales_date, 1, 7), '-', '') = '"&cost_month&"' AND saupbu <> 'ȸ�簣�ŷ�') AS sale_per  " & chr(13) &_
'        "         , (SELECT IFNULL(sum(b.cost_amt), 0) as sales_amt from saupbu_sales b where replace(substring(b.sales_date,1,7),'-','') = '"&cost_month&"' AND a.saupbu = b.saupbu) * "&prosCost&" as saupbu_sale_amt /* ��������2 */  " & chr(13) &_
 '       "         , a.tot_cost_amt                       " & chr(13) &_
 '       "         , (a.saupbu_person * "&privCost&")+ ((SELECT IFNULL(sum(b.cost_amt), 0) as sales_amt from saupbu_sales b where replace(substring(b.sales_date,1,7),'-','') = '"&cost_month&"' AND a.saupbu = b.saupbu) * "&prosCost&") as all_tot_cost_amt                       " & chr(13) &_
 '       "      FROM management_cost a                   " & chr(13) &_
 '       "     WHERE a.cost_month ='"&cost_month&"'       " & chr(13) &_
 '       "     AND a.saupbu <> '��Ÿ�����' " & chr(13) &_
 '       "  GROUP BY a.saupbu                             " & chr(13) &_
 '       "  ORDER BY a.saupbu                             "

objBuilder.Append "SELECT mgct.saupbu /* ����� �� */, "
objBuilder.Append "	mgct.saupbu_person /* ����� �η� */, "
objBuilder.Append "	mgct.tot_person /* ���η� */, "
objBuilder.Append "	mgct.saupbu_per /* ������ */, "
objBuilder.Append "	mgct.saupbu_person * "&privCost&" AS saupbu_cost_amt /* ��������1 */, "
objBuilder.Append "	(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "	FROM saupbu_sales "
objBuilder.Append "	WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&cost_month&"' "
objBuilder.Append "		AND mgct.saupbu = saupbu) AS saupbu_sale, "
objBuilder.Append "	IFNULL(mgct.tot_sale, 0) AS tot_sale /* �� ���� */, "
objBuilder.Append "	(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "	FROM saupbu_sales "
objBuilder.Append "	WHERE REPLACE(SUBSTRING(sales_date, 1, 7),'-','') = '"&cost_month&"' "
objBuilder.Append "		AND mgct.saupbu = saupbu) / "
objBuilder.Append "	(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "	FROM saupbu_sales "
objBuilder.Append "	WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&cost_month&"' "
objBuilder.Append "		AND saupbu <> 'ȸ�簣�ŷ�') AS sale_per, "
objBuilder.Append "	(SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "	FROM saupbu_sales "
objBuilder.Append "	WHERE REPLACE(SUBSTRING(sales_date,1,7), '-', '') = '"&cost_month&"' "
objBuilder.Append "		AND mgct.saupbu = saupbu) * "&prosCost&" AS saupbu_sale_amt /* ��������2 */, "
objBuilder.Append "	mgct.tot_cost_amt, "
objBuilder.Append "	(mgct.saupbu_person * "&privCost&") + "
objBuilder.Append "	((SELECT IFNULL(SUM(cost_amt), 0) AS sales_amt "
objBuilder.Append "	FROM saupbu_sales "
objBuilder.Append "	WHERE REPLACE(SUBSTRING(sales_date, 1, 7), '-', '') = '"&cost_month&"' "
objBuilder.Append "		AND mgct.saupbu = saupbu) * "&prosCost&") AS all_tot_cost_amt "
objBuilder.Append "FROM management_cost AS mgct "
objBuilder.Append "WHERE mgct.cost_month ='"&cost_month&"' "
objBuilder.Append "	AND mgct.saupbu NOT IN ('��Ÿ�����', 'OA���ົ��') "
objBuilder.Append "GROUP BY mgct.saupbu "
objBuilder.Append "ORDER BY FIELD(mgct.saupbu, '����SI����', '����SI����', 'ICT����', 'NI����', '��������', 'SI2����', 'SI1����') DESC "

Set rsComm = Server.CreateObject("ADODB.RecordSet")
rsComm.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()

If saupbu = "" Then
    If rsComm.EOF Then
        saupbu = ""
    Else
        saupbu = rsComm("saupbu")
    End If
End If

title_line = "����� �ο� �� ���� ��� ���� ��Ȳ"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/sales_header.asp" -->
			<!--#include virtual = "/include/profit_loss_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<h3 class="stit">1. �������� ��� ������ ����κ� ���Ϳ��� �ο���, ���纰������ �ش� ����γ��� ����� ������ �����. </h3>
				<h3 class="stit">2. ���纰���Ϳ� ������ ����� ����γ��� ����� ������ �����. </h3>
				<form action="/sales/management_cost_report.asp" method="post" name="frm">
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
									<a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�"></a>

								</p>
							</dd>
						</dl>
					</fieldset>
				</form>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="52%" height="356" valign="top">
				      	<h3 class="stit">* ����κ� �ο� ��Ȳ �� ����</h3>
				      	<table cellpadding="0" cellspacing="0" class="tableList">
                            <colgroup>
                                <col width="*" >
                                <col width="7%" >
                                <col width="10%" >
                                <col width="12%" >
                                <col width="12%" >

                                <col width="14%" >
                                <col width="10%" >
                                <col width="12%" >
                            </colgroup>
				        	<thead>
                            <tr>
                                <th class="first" scope="col" rowspan="2">�����</th>
                                <th scope="col" colspan="4" style="border-bottom:1px solid #e3e3e3;">��������(�ο�)</th>
                                <th scope="col" colspan="3" style="border-bottom:1px solid #e3e3e3;">��������(����)</th>
                                <th scope="col" rowspan="2" style="border-bottom:1px solid #e3e3e3;">���������հ�</th>
                            </tr>
                            <tr>
                                <th scope="col" style="border-left:1px solid #e3e3e3;">�����<br>�η�</th>
                                <th scope="col">������(%)</th>
                                <th scope="col">��������</th>
                                <th scope="col">������</th>

                                <th scope="col">����θ���</th>
                                <th scope="col">������(%)</th>
                                <th scope="col">��������</th>
                            </tr>
                            </thead>
			                <tbody>
			            	<%
							Dim tot_saupbu_person, tot_saupbu_cost_amt, tot_saupbu_per, tot_saupbu_direct
							Dim tot_saupbu_sale, tot_sale_per, tot_saupbu_sale_amt, all_tot_saupbu_sale_amt
							Dim rs_etc, direct_cost

                            tot_saupbu_person   = 0
                            tot_saupbu_cost_amt = 0
                            tot_saupbu_per      = 0
                            tot_saupbu_direct   = 0

                            tot_saupbu_sale     = 0
                            tot_sale_per        = 0
                            tot_saupbu_sale_amt = 0
                            all_tot_saupbu_sale_amt = 0

                            i = 0
                            Do Until rsComm.EOF
                                i = i + 1

                                'sql = "select sum(cost_amt_"&cost_mm&")      " & chr(13) &_
                                '      "  from company_cost                   " & chr(13) &_
                                '      " where (cost_center = '������' )      " & chr(13) &_
                                '      "   and (saupbu = '"&rs("saupbu")&"' ) " & chr(13) &_
                                '      "   and cost_year ='"&cost_year&"'     "
                                objBuilder.Append "SELECT SUM(cost_amt_"&cost_mm&") "
								objBuilder.Append "FROM company_cost "
								objBuilder.Append "WHERE cost_center = '������' "
								objBuilder.Append "	AND saupbu = '"&rsComm("saupbu")&"' "
								objBuilder.Append "	AND cost_year ='"&cost_year&"' "

                                Set rs_etc = DBConn.Execute(objBuilder.ToString())
								objBuilder.Clear()

                                If rs_etc(0) = "" Or IsNull(rs_etc(0)) Then 
                                    direct_cost = 0
                                Else 
                                    direct_cost = CDbl(rs_etc(0))
                                End If								
                                rs_etc.close()

                                saupbu_tab(i,1) = rsComm("saupbu")
                                saupbu_tab(i,2) = direct_cost

                                tot_saupbu_person   = tot_saupbu_person + CDbl(rsComm("saupbu_person"))
                                tot_saupbu_cost_amt = tot_saupbu_cost_amt + CDbl(rsComm("saupbu_cost_amt"))
                                tot_saupbu_per      = tot_saupbu_per + rsComm("saupbu_per")
                                tot_saupbu_direct   = tot_saupbu_direct + direct_cost

                                tot_saupbu_sale     = tot_saupbu_sale + rsComm("saupbu_sale")
                                tot_sale_per        = tot_sale_per + rsComm("sale_per")
                                tot_saupbu_sale_amt = tot_saupbu_sale_amt + rsComm("saupbu_sale_amt")
								all_tot_saupbu_sale_amt = all_tot_saupbu_sale_amt+ rsComm("all_tot_cost_amt")
                                %>
                                <tr>
                                    <!--�����     --> <td class="first"><a href="/sales/management_cost_report.asp?saupbu=<%=rsComm("saupbu")%>&cost_month=<%=cost_month%>&ck_sw=<%="y"%>"><%=rsComm("saupbu")%></a></td>
                                    <!--������η� --> <td class="right"><%=FormatNumber(rsComm("saupbu_person"), 0)%>&nbsp;</td>
                                    <!--������     --> <td class="right"><%=FormatNumber(rsComm("saupbu_per")*100, 3)%>%&nbsp;</td>
                                    <!--�������� --> <td class="right"><%=FormatNumber(rsComm("saupbu_cost_amt"), 0)%>&nbsp;</td>
                                    <!--������     --> <td class="right"><%=FormatNumber(direct_cost, 0)%>&nbsp;</td>

                                    <!--����θ��� --> <td class="right"><%=FormatNumber(rsComm("saupbu_sale"), 0)%>&nbsp;</td>
                                    <!--������     --> <td class="right"><%=FormatNumber(rsComm("sale_per")*100, 3)%>%&nbsp;</td>
                                    <!--�������� --> <td class="right"><%=FormatNumber(rsComm("saupbu_sale_amt"), 0)%>&nbsp;</td>
                                    <!--�������� --> <td class="right"><%=FormatNumber(rsComm("all_tot_cost_amt"), 0)%>&nbsp;</td>
                                </tr>
                                <%
				        	    rsComm.MoveNext()
				        	Loop 
							Set rs_etc = Nothing 
				        	rsComm.close() : Set rsComm = Nothing 
				        	%>
				            <tr bgcolor="#FFE8E8">
                                                      <td class="first">��</td>
                                <!--������η� �� --> <td class="right"><%=FormatNumber(tot_saupbu_person, 0)%>&nbsp;</td>
                                <!--������     �� --> <td class="right"><%=FormatNumber(tot_saupbu_per*100, 3)%>%&nbsp;</td>
                                <!--�������� �� --> <td class="right"><%=FormatNumber(tot_saupbu_cost_amt, 0)%>&nbsp;</td>
                                <!--������     �� --> <td class="right"><%=FormatNumber(tot_saupbu_direct, 0)%>&nbsp;</td>

                                <!--����θ��� �� --> <td class="right"><%=FormatNumber(tot_saupbu_sale, 0)%>&nbsp;</td>
                                <!--������     �� --> <td class="right"><%=FormatNumber(tot_sale_per*100, 3)%>%&nbsp;</td>
                                <!--�������� �� --> <td class="right"><%=FormatNumber(tot_saupbu_sale_amt, 0)%>&nbsp;</td>
                                <!--�������� �� --> <td class="right"><%=FormatNumber(all_tot_saupbu_sale_amt, 0)%>&nbsp;</td>
                            </tr>
                            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="46%" valign="top">
				      	<h3 class="stit">* ����γ� ȸ�纰 ����� ����</h3>
				        <table cellpadding="0" cellspacing="0" summary="" class="tableList">
				        <colgroup>
				          <col width="20%" >
				          <col width="*" >
				          <col width="20%" >
			            </colgroup>
				        <thead>
                            <tr>
                                <th class="first" scope="col">�����</th>
                                <th scope="col">����</th>
                                <th scope="col">����</th>
                            </tr>
                        </thead>
			            <tbody>
                            <%
							Dim tot_cost_amt, tot_charge_per, tot_company_cost, salesDate
							Dim rsSales

                            tot_cost_amt = 0
                            tot_charge_per = 0
                            tot_company_cost = 0

                            salesDate = LEFT(cost_month, 4) & "-" & RIGHT(cost_month, 2)

                            'sql = "    SELECT saupbu, company, sum(cost_amt) as cost_amt   " & chr(13) &_
                            '      "      FROM saupbu_sales                                 " & chr(13) &_
                            '      "     WHERE substring(sales_date,1,7) = '"&salesDate&"'  " & chr(13) &_
                            '      "       AND saupbu ='"&saupbu&"'                         " & chr(13) &_
                            '      "  GROUP BY saupbu ,company                              "
							objBuilder.Append "SELECT saupbu, company, sum(cost_amt) as cost_amt "
							objBuilder.Append "FROM saupbu_sales "
							objBuilder.Append "WHERE substring(sales_date,1,7) = '"&salesDate&"' "
							objBuilder.Append "	AND saupbu ='"&saupbu&"'"
							objBuilder.Append "GROUP BY saupbu ,company "

							Set rsSales = Server.CreateObject("ADODB.RecordSet")
                            rsSales.Open objBuilder.ToString(), DBConn, 1
							objBuilder.Clear()

                            Do Until rsSales.EOF
                                tot_cost_amt = tot_cost_amt + rsSales("cost_amt")
                                %>
                                <tr>
                                    <td class="first"><%=rsSales("saupbu")%></td>
                                    <td><%=rsSales("company")%>&nbsp;</td>
                                    <td class="right"><%=FormatNumber(rsSales("cost_amt"), 0)%>&nbsp;</td>
                                </tr>
                                <%
                                rsSales.MoveNext()
                            Loop 
                            rsSales.close() : Set rsSales = Nothing 
                            %>
                            <tr bgcolor="#FFE8E8">
                                <td class="first">��</td>
                                <td class="right">&nbsp;</td>
                                <td class="right"><%=FormatNumber(tot_cost_amt, 0)%>&nbsp;</td>
                            </tr>
			            </tbody>
			            </table>

                        <%
						Dim rs_emp
						'20170529 KDC����� �� ��� ���� ����Ʈ ���

                        'If Trim(request.cookies("nkpmg_user")("coo_saupbu")&"") = "KDC�����" Then
						If Trim(saupbu) = "����SI����" Then

                            'sql = "SELECT A.pmg_yymm                              " & chr(13) &_
                            '      "     , B.emp_name                              " & chr(13) &_
                            '      "     , B.emp_job                               " & chr(13) &_
                            '      "     , B.emp_type                              " & chr(13) &_
                            '      "     , if(B.cost_except=2,'Y','N') cost_except " & chr(13) &_
                            '      "  FROM pay_month_give  A                       " & chr(13) &_
                            '      "     , emp_master_month B                      " & chr(13) &_
                            '      " WHERE A.pmg_id     = '1'                      " & chr(13) &_
                            '      "   AND A.pmg_emp_no =  B.emp_no                " & chr(13) &_
                            '      "   AND B.cost_except in ('0','1') /*��������*/ " & chr(13) &_
                            '      "   AND A.pmg_yymm   = '" & cost_month & "'     " & chr(13) &_
                            '      "   AND B.emp_month  = '" & cost_month & "'     " & chr(13) &_
                            '      "   AND A.mg_saupbu  = '" & saupbu & "'         "
							objBuilder.Append "SELECT pmgt.pmg_yymm, emmt.emp_name, emmt.emp_job, emmt.emp_type, "
							objBuilder.Append "	IF(emmt.cost_except = 2, 'Y', 'N') AS cost_except "
							objBuilder.Append "FROM pay_month_give AS pmgt "
							objBuilder.Append "INNER JOIN emp_master_month AS emmt ON pmgt.pmg_emp_no = emmt.emp_no "
							objBuilder.Append "	AND emmt.emp_month = '"&cost_month&"' "
							objBuilder.Append "WHERE pmgt.pmg_id = '1' "
							objBuilder.Append "	AND pmgt.pmg_yymm = '"&cost_month&"' "
							objBuilder.Append "	AND emmt.cost_except IN ('0', '1') "
							objBuilder.Append "	AND pmgt.mg_saupbu = '"&saupbu&"' "						

                            Set rs_emp = DBConn.Execute(objBuilder.ToString())
							objBuilder.Clear()
                            %>
                            <h3 class="stit">* �η� ����Ʈ</h3>
                            <table cellpadding="0" cellspacing="0" summary="" class="tableList" style="width:350px;">
                                <colgroup>
                                    <col width="56%" >
                                    <col width="22%" >
                                    <col width="22%" >
                                </colgroup>
                                <thead>
                                    <tr>
                                    <th class="first" scope="col">�̸�</th>
                                    <th scope="col">����</th>
                                    <th scope="col">���� ����</th>
                                    </tr>
                                </thead>
                                <tbody>
                                <%
                                If Not(rs_emp.BOF Or rs_emp.EOF) Then
                                    Do Until rs_emp.EOF
                                        %>
                                        <tr>
                                        <td><%=rs_emp("emp_name")%>&nbsp;<%=rs_emp("emp_job")%></td>
                                        <td><%=rs_emp("emp_type")%></td>
                                        <td><%=rs_emp("cost_except")%></td>
                                        </tr>
                                        <%
                                        rs_emp.MoveNext()
                                    Loop
                                End If
								rs_emp.Close() : Set rs_emp = Nothing 
                                %>
                                </tbody>
                            </table>
                            <%
                        End If
						DBConn.Close() : Set DBConn = Nothing 
                        %>
			          </td>
			        </tr>

				    <tr>
				      <td width="46%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="52%">&nbsp;</td>
			        </tr>
			      </table>
                </div>
			</div>
	</div>
	</body>
</html>

