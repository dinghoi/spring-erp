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
Dim mg_saupbu, cost_month, before_month, cost_id, cost_detail, j
Dim cost_center, cost_year, cost_mm, from_date, end_date, to_date
Dim com_sql, pmg_com_sql, company_sql, de_com_sql, over_com_sql, tran_com_sql, card_com_sql
Dim bonbu_sql, pmg_bonbu_sql, de_bonbu_sql, over_bonbu_sql, tran_bonbu_sql, card_bonbu_sql
Dim saupbu_sql, pmg_saupbu_sql, de_saupbu_sql, over_saupbu_sql, tran_saupbu_sql, card_saupbu_sql
Dim rsCost, title_line

mg_saupbu = f_Request("mg_saupbu")
cost_month = f_Request("cost_month")
before_month = f_Request("before_month")
cost_id = f_Request("cost_id")
cost_detail = f_Request("cost_detail")
j = f_Request("j")

If j = 1 Or j = 6 Then
	cost_center = "상주직접비"
Else
  	cost_center = "직접비"
End If

If j = 1 Or j = 2 Then
	cost_year = CStr(Mid(before_month, 1, 4))
	cost_mm = CStr(Mid(before_month, 5, 2))

	cost_month = before_month
Else
	cost_year = CStr(Mid(cost_month, 1, 4))
	cost_mm = CStr(Mid(cost_month, 5, 2))
End If

from_date = CStr(cost_year) & "-" & CStr(cost_mm) & "-01"
end_date = DateValue(from_date)
end_date = DateAdd("m", 1, from_date)
to_date = CStr(DateAdd("d", -1, end_date))

'//기타사업부 변경
If mg_saupbu = "" Then
	mg_saupbu = ""
End If

If emp_company = "" Then
	com_sql = ""
	pmg_com_sql = ""
	company_sql = ""
	de_com_sql = ""
  	over_com_sql = " "
  	tran_com_sql = " "
  	card_com_sql = " "
Else
  	com_sql = " AND emp_company ='"&emp_company&"' "
  	pmg_com_sql = " AND pmg_company ='"&emp_company&"' "
  	company_sql = " AND company ='"&emp_company&"' "
  	de_com_sql = " AND de_company ='"&emp_company&"' "
  	over_com_sql = " AND overtime.emp_company ='"&emp_company&"' "
  	tran_com_sql = " AND transit_cost.emp_company ='"&emp_company&"' "
  	card_com_sql = " AND card_slip.emp_company ='"&emp_company&"' "
End If

If bonbu = "" Then
	bonbu_sql = ""
	pmg_bonbu_sql = ""
	de_bonbu_sql = ""
  	over_bonbu_sql = " "
  	tran_bonbu_sql = " "
  	card_bonbu_sql = " "
Else
  	bonbu_sql = " AND bonbu ='"&bonbu&"' "
  	pmg_bonbu_sql = " AND pmg_bonbu ='"&bonbu&"' "
  	de_bonbu_sql = " AND de_bonbu ='"&bonbu&"' "
  	over_bonbu_sql = " AND overtime.bonbu ='"&bonbu&"' "
  	tran_bonbu_sql = " AND transit_cost.bonbu ='"&bonbu&"' "
  	card_bonbu_sql = " AND card_slip.bonbu ='"&bonbu&"' "
End If

If saupbu = "" Then
	saupbu_sql = ""
	pmg_saupbu_sql = ""
	de_saupbu_sql = ""
  	over_saupbu_sql = " "
  	tran_saupbu_sql = " "
  	card_saupbu_sql = " "
Else
  	saupbu_sql = " AND saupbu ='"&saupbu&"' "
  	pmg_saupbu_sql = " AND pmg_saupbu ='"&saupbu&"' "
  	de_saupbu_sql = " AND de_saupbu ='"&saupbu&"' "
  	over_saupbu_sql = " AND overtime.saupbu ='"&saupbu&"' "
  	tran_saupbu_sql = " AND transit_cost.saupbu ='"&saupbu&"' "
  	card_saupbu_sql = " AND card_slip.saupbu ='"&saupbu&"' "
End If

If cost_id = "일반경비" Then
	'sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,company,customer,slip_memo,cost FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and mg_saupbu = '"&mg_saupbu&"' and cost_center = '"&cost_center&"' and slip_gubun = '비용' and account = '"&cost_detail&"' order by org_name,emp_name,slip_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT org_name, slip_date, emp_name AS user_name, emp_grade AS user_grade, company, "
	objBuilder.Append "	customer, slip_memo, cost "
	objBuilder.Append "FROM general_cost "
	objBuilder.Append "WHERE cancel_yn = 'N' AND (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') "
	objBuilder.Append "	AND mg_saupbu = '"&mg_saupbu&"' AND cost_center = '"&cost_center&"' "
	objBuilder.Append "	AND slip_gubun = '비용' AND account = '"&cost_detail&"' "
	objBuilder.Append "ORDER BY org_name, emp_name, slip_date ASC "
End If

If cost_id = "임차료" Or cost_id = "외주비" Or cost_id = "자재" Or cost_id = "장비" Or cost_id = "운반비" Then
	'sql = "select slip_date,slip_seq,mg_saupbu as org_name,slip_date,emp_name as user_name,emp_grade as user_grade,company,customer,slip_memo,cost FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and mg_saupbu = '"&mg_saupbu&"' and cost_center = '"&cost_center&"' and slip_gubun = '"&cost_id&"' and account = '"&cost_detail&"' order by org_name,emp_name,slip_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT slip_date, slip_seq, mg_saupbu AS org_name, slip_date, emp_name AS user_name, "
	objBuilder.Append "	emp_grade AS user_grade, company, customer, slip_memo, cost "
	objBuilder.Append "FROM general_cost "
	objBuilder.Append "WHERE (cancel_yn = 'N') AND (slip_date >= '"&from_date&"' AND slip_date <= '"&to_date&"') "
	objBuilder.Append "	AND mg_saupbu = '"&mg_saupbu&"' AND cost_center = '"&cost_center&"' "
	objBuilder.Append "	AND slip_gubun = '"&cost_id&"' AND account = '"&cost_detail&"' "
	objBuilder.Append "ORDER BY org_name, emp_name, slip_date ASC "
End If

If cost_id = "법인카드" Then
	'sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,reside_company as company,customer,concat(account,'-',account_item) as slip_memo,cost FROM card_slip where card_type not like '%주유%' and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and account = '"&cost_detail&"' order by org_name,emp_name,slip_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT org_name, slip_date, emp_name AS user_name, emp_grade AS user_grade, reside_company AS company, "
	objBuilder.Append "	customer, CONCAT(account, '-', account_item) AS slip_memo, cost "
	objBuilder.Append "FROM card_slip "
	objBuilder.Append "WHERE card_type NOT LIKE '%주유%' AND (slip_date >= '"&from_date&"' AND slip_date <= '"&to_date&"') "
	objBuilder.Append "	AND cost_center = '"&cost_center&"' AND mg_saupbu = '"&mg_saupbu&"' AND account = '"&cost_detail&"' "
	objBuilder.Append "ORDER BY org_name, emp_name, slip_date ASC "
End If

If cost_id = "교통비" And cost_detail = "대중교통" Then
	'sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,concat(transit,'-',run_memo) as slip_memo,(fare) as cost FROM transit_cost  where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '"&cost_detail&"' order by org_name,user_name,run_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT org_name, run_date AS slip_date, user_name, user_grade, company, '' AS customer, "
	objBuilder.Append "	CONCAT(transit, '-', run_memo) AS slip_memo, (fare) AS cost "
	objBuilder.Append "FROM transit_cost "
	objBuilder.Append "WHERE cancel_yn = 'N' AND (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') "
	objBuilder.Append "	AND cost_center = '"&cost_center&"' AND mg_saupbu = '"&mg_saupbu&"' AND car_owner = '"&cost_detail&"' "
	objBuilder.Append "ORDER BY org_name, user_name, run_date ASC "
End If

If cost_id = "교통비" And cost_detail = "회사" Then
	'sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,run_memo as slip_memo, (somopum+oil_price+parking+toll) as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '"&cost_detail&"' order by org_name,user_name,run_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT org_name, run_date AS slip_date, user_name, user_grade, company, "
	objBuilder.Append "	'' as customer, run_memo AS slip_memo, (somopum + oil_price + parking + toll) AS cost "
	objBuilder.Append "FROM transit_cost "
	objBuilder.Append "WHERE cancel_yn = 'N' AND (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') "
	objBuilder.Append "	AND cost_center = '"&cost_center&"' AND mg_saupbu = '"&mg_saupbu&"' AND car_owner = '"&cost_detail&"' "
	objBuilder.Append "ORDER BY org_name, user_name, run_date ASC "
End If

If cost_id = "교통비" And cost_detail = "개인" Then
	'sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,run_memo as slip_memo, (oil_price+somopum+parking+toll) as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '"&cost_detail&"' order by org_name,user_name,run_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT org_name, run_date AS slip_date, user_name, user_grade, company, "
	objBuilder.Append "	'' AS customer, run_memo AS slip_memo, (oil_price + somopum + parking + toll) AS cost "
	objBuilder.Append "FROM transit_cost "
	objBuilder.Append "WHERE cancel_yn = 'N' AND (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') "
	objBuilder.Append "	AND cost_center = '"&cost_center&"' AND mg_saupbu = '"&mg_saupbu&"' AND car_owner = '"&cost_detail&"' "
	objBuilder.Append "ORDER BY org_name, user_name, run_date ASC "
End If

If cost_id = "교통비" And cost_detail = "차량수리비" Then
	'sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,run_memo as slip_memo,repair_cost as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '회사'"&tran_com_sql&tran_bonbu_sql&tran_saupbu_sql&" order by org_name,user_name,run_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT org_name, run_date AS slip_date, user_name, user_grade, company, "
	objBuilder.Append "	'' AS customer, run_memo AS slip_memo, repair_cost AS cost "
	objBuilder.Append "FROM transit_cost "
	objBuilder.Append "WHERE cancel_yn = 'N' AND (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') "
	objBuilder.Append "	AND cost_center = '"&cost_center&"' AND mg_saupbu = '"&mg_saupbu&"' AND car_owner = '회사' "
	objBuilder.Append tran_com_sql & tran_bonbu_sql & tran_saupbu_sql & " "
	objBuilder.Append "ORDER BY org_name, user_name, run_date ASC "
End If

'//2016-08-23 알바비 추가
If cost_id = "인건비" and cost_detail = "알바비" Then
	'sql = "select org_name,give_date as slip_date,draft_man as user_name,'' as user_grade,cost_company as company,'' as customer,work_comment as slip_memo,alba_give_total as cost "
	'sql = sql & " from pay_alba_cost "
	'sql = sql & " Where saupbu='"&mg_saupbu&"' "	'기존 주석
	'sql = sql & "	Where bonbu ='"&mg_saupbu&"'"
	'sql = sql & " and give_date between '"&from_date&"' and '"&to_date&"' "
	'sql = sql & " order by give_date asc"
	'rs.Open sql, Dbconn, 1

	objBuilder.Append "SELECT org_name, give_date AS slip_date, draft_man AS user_name, '' AS user_grade, cost_company AS company, "
	objBuilder.Append "	'' AS customer, work_comment AS slip_memo, alba_give_total AS cost "
	objBuilder.Append "FROM pay_alba_cost "
	objBuilder.Append "WHERE bonbu ='"&mg_saupbu&"' "
	'objBuilder.Append "	AND give_date BETWEEN '"&from_date&"' AND '"&to_date&"' "
	objBuilder.Append "	AND rever_yymm = '"&cost_month&"' "
	objBuilder.Append "ORDER BY give_date ASC "
End If

Set rsCost = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

title_line = mg_saupbu & " 사업부별 세부 손익 내역("&cost_id&"-"&cost_detail&")"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html lang="ko">
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}
		</script>
	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>비용년월 : </strong><%=cost_year%>년<%=cost_mm%>월&nbsp;
							<strong>비용유형 : </strong><%=cost_center%>&nbsp;
							<strong>비용구분 : </strong><%=cost_id%>&nbsp;-&nbsp;<%=cost_detail%>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="10%" >
							<col width="10%" >
							<col width="8%" >
							<col width="14%" >
							<col width="16%" >
							<col width="*" >
							<col width="9%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">조직</th>
								<th scope="col">담당자</th>
								<th scope="col">비용일자</th>
								<th scope="col">고객사</th>
								<th scope="col">거래처</th>
								<th scope="col">사용내역</th>
								<th scope="col">사용금액</th>
							</tr>
						</thead>
						<tbody>
         				<%
							Dim cost_cnt, cost_sum, i

							cost_cnt = 0
							cost_sum = 0
							i = 0

							Do Until rsCost.EOF
								i = i + 1

								If rsCost("cost") <> "0" then
									cost_sum = cost_sum + clng(rsCost("cost"))
									cost_cnt = cost_cnt + 1
						%>
							<tr>
								<td class="first"><%=cost_cnt%></td>
								<td><%=rsCost("org_name")%>&nbsp;</td>
								<td><%=rsCost("user_name")%>&nbsp;<%=rsCost("user_grade")%></td>
								<td><%=rsCost("slip_date")%></td>
								<td><%=rsCost("company")%>&nbsp;</td>
								<td class="left"><%=rsCost("customer")%>&nbsp;</td>
								<td class="left">
								<%
									If (cost_id = "임차료" Or cost_id = "외주비" Or cost_id = "자재" Or cost_id = "장비" Or cost_id = "운반비") And (user_id = "900001" Or user_id = "100359" Or user_id = "102592") Then
								%>
									<a href="#" onClick="pop_Window('/sales/tax_bill_in_mod.asp?slip_date=<%=rsCost("slip_date")%>&slip_seq=<%=rsCost("slip_seq")%>&u_type=<%="U"%>','tax_bill_in_mod_pop','scrollbars=yes,width=1000,height=280')"><%=rsCost("slip_memo")%></a>
								<%	Else	%>
									<%=rsCost("slip_memo")%>
								<%	End If %>
                                </td>
								<td class="right"><%=FormatNumber(rsCost("cost"), 0)%></td>
							</tr>
							<%
								End If
								rsCost.MoveNext()
							Loop
							rsCost.Close()
							%>
							<tr>
								<th colspan="7" class="first">합계</th>
								<th class="right"><%=FormatNumber(cost_sum, 0)%></th>
							</tr>
						</tbody>
					</table>
				</div>
	</form>
	</body>
</html>
<!--#include virtual="/common/log_sales_profit.asp" -->
<%
DBConn.Close() : Set DBConn = Nothing
%>