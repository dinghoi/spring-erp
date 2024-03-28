<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
mg_saupbu = Request("mg_saupbu")
cost_month = Request("cost_month")
before_month = Request("before_month")
cost_id = Request("cost_id")
cost_detail = Request("cost_detail")
j = Request("j")
if j = 1 or j = 6 then
	cost_center = "상주직접비"
  else
  	cost_center = "직접비"
end if
if j = 1 or j = 2 then
	cost_year = cstr(mid(before_month,1,4))
	cost_mm = cstr(mid(before_month,5,2))
  else
	cost_year = cstr(mid(cost_month,1,4))
	cost_mm = cstr(mid(cost_month,5,2))
end if

from_date = cstr(cost_year) + "-" + cstr(cost_mm) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

if mg_saupbu = "기타사업부" then
	mg_saupbu = ""
end if
if emp_company = "" then
	com_sql = ""
	pmg_com_sql = ""
	company_sql = ""
	de_com_sql = ""
  	over_com_sql = " "
  	tran_com_sql = " "
  	card_com_sql = " "
  else
  	com_sql = " and emp_company ='"&emp_company&"' "
  	pmg_com_sql = " and pmg_company ='"&emp_company&"' "
  	company_sql = " and company ='"&emp_company&"' "
  	de_com_sql = " and de_company ='"&emp_company&"' "
  	over_com_sql = " and overtime.emp_company ='"&emp_company&"' "
  	tran_com_sql = " and transit_cost.emp_company ='"&emp_company&"' "
  	card_com_sql = " and card_slip.emp_company ='"&emp_company&"' "
end if
if bonbu = "" then
	bonbu_sql = ""
	pmg_bonbu_sql = ""
	de_bonbu_sql = ""
  	over_bonbu_sql = " "
  	tran_bonbu_sql = " "
  	card_bonbu_sql = " "
  else
  	bonbu_sql = " and bonbu ='"&bonbu&"' "
  	pmg_bonbu_sql = " and pmg_bonbu ='"&bonbu&"' "
  	de_bonbu_sql = " and de_bonbu ='"&bonbu&"' "
  	over_bonbu_sql = " and overtime.bonbu ='"&bonbu&"' "
  	tran_bonbu_sql = " and transit_cost.bonbu ='"&bonbu&"' "
  	card_bonbu_sql = " and card_slip.bonbu ='"&bonbu&"' "
end if
if saupbu = "" then
	saupbu_sql = ""
	pmg_saupbu_sql = ""
	de_saupbbu_sql = ""
  	over_saupbu_sql = " "
  	tran_saupbu_sql = " "
  	card_saupbu_sql = " "
  else
  	saupbu_sql = " and saupbu ='"&saupbu&"' "
  	pmg_saupbu_sql = " and pmg_saupbu ='"&saupbu&"' "
  	de_saupbu_sql = " and de_saupbu ='"&saupbu&"' "
  	over_saupbu_sql = " and overtime.saupbu ='"&saupbu&"' "
  	tran_saupbu_sql = " and transit_cost.saupbu ='"&saupbu&"' "
  	card_saupbu_sql = " and card_slip.saupbu ='"&saupbu&"' "
end if

if cost_id = "일반경비" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,company,customer,slip_memo,cost FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and mg_saupbu = '"&mg_saupbu&"' and cost_center = '"&cost_center&"' and slip_gubun = '비용' and account = '"&cost_detail&"' order by org_name,emp_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "임차료" or cost_id = "외주비" or cost_id = "자재" or cost_id = "장비" or cost_id = "운반비" then
	sql = "select slip_date,slip_seq,mg_saupbu as org_name,slip_date,emp_name as user_name,emp_grade as user_grade,company,customer,slip_memo,cost FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and mg_saupbu = '"&mg_saupbu&"' and cost_center = '"&cost_center&"' and slip_gubun = '"&cost_id&"' and account = '"&cost_detail&"' order by org_name,emp_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "법인카드" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,reside_company as company,customer,concat(account,'-',account_item) as slip_memo,cost FROM card_slip where card_type not like '%주유%' and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and account = '"&cost_detail&"' order by org_name,emp_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "교통비" and cost_detail = "대중교통" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,concat(transit,'-',run_memo) as slip_memo,(fare) as cost FROM transit_cost  where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '"&cost_detail&"' order by org_name,user_name,run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "교통비" and cost_detail = "회사" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,run_memo as slip_memo, (somopum+oil_price+parking+toll) as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '"&cost_detail&"' order by org_name,user_name,run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "교통비" and cost_detail = "개인" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,run_memo as slip_memo, (oil_price+somopum+parking+toll) as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '"&cost_detail&"' order by org_name,user_name,run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "교통비" and cost_detail = "차량수리비" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,company,'' as customer,run_memo as slip_memo,repair_cost as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and cost_center = '"&cost_center&"' and mg_saupbu = '"&mg_saupbu&"' and car_owner = '회사'"&tran_com_sql&tran_bonbu_sql&tran_saupbu_sql&" order by org_name,user_name,run_date asc"
	rs.Open sql, Dbconn, 1
end if

title_line = mg_saupbu + " 사업부별 세부 손익 내역"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
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
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
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
							cost_cnt = 0
							cost_sum = 0
							i = 0
							do until rs.eof
								i = i + 1
								if rs("cost") <> "0" then
									cost_sum = cost_sum + clng(rs("cost"))
									cost_cnt = cost_cnt + 1
							%>
							<tr>
								<td class="first"><%=cost_cnt%></td>
								<td><%=rs("org_name")%>&nbsp;</td>
								<td><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("company")%>&nbsp;</td>
								<td class="left"><%=rs("customer")%>&nbsp;</td>
								<td class="left">
							<% if (cost_id = "임차료" or cost_id = "외주비" or cost_id = "자재" or cost_id = "장비" or cost_id = "운반비") and (user_id = "900001" or user_id = "100359") then	%>						
								<a href="#" onClick="pop_Window('tax_bill_in_mod.asp?slip_date=<%=rs("slip_date")%>&slip_seq=<%=rs("slip_seq")%>&u_type=<%="U"%>','tax_bill_in_mod_pop','scrollbars=yes,width=1000,height=280')"><%=rs("slip_memo")%></a>
							<%   else	%>
								<%=rs("slip_memo")%>
							<% end if	%>
                                </td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							</tr>
							<%
								end if
								rs.movenext()
							loop
							rs.close()
							%>
							<tr>
								<th colspan="7" class="first">합계</th>
								<th class="right"><%=formatnumber(cost_sum,0)%></th>
							</tr>
						</tbody>
					</table>
				</div>				        				
	</form>
	</body>
</html>

