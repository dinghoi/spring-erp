<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
emp_company = Request("emp_company")
bonbu = Request("bonbu")
saupbu = Request("saupbu")
team = Request("team")
emp_no = Request("emp_no")
cost_id = Request("cost_id")
cost_detail = Request("cost_detail")
cost_year = Request("cost_year")
cost_month = Request("cost_month")
cost_yymm = cstr(cost_year) + cstr(cost_month)
from_date = cstr(cost_year) + "-" + cstr(cost_month) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

if emp_company = "" then
	com_sql = ""
  else
  	com_sql = " and emp_company ='"&emp_company&"' "
end if

if cost_id = "�ΰǺ�" and cost_detail = "�޿�" then
	sql = "select pmg_company as emp_company, pmg_org_name as org_name,pmg_yymm as slip_date,pmg_emp_name as user_name,pmg_grade as user_grade,pmg_id as slip_memo,pmg_give_total as cost FROM pay_month_give where pmg_yymm = '"&cost_yymm&"' and pmg_team = '"&team&"' and pmg_saupbu = '"&saupbu&"' and pmg_id = '1'"&com_sql&" order by pmg_org_name,pmg_emp_name asc"
	response.write(sql)
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�ΰǺ�" and cost_detail = "��" then
	sql = "select pmg_company as emp_company, pmg_org_name as org_name,pmg_yymm as slip_date,pmg_emp_name as user_name,pmg_grade as user_grade,pmg_id as slip_memo,pmg_give_total as cost FROM pay_month_give where pmg_yymm = '"&cost_yymm&"' and pmg_team = '"&team&"' and pmg_saupbu = '"&saupbu&"' and pmg_id = '1'"&com_sql&" order by pmg_org_name,pmg_emp_name asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�ΰǺ�" and cost_detail = "��������" then
	sql = "select pmg_company as emp_company, pmg_org_name as org_name,pmg_yymm as slip_date,pmg_emp_name as user_name,pmg_grade as user_grade,pmg_id as slip_memo,pmg_give_total as cost FROM pay_month_give where pmg_yymm = '"&cost_yymm&"' and pmg_team = '"&team&"' and pmg_saupbu = '"&saupbu&"' and pmg_id = '1'"&com_sql&" order by pmg_org_name,pmg_emp_name asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�ΰǺ�" and cost_detail = "�˹ٺ�" then
	sql = "select company as emp_company, org_name,rever_yymm as slip_date,draft_man as user_name,draft_tax_id as user_grade,draft_tax_id as slip_memo,alba_give_total as cost FROM pay_alba_cost where rever_yymm = '"&cost_yymm&"' and team = '"&team&"' and saupbu = '"&saupbu&"'"&com_sql&" order by org_name,draft_man asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�ΰǺ�" and cost_detail = "4�뺸��" then
	sql = "select de_company as emp_company, de_org_name as org_name,de_yymm as slip_date,de_emp_name as user_name,de_grade as user_grade,de_id as slip_memo,sum(de_nps_amt+de_nhis_amt+de_epi_amt+de_longcare_amt) as cost from pay_month_deduct where (de_saupbu = '"&saupbu&"') and (de_team ='"&team&"') and (de_yymm ='"&cost_yymm&"') and (de_id ='1')"&com_sql&" group by de_org_name,de_emp_name asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "��Ư��" then
	sql = "select overtime.org_name as org_name,work_date as slip_date,user_name,user_grade,work_item as slip_memo,overtime_amt as cost FROM overtime INNER JOIN memb ON overtime.mg_ce_id = memb.user_id where (cancel_yn = 'N') and  (work_date >= '"&from_date&"' and work_date <= '"&to_date&"') and overtime.team = '"&team&"' and overtime.saupbu = '"&saupbu&"' and overtime.cost_detail = '"&cost_detail&"'"&com_sql&" order by overtime.org_name,memb.user_name, overtime.work_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�Ϲݰ��" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,customer as slip_memo,cost FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and team = '"&team&"' and saupbu = '"&saupbu&"' and slip_gubun = '���' and account = '"&cost_detail&"'"&com_sql&" order by org_name,emp_name,slip_date asc"
	response.write(sql)
	rs.Open sql, Dbconn, 1
end if

if cost_id = "������" or cost_id = "���ֺ�" or cost_id = "����" or cost_id = "���" or cost_id = "��ݺ�" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,slip_memo,cost FROM general_cost where (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and team = '"&team&"' and saupbu = '"&saupbu&"' and slip_gubun = '"&cost_id&"' and account = '"&cost_detail&"'"&com_sql&" order by org_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "���߱���" then
	sql = "select transit_cost.org_name as org_name,run_date as slip_date,user_name,user_grade,concat(transit,'-',run_memo) as slip_memo,(fare) as cost FROM transit_cost INNER JOIN memb ON transit_cost.mg_ce_id = memb.user_id where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and transit_cost.team = '"&team&"' and transit_cost.saupbu = '"&saupbu&"' and transit_cost.car_owner = '"&cost_detail&"'"&com_sql&" order by transit_cost.org_name,memb.user_name, transit_cost.run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "ȸ��" then
	sql = "select transit_cost.org_name as org_name,run_date as slip_date,user_name,user_grade,concat(company,'-',run_memo) as slip_memo,(somopum+oil_price+parking+toll) as cost FROM transit_cost INNER JOIN memb ON transit_cost.mg_ce_id = memb.user_id where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and transit_cost.team = '"&team&"' and transit_cost.saupbu = '"&saupbu&"' and transit_cost.car_owner = '"&cost_detail&"'"&com_sql&" order by transit_cost.org_name,memb.user_name, transit_cost.run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "����" then
	sql = "select transit_cost.org_name as org_name,run_date as slip_date,user_name,user_grade,concat(company,'-',run_memo) as slip_memo,(oil_price+somopum+parking+toll) as cost FROM transit_cost INNER JOIN memb ON transit_cost.mg_ce_id = memb.user_id where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and transit_cost.team = '"&team&"' and transit_cost.saupbu = '"&saupbu&"' and transit_cost.car_owner = '"&cost_detail&"'"&com_sql&" order by transit_cost.org_name,memb.user_name, transit_cost.run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "����������" then
	sql = "select transit_cost.org_name as org_name,run_date as slip_date,user_name,user_grade,run_memo as slip_memo,repair_cost as cost FROM transit_cost INNER JOIN memb ON transit_cost.mg_ce_id = memb.user_id where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and transit_cost.team = '"&team&"' and transit_cost.saupbu = '"&saupbu&"' and transit_cost.car_owner = 'ȸ��'"&com_sql&" order by transit_cost.org_name,memb.user_name, transit_cost.run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "����ī��" then
	sql = "select card_slip.org_name as org_name,card_slip.slip_date,memb.user_name,memb.user_grade,concat(card_slip.customer,'-',card_slip.account_item) as slip_memo,card_slip.cost FROM card_slip INNER JOIN memb ON card_slip.emp_no = memb.user_id where card_type not like '%����%' and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and card_slip.team = '"&team&"' and card_slip.saupbu = '"&saupbu&"' and card_slip.account = '"&cost_detail&"'"&com_sql&" order by card_slip.org_name,memb.user_name, card_slip.slip_date asc"
	rs.Open sql, Dbconn, 1
end if
response.write(sql)
title_line = saupbu + " " + team + " ���κ� ��� ��Ȳ"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
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
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>��� : </strong><%=cost_year%>��<%=cost_month%>��&nbsp;
							<strong>��뱸�� : </strong><%=cost_id%>&nbsp;<%=cost_detail%>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="4%" >
							<col width="*" >
							<col width="14%" >
							<col width="14%" >
							<col width="35%" >
							<col width="13%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">�������</th>
								<th scope="col">��볻��</th>
								<th scope="col">���ݾ�</th>
							</tr>
						</thead>
						<tbody>
         					<% 
							cost_cnt = 0
							cost_sum = 0
							do until rs.eof
								if rs("cost") > "0" then
									cost_sum = cost_sum + clng(rs("cost"))
									cost_cnt = cost_cnt + 1
									user_grade_view = rs("user_grade")
									slip_memo_view = rs("slip_memo")
									if cost_id = "�ΰǺ�" and cost_detail = "4�뺸��" then
										slip_memo_view = "4�뺸��"							
									end if
									if cost_id = "�ΰǺ�" and cost_detail = "�˹ٺ�" then
										user_grade_view = "�˹�"							
									end if
									if cost_id = "�ΰǺ�" and cost_detail <> "4�뺸��" then
										if rs("slip_memo") = "1" then
											slip_memo_view = "�޿�"
										end if
										if rs("slip_memo") = "2" then
											slip_memo_view = "��"
										end if
										if rs("slip_memo") = "4" then
											slip_memo_view = "��������"
										end if
									end if
							%>
							<tr>
								<td class="first"><%=cost_cnt%></td>
								<td><%=rs("org_name")%></td>
								<td><%=rs("user_name")%>&nbsp;<%=user_grade_view%></td>
								<td><%=rs("slip_date")%></td>
								<td><%=slip_memo_view%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
							</tr>
							<%
								end if
								rs.movenext()
							loop
							rs.close()
							%>
							<tr>
								<th colspan="5" class="first">�հ�</th>
								<th class="right"><%=formatnumber(cost_sum,0)%></th>
							</tr>
						</tbody>
					</table>
				</div>				        				
	</form>
	</body>
</html>

