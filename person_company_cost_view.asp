<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
view_sw=Request("view_sw")
reside=Request("reside")
common=Request("common")
direct=Request("direct")
cost_id = Request("cost_id")
cost_detail = Request("cost_detail")
cost_year = Request("cost_year")
cost_month = Request("cost_month")
if cost_month < 10 then
	cost_month = "0" + cstr(cost_month)
end if
cost_yymm = cstr(cost_year) + cstr(cost_month)
from_date = cstr(cost_year) + "-" + cstr(cost_month) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

if view_sw = "1" then
	pmg_sql = " and cost_center = '����������' and pmg_reside_company = '"&reside&"'" 
	alba_sql = " and cost_center = '����������' and cost_company = '"&reside&"'" 
	de_sql = " and cost_center = '����������' and de_reside_company = '"&reside&"'" 
	over_sql = " and cost_center = '����������' and overtime.company = '"&reside&"'" 
	com_sql = " and cost_center = '����������' and company = '"&reside&"'"
	tran_sql = " and cost_center = '����������' and company = '"&reside&"'" 
	card_sql = " and cost_center = '����������' and reside_company = '"&reside&"'" 
end if
if view_sw = "2" then
	pmg_sql = " and cost_center = '"&common&"'" 
	alba_sql = " and cost_company = '"&reside&"'" 
	de_sql = " and cost_center = '"&common&"'" 
	over_sql = " and cost_center = '"&common&"'" 
	com_sql = " and cost_center = '"&common&"'" 
	tran_sql = " and cost_center = '"&common&"'" 
	card_sql = " and cost_center = '"&common&"'" 
end if
if view_sw = "3" then
	pmg_sql = " and mg_saupbu = '"&direct&"' and cost_center = '������'" 
	alba_sql = " and mg_saupbu = '"&direct&"' and cost_center = '������'" 
	de_sql = " and mg_saupbu = '"&direct&"' and cost_center = '������'" 
	over_sql = " and cost_center = '������' and mg_saupbu = '"&direct&"'" 
	com_sql = " and cost_center = '������' and mg_saupbu = '"&direct&"'" 
	card_sql = " and cost_center = '������' and mg_saupbu = '"&direct&"'"
	tran_sql = " and cost_center = '������' and mg_saupbu = '"&direct&"'" 
end if
if view_sw = "4" then
	pmg_sql = " and mg_saupbu = '"&direct&"' and cost_center = '����������'" 
	alba_sql = " and mg_saupbu = '"&direct&"' and cost_center = '����������'" 
	de_sql = " and mg_saupbu = '"&direct&"' and cost_center = '����������'" 
	over_sql = " and cost_center = '����������' and mg_saupbu = '"&direct&"'" 
	com_sql = " and cost_center = '����������' and mg_saupbu = '"&direct&"'" 
	card_sql = " and cost_center = '����������' and mg_saupbu = '"&direct&"'"
	tran_sql = " and cost_center = '����������' and mg_saupbu = '"&direct&"'" 
end if

if cost_id = "�ΰǺ�" and cost_detail = "�޿�" then
	sql = "select pmg_company as emp_company, pmg_org_name as org_name,pmg_yymm as slip_date,pmg_emp_name as user_name,pmg_grade as user_grade,pmg_id as slip_memo,(pmg_give_total) as cost FROM pay_month_give where pmg_yymm = '"&cost_yymm&"' and pmg_id = '1'"&pmg_sql&" order by pmg_org_name,pmg_emp_name asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�ΰǺ�" and cost_detail = "��" then
	sql = "select pmg_company as emp_company, pmg_org_name as org_name,pmg_yymm as slip_date,pmg_emp_name as user_name,pmg_grade as user_grade,pmg_id as slip_memo,pmg_give_total as cost FROM pay_month_give where pmg_yymm = '"&cost_yymm&"' and pmg_id = '2'"&pmg_sql&" order by pmg_org_name,pmg_emp_name asc"
	rs.Open sql, Dbconn, 1
end if

'if cost_id = "�ΰǺ�" and cost_detail = "��������" then
'	sql = "select pmg_company as emp_company, pmg_org_name as org_name,pmg_yymm as slip_date,pmg_emp_name as user_name,pmg_grade as user_grade,pmg_id as slip_memo,pmg_give_total as cost FROM pay_month_give where pmg_yymm = '"&cost_yymm&"' and pmg_id = '4'"&pmg_sql&" order by pmg_org_name,pmg_emp_name asc"
'	rs.Open sql, Dbconn, 1
'end if

'if cost_id = "�ΰǺ�" and cost_detail = "4�뺸��" then
'	sql = "select de_company as emp_company, de_org_name as org_name,de_yymm as slip_date,de_emp_name as user_name,de_grade as user_grade,de_id as slip_memo,sum(de_nps_amt+de_nhis_amt+de_epi_amt+de_longcare_amt) as cost from pay_month_deduct where (de_yymm ='"&cost_yymm&"') and (de_id ='1')"&de_sql&" group by de_org_name,de_emp_name asc"
'	rs.Open sql, Dbconn, 1
'end if

'if pay_grade <> "0" and cost_id = "�ΰǺ�" and (cost_detail = "�޿�" or cost_detail = "��" or cost_detail = "��������" or cost_detail = "4�뺸��") then
'	rs.close()
'	pay_sw = "N"
'	sql = "select pmg_company as emp_company, pmg_org_name as org_name,pmg_yymm as slip_date,pmg_emp_name as user_name,pmg_grade as user_grade,pmg_id as slip_memo,(pmg_give_total - pmg_job_support) as cost FROM pay_month_give where pmg_yymm = '' and pmg_id = '' order by pmg_org_name,pmg_emp_name asc"
'	rs.Open sql, Dbconn, 1
'end if

if cost_id = "�ΰǺ�" and cost_detail = "�˹ٺ�" then
	if team = "" then
	sql = "select company as emp_company, org_name,rever_yymm as slip_date,draft_man as user_name,draft_tax_id as user_grade,draft_tax_id as slip_memo,alba_give_total as cost FROM pay_alba_cost where rever_yymm = '"&cost_yymm&"' and (team = '' or isnull(team))"&company_sql&bonbu_sql&saupbu_sql&" order by org_name,draft_man asc"
	  else
	sql = "select company as emp_company, org_name,rever_yymm as slip_date,draft_man as user_name,draft_tax_id as user_grade,draft_tax_id as slip_memo,alba_give_total as cost FROM pay_alba_cost where rever_yymm = '"&cost_yymm&"'"&alba_sql&" order by org_name,draft_man asc"
	end if
	rs.Open sql, Dbconn, 1
end if

if cost_id = "��Ư��" then
	sql = "select org_name,work_date as slip_date,user_name,user_grade,work_item as slip_memo,overtime_amt as cost FROM overtime where (cancel_yn = 'N') and  (work_date >= '"&from_date&"' and work_date <= '"&to_date&"') and cost_detail = '"&cost_detail&"'"&over_sql&" order by org_name,user_name,work_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�Ϲݰ��" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,customer as slip_memo,cost FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and slip_gubun = '���' and account = '"&cost_detail&"'"&com_sql&" order by org_name,emp_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "������" or cost_id = "���ֺ�" or cost_id = "����" or cost_id = "���" or cost_id = "��ݺ�" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,customer as slip_memo,cost FROM general_cost where (cancel_yn = 'N') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and slip_gubun = '"&cost_id&"' and account = '"&cost_detail&"'"&com_sql&" order by org_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "���߱���" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,concat(transit,'-',run_memo) as slip_memo,(fare) as cost FROM transit_cost  where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and car_owner = '"&cost_detail&"'"&tran_sql&" order by org_name,user_name, run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "ȸ��" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,concat(company,'-',run_memo) as slip_memo,(somopum+oil_price+parking+toll) as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and car_owner = '"&cost_detail&"'"&tran_sql&" order by org_name,user_name,run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "����" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,concat(company,'-',run_memo) as slip_memo,(oil_price+somopum+parking+toll) as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and car_owner = '"&cost_detail&"'"&tran_sql&" order by org_name,user_name,run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "�����" and cost_detail = "����������" then
	sql = "select org_name,run_date as slip_date,user_name,user_grade,run_memo as slip_memo,repair_cost as cost FROM transit_cost where (cancel_yn = 'N') and (run_date >= '"&from_date&"' and run_date <= '"&to_date&"') and car_owner = 'ȸ��'"&tran_sql&" order by org_name,user_name, run_date asc"
	rs.Open sql, Dbconn, 1
end if

if cost_id = "����ī��" then
	sql = "select org_name,slip_date,emp_name as user_name,emp_grade as user_grade,concat(card_slip.customer,'-',card_slip.account_item) as slip_memo,card_slip.cost FROM card_slip where (card_type not like '%����%' or com_drv_yn = 'Y') and (slip_date >= '"&from_date&"' and slip_date <= '"&to_date&"') and account = '"&cost_detail&"'"&card_sql&" order by org_name,emp_name,slip_date asc"
	rs.Open sql, Dbconn, 1
end if
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
								if rs("cost") <> "0" then
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
								<td class="right">
							<% if cost_id <> "�ΰǺ�" then	%>	
								<%=formatnumber(rs("cost"),0)%>
							<%   else	%>
							<%     if pay_grade = 0	then	%>
								<%=formatnumber(rs("cost"),0)%>
							<%       else	%>
								**********
                            <%	   end if	%>
                            <% end if	%>
                                </td>
							</tr>
							<%
								end if
								rs.movenext()
							loop
							rs.close()
							if pay_sw = "N" then
							%>
							<tr>
								<td colspan="6" class="first">��ȸ ������ �����ϴ�.</td>
							</tr>
							<%
							end if
							%>
							<tr>
								<th colspan="5" class="first">�հ�</th>
								<th class="right">
							<% if cost_id <> "�ΰǺ�" then	%>	
								<%=formatnumber(cost_sum,0)%>
							<%   else	%>
							<%     if pay_grade = 0	then	%>
								<%=formatnumber(cost_sum,0)%>
							<%       else	%>
								**********
                            <%	   end if	%>
                            <% end if	%>
                                </th>
							</tr>
						</tbody>
					</table>
				</div>				        				
	</form>
	</body>
</html>

