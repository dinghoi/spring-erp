<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
saupbu = Request("saupbu")
bonbu = Request("bonbu")
emp_company = Request("emp_company")
cost_id = Request("cost_id")
cost_detail = Request("cost_detail")
cost_year = Request("cost_year")
cost_month = int(Request("cost_month"))
be_month = cost_month - 1
if be_month < 10 then
	be_month = "0" + cstr(be_month)
end if
if cost_month < 10 then
	cost_month = "0" + cstr(cost_month)
end if
if be_month = "00" then
	be_month = "12"
	be_year = int(cost_year) - 1
end if

from_date = cstr(cost_year) + "-" + cstr(cost_month) + "-01"
end_date = datevalue(from_date)
end_date = dateadd("m",1,from_date)
to_date = cstr(dateadd("d",-1,end_date))

if cost_month = "01" then
	sql = "select team,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	if emp_company <> "" then
		sql = "select team,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&emp_company&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
	if saupbu <> "" then
		sql = "select team,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
	if saupbu = "" then
		sql = "select team,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
	if bonbu <> "" then
		sql = "select team,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and bonbu ='"&bonbu&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
  else
	sql = "select team,sum(cost_amt_"&be_month&") as be_cost,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	if emp_company <> "" then
		sql = "select team,sum(cost_amt_"&be_month&") as be_cost,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and emp_company ='"&emp_company&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
	if saupbu <> "" then
		sql = "select team,sum(cost_amt_"&be_month&") as be_cost,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
	if saupbu = "" then
		sql = "select team,sum(cost_amt_"&be_month&") as be_cost,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&saupbu&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
	if bonbu <> "" then
		sql = "select team,sum(cost_amt_"&be_month&") as be_cost,sum(cost_amt_"&cost_month&") as cost from org_cost where cost_year ='"&cost_year&"' and bonbu ='"&bonbu&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team order by team"
	end if
end if  
'response.write(sql)
rs.Open sql, Dbconn, 1

title_line = saupbu + " 팀별 비용 현황"

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
							<strong>년월 : </strong><%=cost_year%>년<%=cost_month%>월&nbsp;
							<strong>비용구분 : </strong><%=cost_id%>-<%=cost_detail%>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="30%" >
							<col width="30%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">조직</th>
								<th scope="col">전월금액</th>
								<th scope="col">당월금액</th>
							</tr>
						</thead>
						<tbody>
         					<% 
							tot_be_cost = 0
							tot_cost = 0
							do until rs.eof
								if be_month = "12" then
									sql = "select team,sum(cost_amt_12) as cost from org_cost where cost_year ='"&be_year&"' and saupbu ='"&saupbu&"' and team ='"&rs("team")&"' and cost_id ='"&cost_id&"' and cost_detail ='"&cost_detail&"' group by team"
									set rs_be=dbconn.execute(sql)
									if rs_be.eof or rs_be.bof then
										be_cost = 0
									  else									
										be_cost = rs_be("cost")
									end if
								  else
								  	be_cost = rs("be_cost")
								end if
								tot_be_cost = tot_be_cost + cdbl(be_cost)
								tot_cost = tot_cost + cdbl(rs("cost"))
							%>
							<tr>
								<td class="first"><%=saupbu%>&nbsp;<%=rs("team")%></td>
								<td class="right">
							<% if cost_id <> "인건비" or pay_grade = "0" then	%>
								<%=formatnumber(be_cost,0)%>
							<%   else	%>
                            	**********
							<% end if	%>
                                </td>
								<td class="right"><a href="#" onClick="pop_Window('person_cost_view.asp?cost_year=<%=cost_year%>&cost_month=<%=cost_month%>&cost_id=<%=cost_id%>&cost_detail=<%=cost_detail%>&emp_company=<%=emp_company%>&bonbu=<%=bonbu%>&saupbu=<%=saupbu%>&team=<%=rs("team")%>','person_cost_view_pop','scrollbars=yes,width=800,height=500')">
							<% if cost_id <> "인건비" or pay_grade = "0" then	%>
								<%=formatnumber(rs("cost"),0)%>
							<%   else	%>
                            	**********
							<% end if	%>
                                </a></td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
							<tr>
							  <th class="first">계</th>
							  <th class="right"><%=formatnumber(tot_be_cost,0)%></th>
							  <th class="right"><%=formatnumber(tot_cost,0)%></th>
				          </tr>
						</tbody>
					</table>
				</div>				
	</form>
	</body>
</html>

