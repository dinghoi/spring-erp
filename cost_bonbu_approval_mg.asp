<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_cost = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

cost_month=Request.form("cost_month")
if cost_month = "" then
	be_date = dateadd("m",-1,now())
	be_month = mid(cstr(be_date),1,4) + mid(cstr(be_date),6,2)
	cost_month = be_month
end If

cost_year = mid(cost_month,1,4)
cost_mm = mid(cost_month,5,2)

be_year_year = cstr(int(cost_year) - 1)
be_year_mm = cost_mm

be_mm = int(cost_mm) - 1
be_month_year = cost_year
if be_mm = 0 then
	be_month_mm = 12
	be_month_year = be_year_year
  elseif be_mm > 0 and be_mm < 10 then
	be_month_mm = "0" + cstr(be_mm)
  else
  	be_month_mm = cstr(be_mm)
end if

if position = "본부장" or cost_grade = "1" then
	sql = "select * from emp_org_mst where org_level = '사업부' and org_bonbu = '"&bonbu&"' group by org_name Order By org_name Asc"
  else
	sql = "select * from emp_org_mst where org_level = '사업부' group by org_name Order By org_name Asc"  
end if
if user_id = "100031" then
	sql = "select * from emp_org_mst where org_level = '사업부' and (org_name = 'KAL지원사업부' or org_name = '공항지원사업부') group by org_name Order By org_name Asc"
end if

Rs.Open Sql, Dbconn, 1

title_line = "비용사용 승인 관리"
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
			function getPageCode(){
				return "1 1";
			}
			function frmcheck () {
					document.frm.submit();
			}
			
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_bonbu_approval_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건검색</dt>
                        <dd>
                            <p>
							<label>
							&nbsp;&nbsp;<strong>조회년월&nbsp;</strong> : 
                            <input name="cost_month" type="text" value="<%=cost_month%>" style="width:70px">
							</label>
                            <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="8%" >
							<col width="14%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">본부 / 사업부</th>
								<th rowspan="2" scope="col">전년</th>
								<th rowspan="2" scope="col">전월</th>
								<th rowspan="2" scope="col">당월</th>
								<th colspan="2" style=" border-bottom:1px solid #e3e3e3;" scope="col">전월 증감</th>
								<th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">전년 증감</th>
								<th rowspan="2" scope="col">보고자료</th>
								<th rowspan="2" scope="col">본부장승인</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">증감액</th>
							  <th scope="col">증감율</th>
							  <th scope="col">증감액</th>
							  <th scope="col">증감율</th>
                          </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

							sql="select * from cost_end where saupbu='"&rs("org_name")&"' and end_month ='"&cost_month&"'"
							set rs_cost=dbconn.execute(sql)
							if rs_cost.eof or rs_cost.bof then
								mod_date = ""
								batch_yn = "N"
								bonbu_yn = "N"
								ceo_yn = "N"
							  else														
								batch_yn = rs_cost("batch_yn")
								bonbu_yn = rs_cost("bonbu_yn")
								ceo_yn = rs_cost("ceo_yn")
								if batch_yn = "N" then
									mod_date = ""
								  else	
									mod_date = rs_cost("mod_date")
								end if
							end if					
							if batch_yn = "Y" then
								sql = "select sum(cost_amt_"&cost_mm&") as cost_amt from org_cost where cost_year ='"&cost_year&"' and saupbu ='"&rs("org_name")&"'"
								set rs_cost=dbconn.execute(sql)
								if isnull(rs_cost("cost_amt")) then
									curr_cost = 0
								  else
								  	curr_cost = cdbl(rs_cost("cost_amt"))
								end if
								sql = "select sum(cost_amt_"&cost_mm&") as cost_amt from org_cost where cost_year ='"&be_year_year&"' and saupbu ='"&rs("org_name")&"'"
								set rs_cost=dbconn.execute(sql)
								if isnull(rs_cost("cost_amt")) then
									be_year_cost = 0
								  else
								  	be_year_cost = cdbl(rs_cost("cost_amt"))
								end if
								sql = "select sum(cost_amt_"&be_month_mm&") as cost_amt from org_cost where cost_year ='"&be_month_year&"' and saupbu ='"&rs("org_name")&"'"
								set rs_cost=dbconn.execute(sql)
								if isnull(rs_cost("cost_amt")) then
									be_month_cost = 0
								  else
								  	be_month_cost = cdbl(rs_cost("cost_amt"))
								end if
							  else
								curr_cost = 0
								be_year_cost = 0
								be_month_cost = 0
							end if
							month_cr = curr_cost - be_month_cost
							year_cr = curr_cost - be_year_cost
							if curr_cost = 0 or be_month_cost = 0 then
								month_per = 0
							  else
							  	month_per = month_cr / be_month_cost * 100
							end if
							if curr_cost = 0 or be_year_cost = 0 then
								year_per = 0
							  else
							  	year_per = year_cr / be_year_cost * 100
							end if
						%>
							<tr>
								<td class="first"><%=rs("org_bonbu")%>&nbsp;/&nbsp;<%=rs("org_name")%></td>
								<td class="right"><%=formatnumber(be_year_cost,0)%></td>
								<td class="right"><%=formatnumber(be_month_cost,0)%></td>
								<td class="right"><%=formatnumber(curr_cost,0)%></td>
								<td class="right"><%=formatnumber(month_cr,0)%></td>
							  	<td class="right"><%=formatnumber(month_per,2)%>%</td>
								<td class="right"><%=formatnumber(year_cr,0)%></td>
								<td class="right"><%=formatnumber(year_per,2)%>%</td>
								<td><%=mod_date%>&nbsp;</td>
								<td>
						<% if batch_yn = "Y" and bonbu_yn = "N" then	%>
                                <a href="#" onClick="pop_Window('cost_approval_view.asp?saupbu=<%=rs("org_name")%>&cost_month=<%=cost_month%>','cost_approval_view_pop','scrollbars=yes,width=1250,height=600')" class="btnType04">
                                승인처리</a>
						<%   else	%>
                        		&nbsp;
						<% end if	%>
						<% if bonbu_yn = "Y" then	%>
                                <a href="#" onClick="pop_Window('cost_approval_view.asp?saupbu=<%=rs("org_name")%>&cost_month=<%=cost_month%>','cost_approval_view_pop','scrollbars=yes,width=1250,height=600')" class="btnType04">
                                완료</a>
						<% end if	%>
                                </td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

