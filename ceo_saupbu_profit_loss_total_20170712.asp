<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
dim year_tab(5)
dim sum_amt(20,3,13)
dim saupbu_tab(20)

cost_year=Request.form("cost_year")

if cost_year = "" then
	cost_year = mid(cstr(now()),1,4)
	base_year = cost_year
	view_sw = "0"
end If

be_year = int(cost_year) - 1
for i = 1 to 5
	year_tab(i) = int(cost_year) - i + 1
next

for i = 1 to 20
	saupbu_tab(i) = ""
next

for i = 1 to 20
	for j = 1 to 3
		for k = 1 to 13
			sum_amt(i,j,k) = 0
		next
	next
next

' �������� ����
sql = "select saupbu from sales_org order by sort_seq"
rs.Open sql, Dbconn, 1
i = 0
do until rs.eof
	i = i + 1
	saupbu_tab(i) = rs("saupbu")
	rs.movenext()
loop
rs.close()
i = i + 1 
saupbu_tab(i) = "��Ÿ�����"
i = i + 1 
saupbu_tab(i) = "ȸ�簣�ŷ�"

' ȸ�簣�ŷ�
sql = "select cost_center,sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from company_cost where cost_year = '"&cost_year&"' and (cost_center = 'ȸ�簣�ŷ�') group by cost_center"
rs.Open sql, Dbconn, 1
do until rs.eof
	for k = 1 to 12
		sum_amt(i,2,k) = sum_amt(i,2,k) + cdbl(rs(k))
	next
	rs.movenext()
loop
rs.close()


' ���� ����
sql = "select substring(sales_date,1,7) as sales_month,saupbu,sum(cost_amt) as cost from saupbu_sales where substring(sales_date,1,4) = '"&cost_year&"' group by substring(sales_date,1,7), saupbu"
rs.Open sql, Dbconn, 1
do until rs.eof
	for i = 1 to 20
		if saupbu_tab(i) = rs("saupbu") then
			j = 1
			k = int(mid(rs("sales_month"),6,2))
			sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs("cost"))
			exit for
		end if
	next			
	rs.movenext()
loop
rs.close()

' ��� ����
sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' group by saupbu"
rs.Open sql, Dbconn, 1
do until rs.eof
	for i = 1 to 20
		if saupbu_tab(i) = rs("saupbu") then
			j = 2
			for k = 1 to 12
				sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs(k))
			next
			exit for
		end if
	next			
	rs.movenext()
loop
rs.close()

' ��� ���� (��Ÿ�����)
sql = "select saupbu, sum(cost_amt_01), sum(cost_amt_02), sum(cost_amt_03), sum(cost_amt_04), sum(cost_amt_05), sum(cost_amt_06), sum(cost_amt_07), sum(cost_amt_08), sum(cost_amt_09), sum(cost_amt_10), sum(cost_amt_11), sum(cost_amt_12) from saupbu_profit_loss where cost_year = '"&cost_year&"' and (saupbu = '' or saupbu = '��Ÿ�����') group by saupbu"
rs.Open sql, Dbconn, 1
do until rs.eof
	cost_saupbu = rs("saupbu")
	if cost_saupbu = "" then
		cost_saupbu = "��Ÿ�����"
	end if
	for i = 1 to 20
		if saupbu_tab(i) = cost_saupbu then
			j = 2
			for k = 1 to 12
				sum_amt(i,j,k) = sum_amt(i,j,k) + cdbl(rs(k))
			next
			exit for
		end if 
	next			
	rs.movenext()
loop
rs.close()

' ����� ������ ���⵵ ǥ�� ���� ����
'for i = 1 to 20
'	if saupbu_tab(i) = "" then
'		exit for
'	end if
'	for k = 1 to 12
'		if sum_amt(i,2,k) = 0 then
'			sum_amt(i,1,k) = 0
'		end if
'	next
'next			

' ���Ͱ��
for i = 1 to 20
	if saupbu_tab(i) = "" then
		exit for
	end if
	j = 3
	for k = 1 to 12
		sum_amt(i,j,k) = sum_amt(i,1,k) - sum_amt(i,2,k)
	next
next			

' �� �հ�
for i = 1 to 20
	if saupbu_tab(i) = "" then
		exit for
	end if
	for j = 1 to 3
		for k = 1 to 12
			sum_amt(i,j,13) = sum_amt(i,j,13) + sum_amt(i,j,k)
		next
	next
next			

' �Ѱ�
for i = 1 to 20
	if saupbu_tab(i) = "" then
		exit for
	end if
	for j = 1 to 3
		for k = 1 to 13
			sum_amt(0,j,k) = sum_amt(0,j,k) + sum_amt(i,j,k)
		next
	next
next			

title_line = "����κ� ���� �Ѱ� ��Ȳ"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.cost_year.value == "") {
					alert ("��ȸ���� �Է��ϼ���.");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
			<!--#include virtual = "/include/ceo_cost_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="ceo_saupbu_profit_loss_total.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
							<label>
							&nbsp;&nbsp;<strong>��ȸ��&nbsp;</strong> : 
                            <select name="cost_year" id="cost_year" style="width:70px">
							<% for i = 1 to 5 %>
                              <option value="<%=year_tab(i)%>" <% if cost_year=year_tab(i) then %>selected<% end if %>>&nbsp;<%=year_tab(i)%></option>
							<% next	%>
							</select>
							</label>
                            <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div  style="text-align:right">
				<strong>�ݾ״��� : õ��</strong>
				</div>
                <div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
						</colgroup>
						<thead>
							<tr>
							  <th class="first" scope="col">�����</th>
							  <th scope="col">����</th>
						<% for i = 1 to 12	%>
							  <th scope="col"><%=i%>��</th>
						<% next	%>
							  <th scope="col">�հ�</th>
                          </tr>
						</thead>
						<tbody>
					<%
						for i = 1 to 20
							if saupbu_tab(i) = "" then
								exit for
							end if
					%>							
							<tr>
							  	<td rowspan="3" class="first"><%=saupbu_tab(i)%></td>
								<td>����</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(i,1,k)/1000,0)%></td>
						<%	
							next	
						%>								
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">���</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right">
						<% 	if (k < 13 and sum_amt(i,2,k) > 0) and (saupbu_tab(i) <> "ȸ�簣�ŷ�") then	%>
								<a href="#" onClick="pop_Window('saupbu_profit_loss_report.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>&sales_saupbu=<%=saupbu_tab(i)%>','saupbu_profit_loss_report_pop','scrollbars=yes,width=1230,height=650')"><%=formatnumber(sum_amt(i,2,k)/1000,0)%></a>
						<% 	  else	%>
						<% 		if (k < 13 and sum_amt(i,2,k) > 0) and (saupbu_tab(i) = "ȸ�簣�ŷ�") then	%>
								<a href="#" onClick="pop_Window('company_deal_detail_view.asp?cost_year=<%=cost_year%>&cost_mm=<%=k%>','company_deal_detail_view_pop','scrollbars=yes,width=1000,height=600')"><%=formatnumber(sum_amt(i,2,k)/1000,0)%></a>
						<% 	  		else	%>
								<%=formatnumber(sum_amt(i,2,k)/1000,0)%>
                        <%		end if	%>
                        <%	end if	%>
                                </td>
						<%	
							next	
						%>								
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">����</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(i,3,k)/1000,0)%></td>
						<%	
							next	
						%>								
			              </tr>
					<%
						next
					%>
							<tr>
							  	<td rowspan="3" class="first" bgcolor="#CCFFFF"><strong>��</strong></td>
								<td>����</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(0,1,k)/1000,0)%></td>
						<%	
							next	
						%>								
							</tr>
							<tr>
							  <td style="border-left:1px solid #e3e3e3;">���</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(0,2,k)/1000,0)%></td>
						<%	
							next	
						%>								
			              	</tr>
							<tr bgcolor="#FFDFDF">
							  <td style="border-left:1px solid #e3e3e3;">����</td>
						<%	
							for k = 1 to 13
						%>
								<td class="right"><%=formatnumber(sum_amt(0,3,k)/1000,0)%></td>
						<%	
							next	
						%>								
			              </tr>
						</tbody>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="25%">
                    </td>
				    <td width="50%">
                    </td>
				    <td width="25%">
                    </td>
			      </tr>
				  </table>
				<br>
			</form>
		</div>				
	</div>        				
	</body>
</html>

