<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab(20)
dim grade_tab(20)
dim grade_pay(20,20)
dim sum_pay(20)
dim month_tab(24,2)

be_pg = "ceo_pay_grade_org_report_jm.asp"

curr_dd = cstr(datepart("d",now))
to_date = mid(cstr(now()),1,10)
from_date = mid(cstr(now()-curr_dd+1),1,10)

view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
	pmg_yymm=Request.form("pmg_yymm")
	pmg_yymm_end=Request.form("pmg_yymm_end")
  else
	view_condi = request("view_condi")
	condi = request("condi") 
	pmg_yymm=request("pmg_yymm") 
	pmg_yymm_end=request("pmg_yymm_end") 
end if

if view_condi = "" then
	view_condi = "전체"
	condi = "전체"
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	'pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
	pmg_yymm_end = pmg_yymm
end if

'response.write(view_condi)
'response.write(company)

for i = 1 to 20
    com_tab(i) = ""
	grade_tab(i) = ""
next

for i = 1 to 20
    for j = 1 to 20
	    grade_pay(i,j) = 0
    next
	sum_pay(i) = 0
next

cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
'cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

' 직급테이블에 직급명칭 가져오기
Sql="select * from emp_etc_code where emp_etc_type = '02' order by emp_etc_code DESC"
Rs_etc.Open Sql, Dbconn, 1
k = 0
while not Rs_etc.eof
	k = k + 1
	grade_tab(k) = Rs_etc("emp_etc_name")
	Rs_etc.movenext()
Wend
Rs_etc.close()	

' 회사테이블에 회사 또는 본부명칭 가져오기
if view_condi = "전체" then
	' 2019.02.22 박정신 요청 회사리스트를 빼고자 할시 org_end_date에 null 이 아닌 만료일자를 셋팅하면 리스트에 나타나지 않는다.
	Sql = "SELECT * FROM emp_org_mst WHERE ISNULL(org_end_date) AND org_level = '회사'  ORDER BY org_company ASC"
   Rs_org.Open Sql, Dbconn, 1
   k = 0
   while not Rs_org.eof
	   k = k + 1
	   com_tab(k) = Rs_org("org_name")
	   Rs_org.movenext()
   Wend
 elseif condi = "전체" then
           Sql="select * from emp_org_mst where (org_level = '본부') and (org_company='"+view_condi+"') order by org_code ASC"
           Rs_org.Open Sql, Dbconn, 1
           k = 0
           while not Rs_org.eof
	             k = k + 1
	             com_tab(k) = Rs_org("org_name")
	            Rs_org.movenext()
           Wend   
		else 
		   Sql="select * from emp_org_mst where (org_level = '사업부') and (org_company='"+view_condi+"') and (org_bonbu='"+condi+"') order by org_code ASC"
           Rs_org.Open Sql, Dbconn, 1
           k = 0
           while not Rs_org.eof
	             k = k + 1
	             com_tab(k) = Rs_org("org_name")
	            Rs_org.movenext()
           Wend   
end if
Rs_org.close()
k_org = k	

if view_condi = "전체" then
   Sql = "select * from pay_month_give where (pmg_yymm between '"+pmg_yymm+"' and '" & pmg_yymm_end &"' ) and (pmg_id = '1')"
elseif condi = "전체" then  
			Sql = "select * from pay_month_give where (pmg_yymm between '"+pmg_yymm+"' and '" & pmg_yymm_end &"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"')"
		  else 
			Sql = "select * from pay_month_give where (pmg_yymm between '"+pmg_yymm+"' and '" & pmg_yymm_end &"' ) and (pmg_id = '1') and (pmg_company = '"+view_condi+"') and (pmg_bonbu='"+condi+"')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof 
     
	pmg_give_tot = rs("pmg_give_total") 
    
    Sql = "select * from pay_month_deduct where (de_yymm between '"+pmg_yymm+"' and '" & pmg_yymm_end & "' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"') and (de_company = '"+view_condi+"')"
    Set Rs_dct = DbConn.Execute(SQL)
	if not Rs_dct.eof then
			de_deduct_tot = Rs_dct("de_deduct_total")
	   else
			de_deduct_tot = 0
    end if
    Rs_dct.close()
	pmg_curr_pay = pmg_give_tot - de_deduct_tot

   if view_condi = "전체" then
      com_name = rs("pmg_company")
      elseif condi = "전체" then 
                com_name = rs("pmg_bonbu")
			 else
			    com_name = rs("pmg_saupbu")
   end if

   k = 0                                       
   for i = 1 to k_org
       if com_tab(i) = com_name then
          k = i
	   end if
    next
	
    if k = 0 then   
	   k = k_org + 1
	   if condi = "전체" then 
	          com_tab(k) = view_condi
		  else
		      com_tab(k) = condi
	   end if
	 end if
	 
    j = 0
    for i = 0 to 20
       if grade_tab(i) = rs("pmg_grade") then
	      j = i
	   end if
    next
	
	if j = 0 then   
	   j = 1
	 end if
	
	grade_pay(k,j) = grade_pay(k,j) + pmg_curr_pay
	sum_pay(j) = sum_pay(j) + pmg_curr_pay
	
    rs.movenext()
loop
rs.close()

tot_pay = 0
for j = 1 to 14 
    tot_pay = tot_pay + sum_pay(j)
next


title_line = ""+ view_condi +" - 직급별 급여분포 현황 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>임원 정보 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.view_condi.value =="회사별") {
					if(document.frm.condi.value =="전체") {
						alert('회사를 선택하세요');
						frm.condi.focus();
						return false;}}		
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
            <!--#include virtual = "/include/ceo_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<% '<form action="waiting.asp?pg_name=insa_grade_count.asp" method="post" name="frm"> %>
                <form action="<%=be_pg%>?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                               <label>
                               <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where (isNull(org_end_date) or org_end_date = '1900-01-01') and (org_level = '회사') ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
								<select name="view_condi" id="view_condi" type="text" style="width:150px">
                                  <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
                                </label>
                                <label>
            					</select>
								<strong>조건 : </strong>
                              <%
								Sql="select * from emp_org_mst where isNull(org_end_date) and org_level = '본부' and org_company = '"+view_condi+"' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
								<select name="condi" id="condi" type="text" style="width:150px">
                                  <option value="전체" <%If condi = "전체" then %>selected<% end if %>>전체</option>
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If condi = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
                			  <%
									rs_org.movenext()  
								loop 
								rs_org.Close()
							  %>
            					</select>  
                                </label>
                                <label>
								<strong>귀속년월 : </strong>
                                    <select name="pmg_yymm" id="pmg_yymm" type="text" value="<%=pmg_yymm%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								 ~
                                    <select name="pmg_yymm_end" id="pmg_yymm_end" type="text" value="<%=pmg_yymm_end%>" style="width:90px">
                                    <%	for i = 24 to 1 step -1	%>
                                    <option value="<%=month_tab(i,1)%>" <%If pmg_yymm_end = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                <div  style="text-align:right">
				<strong>금액단위 : 만원</strong>
				</div>
				<div class="gView">
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
                            <col width="6%" >
                            <% for i = 1 to 14 %>
							       <col width="5%" >
                            <% next	%>
                            <col width="6%" >
						</colgroup>
                        <thead>    
                            <tr>
								<th class="first" scope="col">조&nbsp;&nbsp;&nbsp;직</th>
                                <th scope="col">분포율</th>
                                <% 
								for i = 1 to 14 
								%>
                                <th scope="col"><%=grade_tab(i)%></th>
								<%
								next
								%>
                                <th scope="col" style=" border-left:1px solid #e3e3e3;">소계</th>
							</tr>
						</thead>
						<tbody>
                        <%
                        for i = 0 to 20 
                        	if	com_tab(i) <> "" then
						%>	
                            <tr>
                                <% 
								hap_pay = 0
								for j = 1 to 14 
								    hap_pay = hap_pay + grade_pay(i,j)
								next
								
								if tot_pay = 0 then
								      cr_pro = 0
								   else
								      cr_pro = (hap_pay / tot_pay) * 100
								end if
					
								%>
                                <td><%=com_tab(i)%></td>
                                <td><%=formatnumber(cr_pro,2)%>%</td>
                                <% 
								for j = 1 to 14 
								    'ost_amt = cdbl(cost_amt) / 10000 금액단위 짜르는것
								%>
                                    <td class="right"><%=formatnumber(grade_pay(i,j)/10000,0)%></td>
								<%
								next
								%>
                                <td class="right"><%=formatnumber(hap_pay/10000,0)%></td> 
                             </tr>
                        <%
							end if
						next
						
						'증감율퍼센트 
						'cr_cost = curr_cost - be_cost							
						'if cr_cost = 0 then
						'	cr_pro = 0
						'  elseif bi_cost = 0 then
						'  	cr_pro = 100 
						'  else
						'  	cr_pro = cr_cost / be_cost
						'end if
                        %>
							<tr>
                              <th colspan="2">총계</th>
                              <% 
								hap_pay = 0
								for j = 1 to 14 
								    hap_pay = hap_pay + sum_pay(j)
								    'ost_amt = cdbl(cost_amt) / 10000 금액단위 짜르는것
								%>
                                    <th class="right"><%=formatnumber(sum_pay(j)/10000,0)%></th>
								<%
								next
								%>
                                <th class="right"><%=formatnumber(hap_pay/10000,0)%></th> 
							</tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

