<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim com_tab(20)
dim grade_tab(20)
dim grade_cnt(20,20)
dim sum_cnt(20)

curr_dd = cstr(datepart("d",now))
to_date = mid(cstr(now()),1,10)
from_date = mid(cstr(now()-curr_dd+1),1,10)
view_condi = request("view_condi")
condi = request("condi")  

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	condi = request.form("condi")
  else
	view_condi = request("view_condi")
	condi = request("condi")  
end if

if view_condi = "" then
	view_condi = "전체"
	condi = "전체"
end if

'response.write(view_condi)
'response.write(company)

for i = 0 to 20
    com_tab(i) = ""
	grade_tab(i) = ""
next

for i = 0 to 20
    for j = 0 to 20
	    grade_cnt(i,j) = 0
    next
	sum_cnt(i) = 0
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_as = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
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
 else
   Sql="select * from emp_org_mst where (org_level = '본부') and (org_company = '"+condi+"') order by org_name ASC"
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

'전체건수대비 율을 계산하기위한....
'sql = " select count(*) as tot_cnt from as_acpt Where (mg_group='"+mg_group+"') and (Cast(acpt_date as date) >= '" + from_date + "' and 'Cast(acpt_date as date) <= '"+to_date+"')"
'Set rs=DbConn.Execute(SQL)
'tot_cnt = cint(rs("tot_cnt"))
'if tot_cnt = "" or isnull(tot_cnt) then
'	tot_cnt = 0
'end if
'rs.close()
' 화면뿌려줄때...%=formatnumber(com_cnt(0)/tot_cnt*100,2)

'if view_condi = "전체" then
'   sql = " select emp_company, emp_grade, count(*) from emp_master group by emp_company, emp_grade where isNull(emp_end_date) or emp_end_date = '1900-01-01'"
'   Rs.Open Sql, Dbconn, 1
'   else
'   sql = " select emp_bonbu, emp_grade, count(*) from emp_master group by emp_bonbu, emp_grade where (emp_company = '"+company+"') and isNull(emp_end_date) or emp_end_date = '1900-01-01'"
'   Rs.Open Sql, Dbconn, 1
'end if

Sql = "SELECT count(*) FROM emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

'response.write(tottal_record)

if view_condi = "전체" then
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
   else  
   Sql = "select * from emp_master where (emp_company='"+condi+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
end if
Rs.Open Sql, Dbconn, 1

do until rs.eof 
   if view_condi = "전체" then
      com_name = rs("emp_company")
      else
      com_name = rs("emp_bonbu")
   end if
   k = 0                                       
   for i = 1 to k_org
       if com_tab(i) = com_name then
          k = i
	   end if
    next
	
    if k = 0 then   '임시로... 데이타가 잘못되어 비교가 안됨
	   k = k_org + 1
	   com_tab(k) = condi
	 end if
	 
    j = 0
    for i = 0 to 20
       if grade_tab(i) = rs("emp_grade") then
	      j = i
	   end if
    next
	
	if j = 0 then   '임시로... 데이타가 잘못되어 비교가 안됨
	   j = 1
	 end if
	
	grade_cnt(k,j) = grade_cnt(k,j) + 1
	sum_cnt(j) = sum_cnt(j) + 1
	
    rs.movenext()
loop
rs.close()


title_line = "조직별 직급별 인원현황"
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
				return "3 1";
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
				if (chkfrm()) {
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/ceo_header.asp" -->
			<!--#include virtual = "/include/ceo_insa_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<% '<form action="waiting.asp?pg_name=insa_grade_count.asp" method="post" name="frm"> %>
                <form action="ceo_insa_grade_count.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>선택</strong>
                                <select name="view_condi" id="view_condi" style="width:100px">
                                    <option value="전체" <%If view_condi = "전체" then %>selected<% end if %>>전체</option>
                                    <option value="회사별" <%If view_condi = "회사별" then %>selected<% end if %>>회사별</option>
                                </select>
								</label>
								<strong>회사</strong>
							  	<%
									Sql="select * from emp_org_mst where (org_level = '회사') ORDER BY org_code ASC"
                                    rs_org.Open Sql, Dbconn, 1
                                %>
								<label>
        						<select name="condi" id="condi" type="text" style="width:150px" value="<%=condi%>">
                                    <option value="전체" <%If condi = "전체" then %>selected<% end if %>>전체</option>
          					<% 
								While not rs_org.eof 
							%>
          							<option value='<%=rs_org("org_name")%>' <%If rs_org("org_name") = condi  then %>selected<% end if %>><%=rs_org("org_name")%></option>
          					<%
									rs_org.movenext()  
								Wend 
								rs_org.Close()
							%>
        						</select>
                                </label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">회사/본부</th>
								<th scope="col"><%=grade_tab(1)%></th>
								<th scope="col"><%=grade_tab(2)%></th>
								<th scope="col"><%=grade_tab(3)%></th>
								<th scope="col"><%=grade_tab(4)%></th>
								<th scope="col"><%=grade_tab(5)%></th>
								<th scope="col"><%=grade_tab(6)%></th>
								<th scope="col"><%=grade_tab(7)%></th>
								<th scope="col"><%=grade_tab(8)%></th>
								<th scope="col"><%=grade_tab(9)%></th>
								<th scope="col"><%=grade_tab(10)%></th>
								<th scope="col"><%=grade_tab(11)%></th>
								<th scope="col"><%=grade_tab(12)%></th>
								<th scope="col"><%=grade_tab(13)%></th>
								<th scope="col"><%=grade_tab(14)%></th>
                                <th scope="col" style=" border-left:1px solid #e3e3e3;">소계</th>
							</tr>
						</thead>
						<tbody>
                        <%
                        for i = 0 to 20 
                        	if	com_tab(i) <> "" then
						%>	
                            <tr>
                                <td><%=com_tab(i)%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(1)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,1),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(2)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,2),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(3)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,3),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(4)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,4),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(5)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,5),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(6)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,6),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(7)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,7),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(8)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,8),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(9)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,9),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(10)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,10),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(11)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,11),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(12)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,12),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(13)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,13),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_grade_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&grade=<%=grade_tab(14)%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(grade_cnt(i,14),0)%></a>
                                </td>
                                <td><%=formatnumber(clng(grade_cnt(i,1)+grade_cnt(i,2)+grade_cnt(i,3)+grade_cnt(i,4)+grade_cnt(i,5)+grade_cnt(i,6)+grade_cnt(i,7)+grade_cnt(i,8)+grade_cnt(i,9)+grade_cnt(i,10)+grade_cnt(i,11)+grade_cnt(i,12)+grade_cnt(i,13)+grade_cnt(i,14)),0)%>&nbsp;</td>
                             </tr>
                        <%
							end if
						next
                        %>
							<tr>
                              <th>총계</th>
                              <th><%=formatnumber(sum_cnt(1),0)%></th>
                              <th><%=formatnumber(sum_cnt(2),0)%></th>
                              <th><%=formatnumber(sum_cnt(3),0)%></th>
                              <th><%=formatnumber(sum_cnt(4),0)%></th>
                              <th><%=formatnumber(sum_cnt(5),0)%></th>
                              <th><%=formatnumber(sum_cnt(6),0)%></th>
                              <th><%=formatnumber(sum_cnt(7),0)%></th>
                              <th><%=formatnumber(sum_cnt(8),0)%></th>
                              <th><%=formatnumber(sum_cnt(9),0)%></th>
                              <th><%=formatnumber(sum_cnt(10),0)%></th>
                              <th><%=formatnumber(sum_cnt(11),0)%></th>
                              <th><%=formatnumber(sum_cnt(12),0)%></th>
                              <th><%=formatnumber(sum_cnt(13),0)%></th>
                              <th><%=formatnumber(sum_cnt(14),0)%></th>
                              <th><%=formatnumber(clng(sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)+sum_cnt(10)+sum_cnt(11)+sum_cnt(12)+sum_cnt(13)+sum_cnt(14)),0)%>&nbsp;</th>
							</tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

