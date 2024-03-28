<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim com_tab(20)
dim gun_cnt(20,20)
dim sum_cnt(20)
dim acpt_per(20)
dim acpt_pro(20) 
dim per_cnt

be_pg = "insa_report_gunsok_org.asp"

Page=Request("page")
view_condi = request("view_condi")
condi = request("condi")  

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)
target_date = curr_date

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

for i = 0 to 20
    com_tab(i) = ""
next

for i = 0 to 20
    for j = 0 to 20
	    gun_cnt(i,j) = 0
    next
	sum_cnt(i) = 0
	acpt_per(i) = 0
	acpt_pro(i) = 0
next

per_cnt = 200

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_tab = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

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

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

if view_condi = "전체" then
   Sql = "select * from emp_master where (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
   elseif condi = "전체" then  
            Sql = "select * from emp_master where (emp_company='"+view_condi+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
		  else 
		    Sql = "select * from emp_master where (emp_company='"+view_condi+"') and (emp_bonbu='"+condi+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000')"
end if
Rs.Open Sql, Dbconn, 1


do until rs.eof 
   if view_condi = "전체" then
      com_name = rs("emp_company")
      elseif condi = "전체" then 
                com_name = rs("emp_bonbu")
			 else
			    com_name = rs("emp_saupbu")
   end if
   
   if Rs("emp_first_date") = "1900-01-01" then
      emp_first_date = ""
      else 
      emp_first_date = Rs("emp_first_date")
   end if
                
   if emp_first_date <> "" then 
      year_cnt = datediff("yyyy", Rs("emp_first_date"), target_date)
      mon_cnt = datediff("m", Rs("emp_first_date"), target_date)
      day_cnt = datediff("d", Rs("emp_first_date"), target_date) 
      else 
      year_cnt = datediff("yyyy", Rs("emp_first_date"), target_date)
      mon_cnt = datediff("m", Rs("emp_first_date"), target_date)
      day_cnt = datediff("d", Rs("emp_first_date"), target_date) 
   end if

   target_cnt = cint(year_cnt)
   j = target_cnt
   
   'response.write(target_cnt)

   k = 0                                       
   for i = 1 to k_org
       if com_tab(i) = com_name then
          k = i
	   end if
    next
	
    if k = 0 then   '임시로... 데이타가 잘못되어 비교가 안됨
	   k = k_org + 1
	   if condi = "전체" then 
	          com_tab(k) = view_condi
		  else
		      com_tab(k) = condi
	   end if
	 end if
	
	if j = 0 then   '임시로... 데이타가 잘못되어 비교가 안됨
	   j = 1
	   else
	   j = target_cnt + 1
	end if
	
	if j > 16 then
	   j = 17
	end if
	
	gun_cnt(k,j) = gun_cnt(k,j) + 1
	sum_cnt(j) = sum_cnt(j) + 1
	
    rs.movenext()
loop
rs.close()

for i = 1 to 20
	acpt_per(i) = sum_cnt(i) / per_cnt * 100
next

title_line = ""+ view_condi +" - 근속 현황 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
        <script type="text/javascript" id='dummy'></script>
		<script type="text/javascript">
			function getPageCode(){
				return "5 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
			}			
			
			function selchk() {
				var fm = document.frm;
				opt1= fm.professor.options[fm.professor.selectedIndex].value;
				var scpt= document.getElementById('dummy');
				scpt.src='/table_add_write_select.asp?opt1='+opt1;
			}
			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_report_gunsok_org.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
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
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
                            <col width="5%" >
                            <col width="5%" >
						</colgroup>
						<thead>
							<tr>
				                <th rowspan="2" class="first" scope="col">조직</th>
                                <th colspan="17" scope="col" style=" border-bottom:1px solid #e3e3e3;">근속년수(년)</th>
				                <th rowspan="2" scope="col">계</th>
			                </tr>
                            <tr>
								<th class="first" scope="col" style=" border-left:1px solid #e3e3e3;">1년<br>미만</th>
								<th scope="col">1</th>
								<th scope="col">2</th>
								<th scope="col">3</th>
								<th scope="col">4</th>
                                <th scope="col">5</th>
                                <th scope="col">6</th>
                                <th scope="col">7</th>
                                <th scope="col">8</th>
                                <th scope="col">9</th>
                                <th scope="col">10</th>
                                <th scope="col">11</th>
                                <th scope="col">12</th>
                                <th scope="col">13</th>
                                <th scope="col">14</th>
                                <th scope="col">15</th>
                                <th scope="col">16년<br>이상</th>
							</tr>
						</thead>
						<tbody>
                        <%
                        for i = 0 to 20 
                        	if	com_tab(i) <> "" then
						%>	
                            <tr>
                                <td><%=com_tab(i)%></td>
                                <td><%=formatnumber(gun_cnt(i,1),0)%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="1"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,2),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="2"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,3),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="3"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,4),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="4"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,5),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="5"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,6),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="6"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,7),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="7"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,8),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="8"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,9),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="9"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,10),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="10"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,11),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="11"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,12),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="12"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,13),0)%></a>
                                </td>
                                <td>
                                 <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="13"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,14),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="14"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,15),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="15"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,16),0)%></a>
                                </td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_gunsok_count_view.asp?emp_company=<%=view_condi%>&condi=<%=condi%>&emp_bonbu=<%=com_tab(i)%>&gunsok_yy=<%="16"%>','insa_gunsok_count_pop','scrollbars=yes,width=890,height=600')"><%=formatnumber(gun_cnt(i,17),0)%></a>
                                </td>
                                <td><%=formatnumber(clng(gun_cnt(i,1)+gun_cnt(i,2)+gun_cnt(i,3)+gun_cnt(i,4)+gun_cnt(i,5)+gun_cnt(i,6)+gun_cnt(i,7)+gun_cnt(i,8)+gun_cnt(i,9)+gun_cnt(i,10)+gun_cnt(i,11)+gun_cnt(i,12)+gun_cnt(i,13)+gun_cnt(i,14)+gun_cnt(i,15)+gun_cnt(i,16)+gun_cnt(i,17)+gun_cnt(i,18)),0)%>&nbsp;
                                </td>
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
                              <th><%=formatnumber(sum_cnt(15),0)%></th>
                              <th><%=formatnumber(sum_cnt(16),0)%></th>
                              <th><%=formatnumber(sum_cnt(17),0)%></th>
                              <th><%=formatnumber(clng(sum_cnt(1)+sum_cnt(2)+sum_cnt(3)+sum_cnt(4)+sum_cnt(5)+sum_cnt(6)+sum_cnt(7)+sum_cnt(8)+sum_cnt(9)+sum_cnt(10)+sum_cnt(11)+sum_cnt(12)+sum_cnt(13)+sum_cnt(14)+sum_cnt(15)+sum_cnt(16)+sum_cnt(17)+sum_cnt(18)),0)%>&nbsp;</th>
							</tr>
                            <tr valign="bottom">
                                <td class="first" height="200" valign="middle" style="background:#CFF"><strong>0<br>~<br><%=per_cnt%><br>기준</strong></td>
                  				<% 
								for i = 0 to 20 
									acpt_pro(i) = int(acpt_per(i)*200/100)
								next
								%>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(1)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(2)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(3)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(4)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(5)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(6)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(7)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(8)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(9)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(10)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(11)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(12)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(13)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(14)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(15)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(16)%>" align="center"></td>
                                <td><img src="image/graph01.gif" width="15" height="<%=acpt_pro(17)%>" align="center"></td>
                                <td>&nbsp;</td>
							</tr>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

