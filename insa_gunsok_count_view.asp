<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

emp_company = request("emp_company")
condi = request("condi")
emp_bonbu = request("emp_bonbu")
gunsok_yy = request("gunsok_yy")
view_gunsok = gunsok_yy + "년"
sel_cnt = cint(gunsok_yy)

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)
target_date = curr_date

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if emp_company = "전체" then
      Sql = "select * from emp_master where (emp_company='"+emp_bonbu+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code ASC"
   elseif condi = "전체" then  
            if emp_company = emp_bonbu then
                   Sql = "select * from emp_master where (emp_company='"+emp_company+"') and (isNull(emp_bonbu) or emp_bonbu='') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code ASC"
			  else 
			       Sql = "select * from emp_master where (emp_company='"+emp_company+"') and (emp_bonbu='"+emp_bonbu+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code ASC"
		     end if
		  else
		     if condi = emp_bonbu then
		             Sql = "select * from emp_master where (emp_company='"+emp_company+"') and (emp_bonbu='"+condi+"') and (isNull(emp_saupbu) or emp_saupbu='') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code ASC"
			     else
				     Sql = "select * from emp_master where (emp_company='"+emp_company+"') and (emp_bonbu='"+condi+"') and (emp_saupbu='"+emp_bonbu+"') and (isNull(emp_end_date) or emp_end_date = '1900-01-01')  and (emp_no < '900000') ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code ASC"
			 end if
end if
Rs.Open Sql, Dbconn, 1

title_line = ""+ view_gunsok +" - 근속자 내역 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>근속자 내역</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}					
			function chkfrm() {
				if(document.frm.in_name.value =="") {
					alert('성명을 입력하세요');
					frm.in_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_gunsok_count_view.asp?emp_company=<%=emp_company%>" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="13%">
							<col width="12%">
                            <col width="10%">
                            <col width="9%">
                            <col width="9%">
                            <col width="9%">
                            <col width="*">
						</colgroup>
						<thead>
							<tr>
                                <th class="first" scope="col">소속</th>
                                <th scope="col">성&nbsp;&nbsp;명</th>
                                <th scope="col">최초입사일</th>
                                <th scope="col">직급</th>
                                <th scope="col">직책</th>
                                <th scope="col">생년월일</th>
                                <th scope="col">조직</th>
 							</tr>
						</thead>
						<tbody>
						<%
						    v_cnt = 0
							do until rs.eof or rs.bof
							    
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
						    if sel_cnt = 16 and target_cnt >= sel_cnt then
							   v_cnt = v_cnt + 1
						%>	
							<tr>
								<td class="left"><%=rs("emp_org_name")%>&nbsp;</td>
								<td>
                                <a href="#" onClick="pop_Window('insa_emp_master_view.asp?view_condi=<%=rs("emp_company")%>&emp_no=<%=rs("emp_no")%>&u_type=<%=""%>','insa_emp_modify_popup','scrollbars=yes,width=1250,height=480')"><%=rs("emp_name")%>(<%=rs("emp_no")%>)</a>
								</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%>&nbsp;</td>
							</tr>
						<%
						        elseif target_cnt = sel_cnt then 
								       v_cnt = v_cnt + 1
						%>
							<tr>
								<td class="left"><%=rs("emp_org_name")%>&nbsp;</td>
								<td>
                                <a href="#" onClick="pop_Window('insa_emp_master_view.asp?view_condi=<%=rs("emp_company")%>&emp_no=<%=rs("emp_no")%>&u_type=<%=""%>','insa_emp_modify_popup','scrollbars=yes,width=1250,height=480')"><%=rs("emp_name")%>(<%=rs("emp_no")%>)</a>
								</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%>&nbsp;</td>
							</tr>								
						<%		
						    end if
								rs.movenext()
							loop
							rs.close()
						%>
							<tr>
                        <% if v_cnt = 0 then %>
							     <td class="first" colspan="7" style=" border-top:1px solid #e3e3e3;">조회 내용이 없습니다.</td>
                        <%     else %>
								<td class="first" colspan="7" style=" border-top:1px solid #e3e3e3;">(<%=v_cnt%>)&nbsp;건이 조회되었습니다.</td>
                        <% end if %>
							</tr>
						</tbody>
					</table>
				</div>
			</div>				
	   </div>     
                   	<br>
               		<div align=right>
						<a href="#" class="btnType04" onclick="javascript:goAction()" >닫기</a>&nbsp;&nbsp;
					</div>
                    <br>       				
	</form>
	</body>
</html>

