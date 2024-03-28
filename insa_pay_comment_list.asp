<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw
dim month_tab(24,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

view_condi = request("view_condi")
owner_view=request("owner_view")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_company = request.form("view_company")
	view_condi = request.form("view_condi")
	owner_view=Request.form("owner_view")
	pmg_yymm=Request.form("pmg_yymm")
  else
	view_company = request("view_company")
	view_condi = request("view_condi")
	owner_view=request("owner_view")
	pmg_yymm=request("pmg_yymm")
end if

if view_condi = "" then
	view_condi = ""
	owner_view = "C"
	ck_sw = "n"
	view_company = "케이원정보통신"
	from_date = mid(cstr(now()-curr_dd+1),1,10)
	'pmg_yymm = mid(cstr(from_date),1,4) + mid(cstr(from_date),6,2)
	pmg_yymm = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
end if

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
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_give = Server.CreateObject("ADODB.Recordset")
Set Rs_dct = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi <> "" then
     if owner_view = "C" then  
	     Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_company+"') and (pmg_emp_name like '%"+view_condi+"%') ORDER BY pmg_emp_no ASC"
       else
		 Sql = "select * from pay_month_give where (pmg_yymm = '"+pmg_yymm+"') and (pmg_id = '1') and (pmg_company = '"+view_company+"') and (pmg_emp_no = '"+view_condi+"') ORDER BY pmg_emp_no ASC"
     end if
	 Rs.Open Sql, Dbconn, 1
end if
'Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " 급여 특이사항 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>급여관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
			}
			function goAction () {
			   window.close () ;
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("조건을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}

		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pay_header.asp" -->
			<!--#include virtual = "/include/insa_pay_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_comment_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
                                <strong>회사 : </strong>
                              <%
								Sql="select * from emp_org_mst where  org_level = '회사' ORDER BY org_code ASC"
	                            rs_org.Open Sql, Dbconn, 1	
							  %>
                                <label>
								<select name="view_company" id="view_company" type="text" style="width:130px">
                			  <% 
								do until rs_org.eof 
			  				  %>
                					<option value='<%=rs_org("org_name")%>' <%If view_company = rs_org("org_name") then %>selected<% end if %>><%=rs_org("org_name")%></option>
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
								</label>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">사번
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">성명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="9%" >
							<col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
                            <col width="7%" >
							<col width="*" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
                                <th scope="col">입사일</th>
                                <th scope="col">소속</th>
								<th scope="col">기본급</th>
                                <th scope="col">지급액계</th>
                                <th scope="col">공제액계</th>
                                <th scope="col">차인지급액</th>
								<th scope="col">특이사항</th>
                                <th scope="col">등록</th>
                                <th scope="col">수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						if  view_condi <> "" then 
						do until rs.eof
						      emp_no = rs("pmg_emp_no")
							  pmg_give_tot = rs("pmg_give_total")
							   
							  Sql = "SELECT * FROM emp_master where emp_no = '"+emp_no+"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not rs_emp.eof then
									emp_first_date = rs_emp("emp_first_date")
									emp_in_date = rs_emp("emp_in_date")
	                             else
									emp_first_date = ""
									emp_in_date = ""
                              end if
                              rs_emp.close() 
						%>
							<tr>
                              <td class="first"><%=rs("pmg_emp_no")%>&nbsp;</td>
                              <td><%=rs("pmg_emp_name")%></td>
                              <td><%=rs("pmg_grade")%>&nbsp;</td>
                              <td><%=rs("pmg_position")%>&nbsp;</td>
                              <td><%=emp_in_date%>&nbsp;</td>
                              <td><%=rs("pmg_org_name")%>&nbsp;</td>
                              <td class="right"><%=formatnumber(rs("pmg_base_pay"),0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(rs("pmg_give_total"),0)%>&nbsp;</td>
                         <%
						      Sql = "select * from pay_month_deduct where (de_yymm = '"+pmg_yymm+"' ) and (de_id = '1') and (de_emp_no = '"+emp_no+"')"
                              Set Rs_dct = DbConn.Execute(SQL)
							  if not Rs_dct.eof then
									de_deduct_tot = Rs_dct("de_deduct_total")
	                             else
									de_deduct_tot = 0
                              end if
                              Rs_dct.close()
							  
							  pmg_curr_pay = pmg_give_tot - de_deduct_tot
							  
							  if rs("pmg_comment") = "" or isnull(rs("pmg_comment")) then
							         task_memo = ""
								     view_memo = ""
								 else
							         task_memo = replace(rs("pmg_comment"),chr(34),chr(39))
							         view_memo = task_memo
							         if len(task_memo) > 16 then
							   	        view_memo = mid(task_memo,1,16) + "..."
							         end if	
							  end if
                          %>     
                              <td class="right"><%=formatnumber(de_deduct_tot,0)%>&nbsp;</td>
                              <td class="right"><%=formatnumber(pmg_curr_pay,0)%>&nbsp;</td>
                              <td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
                              <td ><a href="#" onClick="pop_Window('insa_pay_comment_add.asp?pmg_emp_no=<%=emp_no%>&pmg_emp_name=<%=rs("pmg_emp_name")%>&owner_view=<%=owner_view%>&view_company=<%=view_company%>&pmg_yymm=<%=pmg_yymm%>&u_type=<%=""%>','insa_comment_add_pop','scrollbars=yes,width=900,height=500')">등록</a>
                              </td>
                         <% if insa_grade = "0" then %>     
                              <td ><a href="#" onClick="pop_Window('insa_pay_comment_add.asp?pmg_emp_no=<%=emp_no%>&pmg_emp_name=<%=rs("pmg_emp_name")%>&owner_view=<%=owner_view%>&view_company=<%=view_company%>&pmg_yymm=<%=pmg_yymm%>&u_type=<%="U"%>','insa_comment_add_pop','scrollbars=yes,width=900,height=500')">수정</a>
                              </td>
                         <%     else %>
                              <td>&nbsp;</td>
                         <% end if %>          
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						
						end if
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <% if owner_view = "T" then 
                              emp_no = view_condi
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
							  end if
							  rs_emp.close()
				    %>
					<a href="#" onClick="pop_Window('insa_pay_comment_add.asp?pmg_emp_no=<%=emp_no%>&pmg_emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&view_company=<%=view_company%>&pmg_yymm=<%=pmg_yymm%>&u_type=<%=""%>','insa_comment_add_pop','scrollbars=yes,width=900,height=500')" class="btnType04">급여특이사항등록</a>
                    <% end if %>
					</div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="cmt_empno" value="<%=cmt_empno%>" ID="Hidden1">
                  <input type="hidden" name="cmt_date" value="<%=cmt_date%>" ID="Hidden1">
                  <input type="hidden" name="cmt_empname" value="<%=cmt_empname%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

