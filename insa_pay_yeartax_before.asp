<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim year_tab(3,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_before.asp"

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

y_final=Request("y_final")
ck_sw=Request("ck_sw")

if ck_sw = "n" then
	inc_yyyy = request.form("inc_yyyy")
  else
	inc_yyyy = request("inc_yyyy")
end if

if view_condi = "" then
	'inc_yyyy = mid(cstr(now()),1,4)
	inc_yyyy = cint(mid(now(),1,4)) - 1
	ck_sw = "n"
end if

' 최근3개년도 테이블로 생성
'year_tab(3,1) = mid(now(),1,4)
'year_tab(3,2) = cstr(year_tab(3,1)) + "년"
'year_tab(2,1) = cint(mid(now(),1,4)) - 1
'year_tab(2,2) = cstr(year_tab(2,1)) + "년"
'year_tab(1,1) = cint(mid(now(),1,4)) - 2
'year_tab(1,2) = cstr(year_tab(1,1)) + "년"

' 최근3개년도 테이블로 생성
year_tab(3,1) = cint(mid(now(),1,4)) - 1
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 2
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 3
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect


Sql = "select * from emp_master where emp_no = '"&emp_no&"'"
rs_emp.Open Sql, Dbconn, 1
emp_in_date = rs_emp("emp_in_date")
emp_name = rs_emp("emp_name")
emp_grade = rs_emp("emp_grade")
emp_position = rs_emp("emp_position")
emp_company = rs_emp("emp_company")
emp_org_name = rs_emp("emp_org_name")

sql = "select * from pay_yeartax_before where b_year = '"&inc_yyyy&"' and b_emp_no = '"&emp_no&"' ORDER BY b_emp_no,b_seq ASC"
'sql = "select * from emp_family where family_empno = '"&emp_no&"' ORDER BY family_empno,family_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 이전근무지 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
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
			function frmcheck () {
				if (formcheck(document.frm)) {
					document.frm.submit ();
				}
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
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_person_yeartax_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_yeartax_before.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
                                <label>
                             <strong>사번 : </strong>
                                <input name="emp_no" type="text" value="<%=emp_no%>" style="width:50px" readonly="true">
                                -
                                <input name="emp_name" type="text" value="<%=emp_name%>" style="width:60px" readonly="true">
                                </label>
                                <label>
                             <strong>직급 : </strong>
                                <input name="emp_grade" type="text" value="<%=emp_grade%>" style="width:60px" readonly="true">
                                -
                                <input name="emp_position" type="text" value="<%=emp_position%>" style="width:70px" readonly="true">
                                </label>
                                <label>
                             <strong>입사일 : </strong>
                                <input name="emp_in_date" type="text" value="<%=emp_in_date%>" style="width:70px" readonly="true">
                                </label>
                                <label>
                             <strong>소속 : </strong>
                                <input name="emp_company" type="text" value="<%=emp_company%>" style="width:90px" readonly="true">
                                -
                                <input name="emp_org_name" type="text" value="<%=emp_org_name%>" style="width:90px" readonly="true">
                                </label>
                             <strong>귀속년도 : </strong>
                                <select name="inc_yyyy" id="inc_yyyy" type="text" value="<%=inc_yyyy%>" style="width:70px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If inc_yyyy = cstr(year_tab(i,1)) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
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
							<col width="8%" >
							<col width="*" >
							<col width="7%" >
							<col width="7%" >
                            
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
                            
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
				                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">사업자번호</th>
				                <th scope="col">근무처명</th>
                                <th scope="col">근무<br>시작일</th>
                                <th scope="col">근무<br>종료일</th>
                                <th scope="col">급여</th>
                                <th scope="col">상여</th>
                                <th scope="col">인정상여등</th>
                                <th scope="col">비과세</th>
                                <th scope="col">국민연금</th>
                                <th scope="col">건강보험</th>
                                <th scope="col">고용보험</th>
                                <th scope="col">장기<br>요양보험</th>
                                <th scope="col">(결정세액)<br>소득세</th>
                                <th scope="col">(결정세액)<br>주민세</th>
                                <th scope="col">비고</th>
                            </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

	           			%>
							<tr>
                                <td><%=rs("b_company_no")%>&nbsp;</td>
                                <td><%=rs("b_company")%>&nbsp;</td>
                                <td><%=rs("b_from_date")%>&nbsp;</td>
                                <td><%=rs("b_to_date")%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_pay"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_bonus"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_deem_bonus"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_overtime_taxno"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_nps"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_nhis"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_epi"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_longcare"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_income_tax"),0)%>&nbsp;</td>
                                <td class="right"><%=formatnumber(rs("b_wetax"),0)%>&nbsp;</td>
                        <% if y_final <> "Y" then  %>                                
                                <td>
                                <a href="#" onClick="pop_Window('insa_pay_yeartax_before_add.asp?b_year=<%=rs("b_year")%>&b_emp_no=<%=rs("b_emp_no")%>&b_seq=<%=rs("b_seq")%>&b_emp_name=<%=rs("b_emp_name")%>&u_type=<%="U"%>','insa_pay_yeartax_before_add_pop','scrollbars=yes,width=800,height=370')">수정</a></td>
                        <%    else  %>
                                <td>&nbsp;</td>
                        <% end if  %>                                
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
              <% if y_final <> "Y" then  %>
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_pay_yeartax_before_add.asp?b_year=<%=inc_yyyy%>&b_emp_no=<%=emp_no%>&b_emp_name=<%=emp_name%>&u_type=<%=""%>','insa_pay_yeartax_before_add_pop','scrollbars=yes,width=800,height=370')" class="btnType04">이전근무지 입력</a>
					</div>  
              <%   else  %>
                       <br><br>
			  <%   end if  %>                      
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="in_emp_no" value="<%=emp_no%>" ID="Hidden1">      
                <input type="hidden" name="y_final" value="<%=y_final%>" ID="Hidden1">             
			</form>
		</div>				
	</div>        				
	</body>
</html>

