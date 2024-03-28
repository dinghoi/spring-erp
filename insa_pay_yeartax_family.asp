<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim year_tab(3,2)

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_pay_yeartax_family.asp"

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
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_year = Server.CreateObject("ADODB.Recordset")
Set rs_bef = Server.CreateObject("ADODB.Recordset")
Set rs_ins = Server.CreateObject("ADODB.Recordset")
Set rs_ann = Server.CreateObject("ADODB.Recordset")
Set rs_fami = Server.CreateObject("ADODB.Recordset")
Set rs_medi = Server.CreateObject("ADODB.Recordset")
Set rs_edu = Server.CreateObject("ADODB.Recordset")
Set rs_dona = Server.CreateObject("ADODB.Recordset")
Set rs_duct = Server.CreateObject("ADODB.Recordset")
Set rs_cred = Server.CreateObject("ADODB.Recordset")
Set rs_hous = Server.CreateObject("ADODB.Recordset")
Set rs_houm = Server.CreateObject("ADODB.Recordset")
Set rs_savi = Server.CreateObject("ADODB.Recordset")
Set rs_other = Server.CreateObject("ADODB.Recordset")
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

sql = "select * from pay_yeartax where y_year = '"&inc_yyyy&"' and y_emp_no = '"&emp_no&"'"
rs_year.Open Sql, Dbconn, 1
if not rs_year.eof then
       y_final =  rs_year("y_final") 
   else	   
	   y_final =  ""
end if
rs_year.close()	

sql = "select * from emp_family where family_empno = '"&emp_no&"' ORDER BY family_empno,family_seq ASC"
Rs.Open Sql, Dbconn, 1

title_line = "연말정산 - 전산등록 안내 및 주의사항 "
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
				<form action="insa_pay_yeartax_family.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
							<col width="6%" >
							<col width="8%" >
							<col width="10%" >
							<col width="6%" >
							<col width="6%" >
                            
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="8%" >
                            
                            <col width="*" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" scope="col" class="first">관계</th>
                                <th rowspan="2" scope="col">성명</th>
								<th rowspan="2" scope="col">주민등록번호</th>
								<th rowspan="2" scope="col">내외국인<br>구분</th>
                                <th rowspan="2" scope="col">부양여부</th>
								<th colspan="7" scope="col" style=" border-bottom:1px solid #e3e3e3;">구분</th>
                                <th rowspan="2" scope="col">기타</th>
                                <th rowspan="2" scope="col">수정</th>
							</tr>
                            <tr>
				                <th class="first"scope="col" style=" border-left:1px solid #e3e3e3;">장애인</th>
				                <th scope="col">국가유공자</th>
                                <th scope="col">중증환자</th>
                                <th scope="col">수급자</th>
                                <th scope="col">위탁아동</th>
                                <th scope="col">입양여부</th>
                                <th scope="col">입양일자</th>
                            </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
                           family_support_yn = rs("family_support_yn")
						   family_disab = rs("family_disab")
						   family_merit = rs("family_merit")
						   family_serius = rs("family_serius")
						   family_pensioner = rs("family_pensioner")
						   family_witak = rs("family_witak")
						   family_holt = rs("family_holt")
						   if rs("family_holt_date") = "1900-01-01" then
						            family_holt_date = ""
							  else 
							        family_holt_date = rs("family_holt_date")
						   end if
	           			%>
							<tr>
                                <td><%=rs("family_rel")%>&nbsp;</td>
                                <td><%=rs("family_name")%>&nbsp;</td>
                                <td><%=rs("family_person1")%>-<%=rs("family_person2")%>&nbsp;</td>
                                <td><%=rs("family_national")%>&nbsp;</td>
                                <td><input type="checkbox" name="support_check" value="Y" <% if family_support_yn = "Y" then %>checked<% end if %> id="support_check"></td>
                                <td>
								<input type="checkbox" name="disab_check" value="Y" <% if family_disab = "Y" then %>checked<% end if %> id="disab_check"></td>
                                <td><input type="checkbox" name="merit_check" value="Y" <% if family_merit = "Y" then %>checked<% end if %> id="merit_check"></td>
                                <td><input type="checkbox" name="serius_check" value="Y" <% if family_serius = "Y" then %>checked<% end if %> id="serius_check"></td>
                                <td><input type="checkbox" name="pensioner_check" value="Y" <% if family_pensioner = "Y" then %>checked<% end if %> id="pensioner_check"></td>
                                <td><input type="checkbox" name="witak_check" value="Y" <% if family_witak = "Y" then %>checked<% end if %> id="witak_check"></td>
                                <td><input type="checkbox" name="holt_check" value="Y" <% if family_holt = "Y" then %>checked<% end if %> id="holt_check"></td>
                                <td><%=family_holt_date%>&nbsp;</td>
                                <td>&nbsp;</td>
                        <% if y_final <> "Y" then  %>
                                <td>
                                <a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=rs("family_empno")%>&family_seq=<%=rs("family_seq")%>&emp_name=<%=emp_name%>&u_type=<%="U"%>','insa_family_add_pop','scrollbars=yes,width=750,height=370')">수정</a></td>
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
					<a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=emp_no%>&emp_name=<%=emp_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=370')" class="btnType04">부양가족추가</a>
					</div>    
              <%   else  %>
                       <br><br>
			  <%   end if  %>                    
                    </td>
			      </tr>
				  </table>
           <h3 class="stit">※ 연말정산 전산 등록방법 및 주의 사항 ※<br>&nbsp;<br>
                1. 가족사항 필수 등록<br>
                &nbsp;&nbsp;&nbsp;&nbsp;■ 소득공제신고서 등록전에 인사관리>가족사항에서 가족의 소득공제 정보를 확인할 것<br>
                &nbsp;&nbsp;&nbsp;&nbsp;■ 기본공제 또는 특별공제를 받을 가족 둥 미등록자는 반드시 등록을 해야 함(등록시 주민등록번호는 필수 사항임)<br>&nbsp;<br>
                2. 의료비/기부금명세서 작성방법<br>
                &nbsp;&nbsp;&nbsp;&nbsp;■ 소득공제신고서 상에 의료비. 기부금 내역을 등록하면 해당 명세서를 출력할 수 있으므로 별도의 파일에 저장할 필요는 없음.<br>
                &nbsp;&nbsp;&nbsp;&nbsp;■ 의료비 작성시 국세청 자료와 의료기관 자료를 중복하여 입력하면 추후 추징대상이 될 수 있으니 반드시 확인하여 기재<br>
                &nbsp;&nbsp;&nbsp;&nbsp;■ 기부금은 본인과 기본공제대상자인 배우자 및 부양가족이 지급한 기부금이 공제대상임.<br>&nbsp;<br>
                3. 국세청금액/그밖의금액 구분등록<br>
                &nbsp;&nbsp;&nbsp;&nbsp;■ 공제자료 등록시 국세청에서 발급받은 자료는 국세청금액에 입력하고, 그 외 자료는 기타금액(그 밖의금액)에 구분하여 등록해야 함.<br>&nbsp;<br>
                4. 신용카드,현금영수증.직불카드 등록<br>
                &nbsp;&nbsp;&nbsp;&nbsp;■ 신용카드/련금영수증/직불카드 금액 입력시 일반합계와 전통시장 및 대붕교통 사용 합계액을 구분하여 입력.</h3>
                <input type="hidden" name="family_empno" value="<%=in_empno%>" ID="Hidden1">  
                <input type="hidden" name="y_final" value="<%=y_final%>" ID="Hidden1">               
			</form>
		</div>				
	</div>        				
	</body>
</html>

