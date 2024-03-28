<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
family_empno = request("family_empno")
family_seq = request("family_seq")
emp_name = request("emp_name")

family_rel = ""
family_name = ""
family_birthday = ""
family_birthday_id = ""
family_job = ""
family_live = "안함"
family_person1 = ""
family_person2 = ""
family_tel_ddd = ""
family_tel_no1 = ""
family_tel_no2 = ""
family_support_yn = "N"
family_national = "내국인"
family_disab = ""
family_merit = ""
family_serius = ""
family_pensioner = ""
family_witak = ""
family_holt = ""
family_holt_date = ""
family_children = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 가족사항 등록 "
if u_type = "U" then

	Sql="select * from emp_family where family_empno = '"&family_empno&"' and family_seq = '"&family_seq&"'"
	Set rs=DbConn.Execute(Sql)

	family_rel = rs("family_rel")
    family_name = rs("family_name")
    family_birthday = rs("family_birthday")
    family_birthday_id = rs("family_birthday_id")
    family_job = rs("family_job")
    family_live = rs("family_live")
    family_person1 = rs("family_person1")
    family_person2 = rs("family_person2")
	family_tel_ddd = rs("family_tel_ddd")
    family_tel_no1 = rs("family_tel_no1")
    family_tel_no2 = rs("family_tel_no2")
	family_support_yn = rs("family_support_yn")
	if family_birthday = "1900-01-01"  then
	   family_birthday = ""
	end if
	family_national = rs("family_national")
    family_disab = rs("family_disab")
	family_merit = rs("family_merit")
    family_serius = rs("family_serius")
    family_pensioner = rs("family_pensioner")
    family_witak = rs("family_witak")
    family_holt = rs("family_holt")
    family_holt_date = rs("family_holt_date")
	if family_holt_date = "1900-01-01"  then
	   family_holt_date = ""
	end if
	family_children = rs("family_children")
	
	rs.close()

	title_line = " 가족사항 변경 "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=family_birthday%>" );
			});	
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=family_holt_date%>" );
			});	
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.family_birthday.value =="") {
					alert('생년월일을 입력하세요');
					frm.family_birthday.focus();
					return false;}
//				if(document.frm.family_person1.value =="") {
//					alert('주민등록번호를 입력하세요');
//					frm.family_person1.focus();
//					return false;}
//				if(document.frm.family_person2.value =="") {
//					alert('주민등록번호를 입력하세요');
//					frm.family_person2.focus();
//					return false;}
				if(document.frm.family_rel =="") {
					alert('관계항목을 선택하세요');
					frm.family_rel.focus();
					return false;}
				if(document.frm.family_name.value =="") {
					alert('가족성명을 입력하세요');
					frm.family_name.focus();
					return false;}
				if(document.frm.family_tel_no1.value =="") {
					alert('전화번호를 입력하세요');
					frm.family_tel_no1.focus();
					return false;}
				if(document.frm.family_tel_no2.value =="") {
					alert('전화번호를 입력하세요');
					frm.family_tel_no2.focus();
					return false;}
				if(document.frm.family_support_yn.value =="") {
					alert('부양가족여부를 입력하세요');
					frm.family_support_yn.focus();
					return false;}
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_family_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="15%" >
						<col width="18%" >
						<col width="15%" >
						<col width="18%" >
						<col width="15%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="family_empno" type="text" id="family_empno" size="14" value="<%=family_empno%>" readonly="true">
                      <input type="hidden" name="family_seq" value="<%=family_seq%>" ID="Hidden1"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                 	<tr>
                      <th>관계(필수)</th>
                      <td colspan="5" class="left">
					  <select name="family_rel" id="family_rel" value="<%=family_rel%>" style="width:100px">
				          <option value="" <% if family_rel = "" then %>selected<% end if %>>선택</option>
				          <option value='부' <%If family_rel = "부" then %>selected<% end if %>>부</option>
				          <option value='모' <%If family_rel = "모" then %>selected<% end if %>>모</option>
				          <option value='남편' <%If family_rel = "남편" then %>selected<% end if %>>남편</option>
                          <option value='아내' <%If family_rel = "아내" then %>selected<% end if %>>아내</option>
                          <option value='아들' <%If family_rel = "아들" then %>selected<% end if %>>아들</option>
                          <option value='딸' <%If family_rel = "딸" then %>selected<% end if %>>딸</option>
                          <option value='조부' <%If family_rel = "조부" then %>selected<% end if %>>조부</option>
                          <option value='조모' <%If family_rel = "조모" then %>selected<% end if %>>조모</option>
                          <option value='외조부' <%If family_rel = "외조부" then %>selected<% end if %>>외조부</option>
                          <option value='외조모' <%If family_rel = "외조모" then %>selected<% end if %>>외조모</option>
                          <option value='시부' <%If family_rel = "시부" then %>selected<% end if %>>시부</option>
                          <option value='시모' <%If family_rel = "시모" then %>selected<% end if %>>시모</option>
                          <option value='장인' <%If family_rel = "장인" then %>selected<% end if %>>장인</option>
                          <option value='장모' <%If family_rel = "장모" then %>selected<% end if %>>장모</option>
                          <option value='형(형제자매)' <%If family_rel = "형(형제자매)" then %>selected<% end if %>>형(형제자매)</option>
                          <option value='매(형제자매)' <%If family_rel = "매(형제자매)" then %>selected<% end if %>>매(형제자매)</option>
                          <option value='제(형제자매)' <%If family_rel = "제(형제자매)" then %>selected<% end if %>>제(형제자매)</option>
                          <option value='올케' <%If family_rel = "올케" then %>selected<% end if %>>올케</option>
                          <option value='자(형제자매)' <%If family_rel = "자(형제자매)" then %>selected<% end if %>>자(형제자매)</option>
                          <option value='손자' <%If family_rel = "손자" then %>selected<% end if %>>손자</option>
                          <option value='손녀' <%If family_rel = "손녀" then %>selected<% end if %>>손녀</option>
                          <option value='자부' <%If family_rel = "자부" then %>selected<% end if %>>자부</option>
                          <option value='손부' <%If family_rel = "손부" then %>selected<% end if %>>손부</option>
                          <option value='기타관계' <%If family_rel = "기타관계" then %>selected<% end if %>>기타관계</option>
                      </select>
                      &nbsp;위탁아동인경우는 기타관계를 선택하십시요!
                      </td>
                    </tr>
                    <tr>
                      <th>성명(필수)</th>
                      <td colspan="2" class="left">
					  <input name="family_name" type="text" id="family_name" size="14" value="<%=family_name%>"></td>
                      <th>출생지</th>
                      <td colspan="2" class="left">
					  <select name="family_national" id="family_national" value="<%=family_national%>" style="width:90px">
				          <option value="" <% if family_rel = "" then %>selected<% end if %>>선택</option>
				          <option value='내국인' <%If family_national = "내국인" then %>selected<% end if %>>내국인</option>
				          <option value='외국인' <%If family_national = "외국인" then %>selected<% end if %>>외국인</option>
                      </select>
                      </td>
                    </tr>
                    <tr>
                      <th>생년월일(필수)</th>
                      <td colspan="2" class="left">
					  <input name="family_birthday" type="text" value="<%=family_birthday%>" style="width:70px;text-align:center" id="datepicker" readonly="true">
					  &nbsp;&nbsp;
					  <input type="radio" name="family_birthday_id" value="양" <% if family_birthday_id = "양" then %>checked<% end if %>>양
              		  <input name="family_birthday_id" type="radio" value="음" <% if family_birthday_id = "음" then %>checked<% end if %>>음
					  </td>
                      <th>주민등록번호</th>
                      <td colspan="2" class="left">
                      <input name="family_person1" type="text" id="family_person1" size="6" maxlength="6" value="<%=family_person1%>" >
					  -
                      <input name="family_person2" type="text" id="family_person2" size="7" maxlength="7" value="<%=family_person2%>" >
                      &nbsp;(연말정산필수)
				      </td>
                   </tr>
                   <tr>
                      <th>직업</th>
                      <td colspan="2" class="left">
					  <input name="family_job" type="text" id="family_job" style="width:160px; ime-mode:active" onKeyUp="checklength(this,20);" value="<%=family_job%>"></td>
                      <th>전화번호</th>
                      <td colspan="2" class="left">
                      <input name="family_tel_ddd" type="text" id="family_tel_ddd" size="3" maxlength="3" value="<%=family_tel_ddd%>" >
								  -
                      <input name="family_tel_no1" type="text" id="family_tel_no1" size="4" maxlength="4" value="<%=family_tel_no1%>" >
                                  -
                      <input name="family_tel_no2" type="text" id="family_tel_no2" size="4" maxlength="4" value="<%=family_tel_no2%>" >
					  </td>
                    </tr>
                    <tr>
                      <th>동거여부</th>
                      <td colspan="2" class="left">
					  <input type="radio" name="family_live" value="동거" <% if family_live = "동거" then %>checked<% end if %>>동거 
              		  <input name="family_live" type="radio" value="안함" <% if family_live = "안함" then %>checked<% end if %>>안함
                      <th>부양가족</th>
                      <td colspan="2" class="left">
					  <input type="radio" name="family_support_yn" value="Y" <% if family_support_yn = "Y" then %>checked<% end if %>>부양 
              		  <input name="family_support_yn" type="radio" value="N" <% if family_support_yn = "N" then %>checked<% end if %>>안함
					  </td>
                    </tr>
                    <tr>
                      <th>장애인</th>
                      <td colspan="2" class="left">
					  <input type="checkbox" name="disab_check" value="Y" <% if family_disab = "Y" then %>checked<% end if %> id="disab_check">장애인
              		  <input type="checkbox" name="merit_check" value="Y" <% if family_merit = "Y" then %>checked<% end if %> id="merit_check">국가유공자
                      <input type="checkbox" name="serius_check" value="Y" <% if family_serius = "Y" then %>checked<% end if %> id="serius_check">중증환자
					  </td>
                      <th>국민기초생활수급</th>
                      <td colspan="2" class="left">
					  <input type="checkbox" name="pensioner_check" value="Y" <% if family_pensioner = "Y" then %>checked<% end if %> id="pensioner_check">예
                    </tr>
                    <tr>
                      <th>입양여부</th>
                      <td class="left">
					  <input type="checkbox" name="holt_check" value="Y" <% if family_holt = "Y" then %>checked<% end if %> id="holt_check">예
                      </td>
                      <th>입양일자</th>
                      <td class="left">
              		  <input name="family_holt_date" type="text" value="<%=family_holt_date%>" style="width:70px;text-align:center" id="datepicker1" readonly="true">
					  </td>
                      <th>위탁아동</th>
                      <td class="left">
					  <input type="checkbox" name="witak_check" value="Y" <% if family_witak = "Y" then %>checked<% end if %> id="witak_check">예
                    </tr>
                    <tr>
                      <th>자녀양육</th>
                      <td colspan="5" class="left">
					  <input type="checkbox" name="children_check" value="Y" <% if family_children = "Y" then %>checked<% end if %> id="children_check">예&nbsp;&nbsp;(6세미만 자녀의경우 연말정산 추가공제 체크)
                      </td>
                    </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

