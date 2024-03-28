<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
user_id = request("user_id")

out_yn = "Y"
user_name = ""
user_grade = ""
org_name = ""
hp = ""
email = ""
mg_group = "1"
sms = "N"
help_yn = "N"
user_type = "Y"
reside_company = ""
reside_place = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Set Rs_trade = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "CE 등록"
if u_type = "U" then

	Sql="select * from memb where user_id = '" + user_id + "'"
	Set rs=DbConn.Execute(Sql)

	if rs("emp_no") = "999999" then
		out_yn = "Y"
	  else
	  	out_yn = "N"
	end if
	if rs("org_name") = "외주관리" then
		user_type = "N"
	  else
	  	user_type = "Y"
	end if
	user_id = rs("user_id")
	user_name = rs("user_name")
	user_grade = rs("user_grade")
	hp = rs("hp")
	org_name = rs("org_name")
	email = rs("email")
	grade = rs("grade")
	mg_group = rs("mg_group")
	reside = rs("reside")
	reside_place = rs("reside_place")
	reside_company = rs("reside_company")
	reside_company1 = rs("reside_company")
	sms = rs("sms")
	rs.close()

	title_line = "CE 변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function moveNext(varControl,varNext){
				var ctrType="";
			
				if(varControl.value.length == varControl.maxLength){
					varNext.focus();
					ctrType = varNext.type.toUpperCase();
					if(ctrType != "RADIO" && ctrType != "SELECT-ONE")
						varNext.select();
				}
			}

			function chkfrm() {
				if(document.frm.user_id.value =="") {
					alert('아이디를 입력하세요');
					frm.user_id.focus();
					return false;}
				if(document.frm.out_yn.value =="Y") {
					if(document.frm.user_name.value =="") {
						alert('사용자명을 입력하세요');
						frm.user_name.focus();
						return false;}}
				if(document.frm.out_yn.value =="Y") {
					if(document.frm.user_grade.value =="") {
						alert('사용자 직급을 입력하세요');
						frm.user_grade.focus();
						return false;}}
				if(document.frm.out_yn.value =="Y") {
					if(document.frm.org_name.value =="") {
						alert('부서명을 선택하세요');
						frm.org_name.focus();
						return false;}}
				if (eval("document.frm.user_type[0].checked")) {
					if(document.frm.org_name.value =="외주관리") {
						alert('계약직은 외주관리가 될수 없습니다. 부서명을 선택하세요');
						frm.org_name.focus();
						return false;}}
				if(document.frm.hp.value =="") {
					alert('핸드폰 번호를 입력하세요');
					frm.hp.focus();
					return false;}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function user_view() {
				
				if (eval("document.frm.user_type[0].checked")) {
					document.getElementById('org_view').style.display = '';
					document.getElementById('org_reg').style.display = 'none';
				}	
				if (eval("document.frm.user_type[1].checked")) {
					document.frm.org_name.value = "외주관리";
					document.frm.team.value = "외주관리";
					document.getElementById('org_view').style.display = 'none';
					document.getElementById('org_reg').style.display = '';
				}	
			}
		</script>
	</head>
	<body onload="user_view();">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="ce_reg_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">아이디</th>
								<td class="left">
                     			<%	if u_type = "U" then %>
                                <%=user_id%><input type="hidden" name="user_id" value="<%=user_id%>">
                                <%	  else	%>
                                <input name="user_id" type="text" id="user_id" style="width:120px" readonly="true"><a href="#" class="btnType03" onclick="javascript:pop_id_check()" >사용확인</a>
                                <% 	end if	%>
                                </td>
								<th>사용자명/직급</th>
								<td class="left">
								<% if out_yn = "N" then	%>
                                <%=user_name%>&nbsp;<%=user_grade%>
                                <%   else	%>
                                <input name="user_name" type="text" id="user_name" style="width:100px" onKeyUp="checklength(this,20)" value="<%=user_name%>">&nbsp;<input name="user_grade" type="text" id="user_grade" style="width:80px" onKeyUp="checklength(this,20)" value="<%=user_grade%>">
								<% end if	%>
                                </td>
							</tr>
							<tr>
								<th>사용자유형</th>
								<td class="left">
                                <input type="radio" name="user_type" value="Y" <% if user_type = "Y" then %>checked<% end if %> style="width:40px" id="Radio1" onClick="user_view()">계약직
                                <input type="radio" name="user_type" value="N" <% if user_type = "N" then %>checked<% end if %> style="width:40px" id="Radio1" onClick="user_view()">외주관리
                              </td>
								<th>부서명</th>
								<td class="left">
							<% if out_yn = "N" then	%>
                            	<%=org_name%>
                            <%  else	%>
								<div id = "org_view">
                                 <a href="#" onClick="pop_Window('org_search.asp','org_search_pop','scrollbars=yes,width=600,height=400')" class="btnType03">조직조회</a>
                                  <input name="org_name" type="text" value="<%=org_name%>" readonly="true" style="width:150px">
                                  <input name="emp_company" type="hidden" value="<%=emp_company%>">
                                  <input name="bonbu" type="hidden" value="<%=bonbu%>">
                                  <input name="saupbu" type="hidden" value="<%=saupbu%>">
                                  <input name="team" type="hidden" value="<%=team%>">
                                  <input name="reside_place" type="hidden" value="<%=reside_place%>">
                                  <input name="reside_company" type="hidden" value="<%=reside_company%>">
								</div>
								<div id = "org_reg">
                                  <input name="org_name1" type="hidden" value="<%=org_name%>">
                                  <input name="team1" type="hidden" value="<%=team%>">
                                  <strong>상 주 처 : </strong><input name="reside_place1" type="text" value="<%=reside_place%>" style="width:150px"><br>
                                  <strong>상주회사 : </strong>
								<%
								sql="select * from trade where (use_sw = 'Y') and (trade_id = '매출' or trade_id = '공통') order by trade_name asc"
                                rs_trade.Open Sql, Dbconn, 1
                                %>
                                <select name="reside_company1" id="reside_company1" style="width:150px">
          					<% 
								While not rs_trade.eof 
							%>
          							<option value='<%=rs_trade("trade_name")%>' <%If rs_trade("trade_name") = reside_company1  then %>selected<% end if %>><%=rs_trade("trade_name")%></option>
          					<%
									rs_trade.movenext()  
								Wend 
								rs_trade.Close()
							%>
                                </select>
								</div>
							<% end if	%>
								</td>
							</tr>
							<tr>
								<th class="first">핸드폰</th>
								<td class="left"><input name="hp" type="text" id="hp" style="width:120px" onKeyUp="checklength(this,13);" value="<%=hp%>"></td>
								<th>이메일</th>
								<td class="left"><input name="email" type="text" id="email" style="width:200px" onKeyUp="checklength(this,20)" value="<%=email%>"></td>
							</tr>
							<tr>
							  <th class="first">관리등급</th>
							  <td class="left"><select name="grade" id="grade" style="width:120px">
							    <option value="6" <% if grade = "6" then %>selected<% end if %>>권한대기</option>
							    <option value="5" <% if grade = "5" then %>selected<% end if %>>사용자</option>
							    <option value="4" <% if grade = "4" then %>selected<% end if %>>CE</option>
							    <option value="3" <% if grade = "3" then %>selected<% end if %>>상주CE</option>
							    <option value="2" <% if grade = "2" then %>selected<% end if %>>상주관리자</option>
							    <option value="1" <% if grade = "1" then %>selected<% end if %>>관리자</option>
							    <option value="0" <% if grade = "0" then %>selected<% end if %>>마스터</option>
						      </select></td>
							  <th>핼프</th>
							  <td class="left">
                              <input type="radio" name="help_yn" value="N" <% if help_yn = "N" then %>checked<% end if %> style="width:20px">NO
							  <input type="radio" name="help_yn" value="Y" <% if help_yn = "Y" then %>checked<% end if %> style="width:20px">YES 
							  </td>
				          </tr>
							<tr>
								<th class="first">관리그룹</th>
								<td class="left">
                                <input type="radio" name="mg_group" value="1" <% if mg_group = "1" then %>checked<% end if %> style="width:40px" id="Radio3">
일반그룹
  								<input type="radio" name="mg_group" value="2" <% if mg_group = "2" then %>checked<% end if %> style="width:40px" id="Radio4">
한진그룹 </td>
								<th>문자발송여부</th>
                                <td class="left">&nbsp;</td>
                            </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="out_yn" value="<%=out_yn%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

