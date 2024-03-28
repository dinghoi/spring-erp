<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Set dbconn = server.CreateObject("adodb.connection")
dbconn.open dbconnect

Sql="select * from memb where user_id='"&user_id&"'"
Set Rs=dbconn.execute(Sql)
if rs.eof or rs.bof then
	response.write"<script language=javascript>"
	response.write"alert('정보변경을 할 수 없습니다');"		
	response.write"parent.opener.location.reload();"
	response.write"self.close() ;"
	response.write"</script>"
end if
if rs("car_yn") = "" or isnull(rs("car_yn")) or rs("car_yn") = "N" then
	car_yn = "N" 
  else
	car_yn = "Y"
end if

sql = "select * from car_info where owner_emp_no = '"&user_id&"'"
Set rs_car=dbconn.execute(Sql)
if rs_car.eof or rs_car.bof then
	car_no = ""
	car_name = ""
	oil_kind = ""
  else
	car_no = rs_car("car_no")
	car_name = rs_car("car_name")
	oil_kind = rs_car("oil_kind")
end if

title_line = "사용자 정보 변경"
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			

			function chkfrm() {
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.car_yn[" + j + "].checked")) {
						k = j + 1
					}
				}
				if(document.frm.pass.value != document.frm.re_pass.value) {
					alert('비밀번호가 다릅니다.');
					frm.re_pass.focus();
					return false;}
//				if(document.frm.mod_pass.value =="") {
//					alert('변경비밀번호를 입력하세요');
//					frm.mod_pass.focus();
//					return false;}
				if(document.frm.mod_pass.value != document.frm.mod_re_pass.value) {
					alert('변경 확인 비밀번호가 다릅니다');
					frm.mod_pass.focus();
					return false;}
				if(document.frm.hp.value =="") {
					alert('핸드폰 번호를 입력하세요');
					frm.hp.focus();
					return false;}
				if(document.frm.old_car_yn.value =="Y") {
					if(k==1) {
						alert('정말 차량을 보유하지 않습니까??');
						}}
				if(k==2) {
					if(document.frm.car_no.value =="") {
						frm.car_no.focus();
						alert('차량번호를 입력하세요');
						return false;}}
				if(k==2) {
					if(document.frm.car_name.value =="") {
						frm.car_name.focus();
						alert('차종을 입력하세요');
						return false;}}
				if(k==2) {
					if(document.frm.oil_kind.value =="") {
						frm.oil_kind.focus();
						alert('유종을 입력하세요');
						return false;}}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function car_yn_view() 
			{
				k = 0;
				for (j=0;j<2;j++) {
					if (eval("document.frm.car_yn[" + j + "].checked")) {
						k = j + 1
					}
				}
				if (k==1) {
					document.getElementById('car_no_view').style.display = 'none'; 
					document.getElementById('car_name_view').style.display = 'none'; 
					document.getElementById('oil_kind_view').style.display = 'none'; }
				if (k==2) {
					document.getElementById('car_no_view').style.display = ''; 
					document.getElementById('car_name_view').style.display = ''; 
					document.getElementById('oil_kind_view').style.display = ''; }
			}
		</script>

	</head>
	<body onload="car_yn_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="user_mod_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">사용자</th>
								<td class="left"><%=rs("user_name")%>(<%=rs("user_id")%>)</td>
							</tr>
							<tr>
								<th class="first">기존비밀번호</th>
								<td class="left"><input name="re_pass" type="password" id="re_pass" style="width:150px"><input name="pass" type="hidden" id="pass" value="<%=rs("pass")%>"></td>
							</tr>
							<tr>
								<th class="first">변경비밀번호</th>
								<td class="left"><input name="mod_pass" type="password" id="mod_pass" onKeyUp="checklength(this,15);" style="width:150px"></td>
							</tr>
							<tr>
								<th class="first">변경확인비밀번호</th>
								<td class="left"><input name="mod_re_pass" type="password" id="mod_re_pass" style="width:150px"></td>
							</tr>
							<tr>
								<th class="first">핸드폰번호</th>
								<td class="left"><input name="hp" type="text" id="hp" value="<%=rs("hp")%>" style="width:150px"></td>
							</tr>
							<tr>
								<th class="first">차량유무</th>
								<td class="left">
                                <input type="radio" name="car_yn" value="N" <% if car_yn = "N" then %>checked<% end if %> style="width:25px"  onClick="car_yn_view()">미보유
								<input type="radio" name="car_yn" value="Y" <% if car_yn = "Y" then %>checked<% end if %> style="width:25px" onClick="car_yn_view()">보유
                                </td>
            				</tr>
							<tr id="car_no_view">
							  <th class="first">차량번호</th>
							  <td class="left"><input name="car_no" type="text" id="car_no" value="<%=car_no%>" style="width:150px"></td>
					        </tr>
							<tr id="car_name_view">
							  <th class="first">차종</th>
							  <td class="left"><input name="car_name" type="text" id="car_name" value="<%=car_name%>" style="width:150px"></td>
					        </tr>
							<tr id="oil_kind_view">
							  <th class="first">유종</th>
							  <td class="left">
                                <select name="oil_kind" id="oil_kind" style="width:150px">
								  <option value="">선택</option>
								  <option value="휘발유" <%If oil_kind = "휘발유" then %>selected<% end if %>>휘발유</option>
								  <option value="디젤" <%If oil_kind = "디젤" then %>selected<% end if %>>디젤</option>
								  <option value="가스" <%If oil_kind = "가스" then %>selected<% end if %>>가스</option>
							    </select>
                              </td>
					        </tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="변경" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
	                <input name="old_car_yn" type="hidden" id="old_car_yn" value="<%=car_yn%>">
	                <input name="old_car_no" type="hidden" id="old_car_no" value="<%=car_no%>">
				</form>
		</div>				
	</body>
</html>

