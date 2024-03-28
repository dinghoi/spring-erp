<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
user_id = request("user_id")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql="select * from memb where user_id = '" + user_id + "'"
Set rs=DbConn.Execute(Sql)
cost_grade = rs("cost_grade")
title_line = "사용자별 비용 권한 변경"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
			function chkfrm() {

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
				<h3 class="tit"><%=title_line%></h3>
				<form action="cost_grade_mod_ok.asp" method="post" name="frm">
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
								<td class="left"><%=user_id%><input type="hidden" name="user_id" value="<%=rs("user_id")%>"></td>
								<th>사용자명</th>
								<td class="left"><%=rs("user_name")%>&nbsp;<%=rs("user_grade")%></td>
							</tr>
							<tr>
								<th class="first">부서명</th>
								<td class="left"><%=rs("org_name")%></td>
								<th>비용권한</th>
								<td class="left">
                                <select name="cost_grade" id="cost_grade" style="width:150px">
								  <option value="7" <% if cost_grade = "7" then %>selected<% end if %>>권한없음</option>
								  <option value="6" <% if cost_grade = "6" then %>selected<% end if %>>일반CE</option>
								  <option value="5" <% if cost_grade = "5" then %>selected<% end if %>>일반CE/관리</option>
								  <option value="4" <% if cost_grade = "4" then %>selected<% end if %>>영업및관리</option>
								  <option value="3" <% if cost_grade = "3" then %>selected<% end if %>>비용대행</option>
								  <option value="2" <% if cost_grade = "2" then %>selected<% end if %>>사업부장권한</option>
								  <option value="1" <% if cost_grade = "1" then %>selected<% end if %>>본부장권한</option>
							<% if emp_no = "100031" or emp_no = "900001" or emp_no = "100359" then	%>
								  <option value="0" <% if cost_grade = "0" then %>selected<% end if %>>마스터</option>
							<% end if	%>
							    </select>
								</td>
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
				</form>
		</div>				
	</body>
</html>

