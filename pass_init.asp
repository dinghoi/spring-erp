<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
user_id = Request("user_id")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

SQL = "select * from memb where user_id = '" + user_id + "'"
set rs=dbconn.execute(sql)

emp_view = "정직"
emp_yn = "Y"
sql = "select * from emp_master where emp_no = '"+rs("emp_no")+"'"
set rs_emp=dbconn.execute(sql)
if rs_emp.eof or rs_emp.bof then
	emp_view = "외주직원"
	emp_yn = "N"
end if


title_line = "비밀번호 초기화"
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
			function chkfrm() {

				{
				a=confirm('초기화 하시겠습니까?')
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
				<form action="pass_init_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="40%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">이름 / 아이디</th>
								<td class="left"><%=rs("user_name")%>&nbsp;/&nbsp;<%=rs("user_id")%></td>
							</tr>
							<tr>
								<th class="first">직원여부</th>
								<td class="left"><%=emp_view%></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <strong>정직은 주민번호 뒤 7자리, 외주직원 '1111' 초기화</strong>
                </div>
				<br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="비밀번호 초기화" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="user_id" value="<%=user_id%>" ID="Hidden1">
				<input type="hidden" name="emp_yn" value="<%=emp_yn%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

