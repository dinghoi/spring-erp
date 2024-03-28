<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
company = request("company")
org_gubun = request("org_gubun")
org_name = ""
used_sw = "Y"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

etc_code = "75" + company
Sql="select * from etc_code where etc_code = '" + etc_code + "'"
Set rs_etc=DbConn.Execute(SQL)
if rs_etc.eof or rs_etc.bof then 
	company_name = "없음"
  else 
	company_name = rs_etc("etc_name")
end if
rs_etc.close()						

if org_gubun = "1" then
	org_gubun_name = "관리조직"
end if
if org_gubun = "2" then
	if company = "01" then	
		org_gubun_name = "법인명"
	  else
		org_gubun_name = "상위조직"
	end if
end if
if u_type = "U" then
	org_code = request("org_code")
	sql="select * from org_code where org_company='" + company + "' and org_gubun = '" + org_gubun + "' and org_code = '" + org_code + "'"
	set rs=dbconn.execute(sql)
	org_name = rs("org_name")
	used_sw = rs("used_sw")
	rs.close()	
	title_line = "조직관리 코드 변경"
end if

title_line = "조직관리 코드 등록"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
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
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {

				if(document.frm.org_name.value == "") {
					alert('코드명을 입력하세요!!');
					frm.org_name.focus();
					return false;}
				{
				a=confirm('등록을 하시겠습니까?')
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
				<form action="org_code_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">회사</th>
								<td class="left"><%=company_name%><input name="company" type="hidden" id="company" value="<%=company%>"></td>
							</tr>
							<tr>
								<th class="first">구분코드</th>
								<td class="left"><%=org_gubun_name%><input name="org_gubun" type="hidden" id="org_gubun" value="<%=org_gubun%>"><input name="org_code" type="hidden" id="org_code" value="<%=org_code%>"></td>
							</tr>
							<tr>
								<th class="first">코드명</th>
								<td class="left"><input name="org_name" type="text" id="org_name" value="<%=org_name%>" style="width:200px" onKeyUp="checklength(this,30)"></td>
							</tr>
							<tr>
								<th class="first">사용유무</th>
								<td class="left">
								<input name="used_sw" type="radio" value="Y" <% if used_sw = "Y" then %>checked<% end if %>>사용
              					<input type="radio" name="used_sw" value="N" <% if used_sw = "N" then %>checked<% end if %>>미사용
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

