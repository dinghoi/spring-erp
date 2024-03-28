<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
trade_code = request("trade_code")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_trade = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "서비스 회사 항목 변경"

Sql="select * from trade where trade_code = '"&trade_code&"'"
Set rs=DbConn.Execute(Sql)

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//ENrs("customer_no")http://www.w3.org/TR/html4/loose.dtd">
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
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.support_company.value =="") {
					alert('지원회사를 입력하세요');
					frm.support_company.focus();
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
				<h3 class="tit"><%=title_line%></h3>
				<form action="company_mod_ok.asp" method="post" name="frm">
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
				        <th class="first">회사명</th>
				        <td class="left"><%=rs("trade_name")%></td>
				        <th>거래처유형</th>
				        <td class="left"><%=rs("trade_id")%></td>
			          </tr>
				      <tr>
				        <th class="first">그룹명</th>
				        <td class="left"><input name="group_name" type="text" id="group_name" style="width:150px;" value="<%=rs("group_name")%>" onKeyUp="checklength(this,30);"></td>
				        <th>관리그룹</th>
				        <td class="left"><select name="mg_group" id="mg_group" style="width:150px">
				          <option value="1" <% if rs("mg_group") = "1" then %>selected<% end if %>>일반그룹</option>
				          <option value="2" <% if rs("mg_group") = "2" then %>selected<% end if %>>한진그룹</option>
			            </select></td>
			          </tr>
				      <tr>
				        <th class="first">지원회사</th>
				        <td class="left"><input name="support_company" type="text" id="support_company" style="width:150px;" value="<%=rs("support_company")%>" onKeyUp="checklength(this,30);"></td>
				        <th>사용유무</th>
				        <td class="left">
                        <input type="radio" name="use_sw" value="Y" <% if rs("use_sw") = "Y" then %>checked<% end if %> style="width:30px">사용
  						<input type="radio" name="use_sw" value="N" <% if rs("use_sw") = "N" then %>checked<% end if %> style="width:30px">미사용
						</td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="변경" onClick="javascript:frmcheck();" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onClick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="trade_code" value="<%=trade_code%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

