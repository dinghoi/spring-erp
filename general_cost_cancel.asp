<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
slip_date = request("slip_date")
slip_seq = request("slip_seq")

Sql="select * from general_cost where slip_date = '"&slip_date&"' and slip_seq = '"&slip_seq&"'"
Set rs=DbConn.Execute(Sql)

org_name = rs("org_name")
account = rs("account") + "-" + rs("account_item")
sign_no = rs("sign_no")
pay_method = rs("pay_method")
price = rs("price")
company = rs("company")
vat_yn = rs("vat_yn")
pay_yn = rs("pay_yn")
customer = rs("customer")
emp_name = rs("emp_name")
emp_no = rs("emp_no")
slip_memo = rs("slip_memo")
end_yn = rs("end_yn")
cancel_yn = rs("cancel_yn")
reg_id = rs("reg_id")
reg_date = rs("reg_date")
mod_id = rs("mod_id")
mod_date = rs("mod_date")
mod_user = rs("mod_user")
rs.close()

if end_yn = "Y" then
	end_view = "마감"
  else
  	end_view = "진행"
end if

title_line = "일반경비 지급 변경"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>관리회계시스템</title>
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
				<form action="general_cost_cancel_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="13%" >
				      <col width="37%" >
				      <col width="13%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">발생일자</th>
				        <td class="left"><%=slip_date%>&nbsp;
				        <input name="slip_date" type="hidden" value="<%=slip_date%>">
				        <input name="slip_seq" type="hidden" value="<%=slip_seq%>">
                        </td>
				        <th>소속</th>
				        <td class="left"><%=org_name%></td>
			          </tr>
				      <tr>
				        <th class="first">사용자</th>
				        <td class="left"><%=emp_name%>&nbsp;(&nbsp;<%=emp_no%>&nbsp;)</td>
				        <th>비용항목</th>
				        <td class="left"><%=account%>&nbsp;<%=account_item%>&nbsp;</td>
			          </tr>
				      <tr>
				        <th class="first">사용구분/금액</th>
				        <td class="left"><%=pay_method%><input name="pay_method" type="hidden" value="<%=pay_method%>"></td>
				        <th>사용회사</th>
				        <td class="left"><%=company%></td>
			          </tr>
				      <tr>
				        <th class="first">사용처</th>
				        <td class="left"><%=customer%></td>
				        <th>전자결재NO</th>
				        <td class="left"><%=sign_no%></td>
			          </tr>
				      <tr>
				        <th class="first">정산여부</th>
				        <td class="left"><%=pay_yn%></td>
				        <th>비고</th>
				        <td class="left"><%=slip_memo%></td>
			          </tr>
    				  <tr>
						<th class="first">취소여부</th>
						<td class="left">
						<input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:30px" ID="Radio1">취소           
                        <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:30px" ID="Radio2">지급
						</td>
                        <th>마감여부</th>
						<td class="left"><%=end_view%></td>
					</tr>
					<tr>
						<th class="first">등록정보</th>
						<td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
                    	<th>변경정보</th>
						<td class="left"><%=mod_user%>&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
					</tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	if end_yn = "N" or end_yn = "C" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="end_yn" value="<%=end_yn%>" ID="Hidden1">
				<input type="hidden" name="end_date" value="<%=end_date%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

