<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<!--#include virtual="/include/end_check.asp" -->
<%
cost_year = request("cost_year")
cost_mm = int(request("cost_mm"))
saupbu = request("saupbu")
memo_sw = request("memo_sw")
if cost_mm < 10 then
	cost_mm = "0" + cstr(cost_mm)
  else
  	cost_mm = cstr(cost_mm)
end if
cost_month = cost_year + cost_mm

if position = "사업부장" then
	memo_id = "1"
end if
if position = "본부장" then
	memo_id = "2"
end if

if user_id = "100167" and (saupbu = "공항지원사업부") then
	memo_id = "1"
end if
if user_id = "100031" and (saupbu = "KAL지원사업부" or saupbu = "공항지원사업부") then
	memo_id = "2"
end if

if memo_id = "1" then
	memo_user = "사업부장"
end if
if memo_id = "2" then
	memo_user = "본부장"
end if

title_line = saupbu + " 의견 " + memo_sw

Sql="select * from saupbu_memo where cost_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
Set rs=DbConn.Execute(Sql)

if rs.eof or rs.bof then
	saupbu_memo = ""
	saupbu_reg_name = ""
	saupbu_reg_date = ""
	bonbu_memo = ""
	bonbu_reg_name = ""
	bonbu_reg_date = ""
	end_yn = "N"
  else
	saupbu_memo = rs("saupbu_memo")
	if saupbu_memo = "" or isnull(saupbu_memo) then
		saupbu_memo = rs("saupbu_memo")
	  else
		saupbu_memo = replace(saupbu_memo,chr(10),"<br>")
	end if

	saupbu_reg_name = rs("saupbu_reg_name")
	saupbu_reg_date = rs("saupbu_reg_date")

	bonbu_memo = rs("bonbu_memo")
	if bonbu_memo = "" or isnull(bonbu_memo) then
		bonbu_memo = rs("bonbu_memo")
	  else
		bonbu_memo = replace(bonbu_memo,chr(10),"<br>")
	end if
	bonbu_reg_name = rs("bonbu_reg_name")
	bonbu_reg_date = rs("bonbu_reg_date")
	end_yn = rs("end_yn")
end if
rs.close()

Sql="select * from cost_end where end_month = '"&cost_month&"' and saupbu = '"&saupbu&"'"
Set rs=DbConn.Execute(Sql)
if rs.eof or rs.bof then
	bonbu_yn = "N"
	batch_yn = "N"
  else
	bonbu_yn = rs("bonbu_yn")
	batch_yn = rs("batch_yn")
end if
rs.close()
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 관리 시스템</title>
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
				if(document.frm.memo_id.value =="1") {
					if(document.frm.saupbu_memo.value =="") {
						alert('의견을 입력하세요');
						frm.saupbu_memo.focus();
						return false;}}
				if(document.frm.memo_id.value =="2") {
					if(document.frm.bonbu_memo.value =="") {
						alert('의견을 입력하세요');
						frm.bonbu_memo.focus();
						return false;}}

				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function approve_submit() 
				{
				a=confirm('결재 상신하겠습니까 ?')
				if (a==true) {
					document.frm.action = "saupbu_approve_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="saupbu_memo_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <colgroup>
				      <col width="20%" >
				      <col width="*" >
			        </colgroup>
				    <tbody>
				      <tr>
				        <th class="first">비용년월</th>
				        <td class="left"><%=cost_year%>년<%=cost_mm%>월</td>
			          </tr>
				      <tr>
				        <th class="first">사업부장 의견</th>
				        <td class="left">
					<% if memo_id = "1" and memo_sw = "등록" then	%>
                        <textarea name="saupbu_memo" id="textarea" rows="15"><%=saupbu_memo%></textarea>
					<%   else	%>
					<%   if saupbu_memo = "" or isnull(saupbu_memo) then	%>
                    	&nbsp;
					<%     else	%>
						<%=saupbu_memo%><br><br><%=saupbu_reg_name%>(<%=saupbu_reg_date%>)
					<%   end if	%>
                    <% end if	%>
                        </td>
			          </tr>
				      <tr>
				        <th class="first">본부장 의견</th>
				        <td class="left">
					<% if memo_id = "2" and memo_sw = "등록" then	%>
                        <textarea name="bonbu_memo" id="textarea" rows="10"><%=bonbu_memo%></textarea>
					<%   else	%>
					<%   if bonbu_memo = "" or isnull(bonbu_memo) then	%>
                    	&nbsp;
					<%     else	%>
						<%=bonbu_memo%><br><br><%=bonbu_reg_name%>(<%=bonbu_reg_date%>)
                    <%   end if	%>
					<% end if	%>
                        </td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
			<% if memo_sw = "등록" then	%>
				<%	if end_yn <> "Y" then	%>
					<% if (memo_id = "1" and batch_yn = "N") or (memo_id = "2" and bonbu_yn = "N") then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="승인" onclick="javascript:approve_submit();"  NAME="Button1"></span>
        			<% end if	%>
				<%	end if	%>
                    <span class="btnType01"><input type="button" value="담기" onclick="javascript:goAction();"></span>
			<% end if	%>
                </div>
				<br>
                    <input type="hidden" name="cost_month" value="<%=cost_month%>" ID="Hidden1">
                    <input type="hidden" name="saupbu" value="<%=saupbu%>" ID="Hidden1">
                    <input type="hidden" name="memo_id" value="<%=memo_id%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

