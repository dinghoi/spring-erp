<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
acpt_no = request("acpt_no")

Sql="select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs=DbConn.Execute(Sql)

title_line = "장애내용 변경"

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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=acpt_date%>" );
			});	  
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

				if(document.frm.as_memo.value == "") {
					alert('장애내용을 입력하세요');
					frm.as_memo.focus();
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
				<form action="as_memo_mod_ok.asp" method="post" name="frm">
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
				        <th class="first">접수일자</th>
				        <td class="left"><%=rs("acpt_date")%><input name="old_date" type="hidden" id="old_date" value="<%=old_date%>"></td>
				        <th>사용자</th>
				        <td class="left"><%=rs("acpt_user")%></td>
			          </tr>
				      <tr>
				        <th class="first">회사</th>
				        <td class="left"><%=rs("company")%></td>
				        <th>조직명</th>
				        <td class="left"><%=rs("dept")%></td>
			          </tr>
				      <tr>
				        <th class="first">처리유형</th>
				        <td class="left"><%=rs("as_type")%> / <%=rs("as_process")%></td>
				        <th><span class="first">담당CE</span></th>
				        <td class="left"><%=rs("mg_ce")%>(<%=rs("mg_ce_id")%>)</td>
			          </tr>
				      <tr>
				        <th class="first">장애내용</th>
				        <td colspan="3" class="left"><textarea name="as_memo" rows="5" id="textarea"><%=rs("as_memo")%></textarea></td>
			          </tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
                  	<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
                <input type="hidden" name="acpt_no" value="<%=acpt_no%>" ID="Hidden1">
                <input name="old_acpt_date" type="hidden" id="old_acpt_date" value="<%=old_acpt_date%>">
				</form>
		</div>				
	</body>
</html>

