<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
mg_ce_id = Request.form("mg_ce_id")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

SQL = "select * from memb where user_id = '" + mg_ce_id + "'"
set rs=dbconn.execute(sql)

if mg_ce_id = "" or mg_ce_id = null then
	use_msg = "아이디를 입력하세요 !!!!"
	use_ok = "N"
 else
	If Rs.Eof or Rs.Bof Then
		use_msg = "사용가능한 아이디입니다. 사용하시겠습니까?"
		use_ok = "Y"
	 else
		use_msg = "이미 사용하고 있는 아이디입니다 !!!!"
		use_ok = "N"
	End if
end if

title_line = "아이디 중복 Check"
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
			function id_move(mg_ce_id)
			{
				opener.document.frm.user_id.value = mg_ce_id;
				window.close();
			
			}
			
			function chkfrm(){
				if(document.frm.mg_ce_id.value =="") {
					alert('아이디를 입력하세요');
					frm.mg_ce_id.focus();
					return false;}
				document.frm.submit ();
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="ce_id_check.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="25%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">아이디</th>
								<td class="left"><input name="mg_ce_id" type="text" id="mg_ce_id" value="<%=mg_ce_id%>" style="width:120px" onKeyUp="checklength(this,15)"><a href="#" class="btnType03" onclick="javascript:chkfrm();" >중복확인</a></td>
							</tr>
							<tr>
								<th class="first">가능여부</th>
								<td class="left"><%=use_msg%>&nbsp;
								  <% if use_ok = "Y" then %>
                                    <a href="#" class="btnType03" onClick="id_move('<%=mg_ce_id%>');">사용</a>
                                  <% end if %>
                                </td>
							</tr>
						</tbody>
					</table>
				</div>
				</form>
		</div>				
	</body>
</html>

