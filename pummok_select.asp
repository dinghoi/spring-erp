<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
code_ary = request("code_ary")
srv_type = Request.Form("srv_type")
Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if srv_type = "" then
	sql = "select * from pummok_code where srv_type = '" + srv_type + "'"
  else
	sql = "select * from pummok_code where srv_type like '%" + srv_type + "%' ORDER BY pummok_name ASC"
end if
Rs.Open Sql, Dbconn, 1

title_line = "ǰ�� �˻�"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>ǰ�� �˻�</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			function frmcheck1 () {
//				if (chkfrm1()) {
				document.frm1.submit ();
//				}
			}			
			
			function chkfrm() {
				if(document.frm.srv_type.value == "" || document.frm.srv_type.value == " ") {
					alert('���������� �Է��ϼ���');
					frm.srv_type.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="pummok_select.asp?code_ary=<%=code_ary%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>���������� �Է��ϼ��� </strong>
								<label>
        						<input name="srv_type" type="text" id="srv_type" value="<%=srv_type%>" style="width:150px; text-align:left; ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<form action="pummok_select_ok.asp" method="post" name="frm1">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="23%" >
							<col width="23%" >
							<col width="23%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">��������</th>
								<th scope="col">ǰ���</th>
								<th scope="col">�԰�</th>
								<th scope="col">���</th>
							</tr>
						</thead>
						<tbody>
						<%
							i = 0
							do until rs.eof or rs.bof
								i = i + 1
							%>
							<tr>
								<td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="<%=rs("pummok_code")%>"></td>
								<td><%=rs("srv_type")%></td>
								<td><%=rs("pummok_name")%></td>
								<td><%=rs("standard")%></td>
								<td><%=rs("pummok_memo")%>&nbsp;</td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
						<%
						  if i = 0 then
						%>
							<tr>
								<td class="first" colspan="5">������ �����ϴ�</td>
							</tr>
                        <%
						end if
						%>
							<tr>
								<td class="first; left" colspan="5"><span class="btnType04"><input type="button" value="����" onclick="javascript:frmcheck1();"></span></td>
							</tr>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="code_ary" value="<%=code_ary%>">
				</form>
		</div>        				
	</body>
</html>

