<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
mg_ce = Request.form("mg_ce")
seq = Request("seq")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if mg_ce = "" then
	SQL = "select * from memb where grade < '5' and user_name = '" + mg_ce + "' ORDER BY user_name ASC"
 else
	SQL = "select * from memb where grade < '5'  and user_name like '%" + mg_ce + "%' ORDER BY user_name ASC"
end if
Rs.open SQL, Dbconn, 1

title_line = "CE �˻�"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>CE �˻�</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function ce_list(mg_ce,mg_ce_id,grade,emp_company,bonbu,saupbu,team,org_name,reside_place,reside,reside_company,seq)
			{
				opener.document.frm1.mg_ce<%=seq%>.value = mg_ce;
				opener.document.frm1.mg_ce_id<%=seq%>.value = mg_ce_id;
				opener.document.frm1.grade<%=seq%>.value = grade;
				opener.document.frm1.emp_company<%=seq%>.value = emp_company;
				opener.document.frm1.bonbu<%=seq%>.value = bonbu;
				opener.document.frm1.saupbu<%=seq%>.value = saupbu;
				opener.document.frm1.team<%=seq%>.value = team;
				opener.document.frm1.reside<%=seq%>.value = reside;
				opener.document.frm1.reside_place<%=seq%>.value = reside_place;
				opener.document.frm1.reside_company<%=seq%>.value = reside_company;
				opener.document.frm1.org_name<%=seq%>.value = org_name;
				window.close();
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.mg_ce.value =="") {
					alert('CE���� �Է��ϼ���');
					frm.mg_ce.focus();
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
				<form action="ce_search.asp?seq=<%=seq%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>CE���� �Է��ϼ��� </strong>
								<label>
        						<input name="mg_ce" type="text" id="mg_ce" value="<%=mg_ce%>" style="width:150px;text-align:left;ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="25%" >
							<col width="25%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">CE��</th>
								<th scope="col">���̵�</th>
								<th scope="col">����</th>
								<th scope="col">�μ���</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
						%>
							<tr>
								<td class="first"><a href="#" onClick="ce_list('<%=rs("user_name")%>','<%=rs("user_id")%>','<%=rs("user_grade")%>','<%=rs("emp_company")%>','<%=rs("bonbu")%>','<%=rs("saupbu")%>','<%=rs("team")%>','<%=rs("org_name")%>','<%=rs("reside_place")%>','<%=rs("reside")%>','<%=rs("reside_company")%>','<%=seq%>');"><%=rs("user_name")%></a>
                                </td>
								<td><%=rs("user_id")%></td>
								<td><%=rs("user_grade")%></td>
								<td><%=rs("org_name")%></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="4">������ �����ϴ�</td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>        				
	</body>
</html>

