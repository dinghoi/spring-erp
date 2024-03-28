<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/srvmg_dbcon.asp" -->
<!--#include virtual="/include/srvmg_user.asp" -->
<%
bonbu = request("bonbu")
saupbu = request("saupbu")
use_man = Request.form("use_man")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if mg_ce = "" then
	SQL = "select * from memb where mg_group = '"&mg_group&"' and bonbu = '"&mg_ce&"' and user_name = '"&mg_ce&"' and user_name = '"&mg_ce&"' ORDER BY user_name ASC"
 else
	SQL = "select * from memb where mg_group = '" + mg_group + "' and user_name like '%" + mg_ce + "%' ORDER BY user_name ASC"
end if
Rs.open SQL, Dbconn, 1

title_line = "CE 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>CE 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function ce_list(mg_ce,mg_ce_id,grade,belong,seq)
			{
				if (seq == '1' ) {
					opener.document.frm.mg_ce1.value = mg_ce;
					opener.document.frm.mg_ce_id1.value = mg_ce_id;
					opener.document.frm.grade1.value = grade;
					opener.document.frm.belong1.value = belong;
				}
				if (seq == '2' ) {
					opener.document.frm.mg_ce2.value = mg_ce;
					opener.document.frm.mg_ce_id2.value = mg_ce_id;
					opener.document.frm.grade2.value = grade;
					opener.document.frm.belong2.value = belong;
				}
				if (seq == '3' ) {
					opener.document.frm.mg_ce3.value = mg_ce;
					opener.document.frm.mg_ce_id3.value = mg_ce_id;
					opener.document.frm.grade3.value = grade;
					opener.document.frm.belong3.value = belong;
				}
				if (seq == '4' ) {
					opener.document.frm.mg_ce4.value = mg_ce;
					opener.document.frm.mg_ce_id4.value = mg_ce_id;
					opener.document.frm.grade4.value = grade;
					opener.document.frm.belong4.value = belong;
				}
				if (seq == '5' ) {
					opener.document.frm.mg_ce5.value = mg_ce;
					opener.document.frm.mg_ce_id5.value = mg_ce_id;
					opener.document.frm.grade5.value = grade;
					opener.document.frm.belong5.value = belong;
				}
				if (seq == '6' ) {
					opener.document.frm.mg_ce6.value = mg_ce;
					opener.document.frm.mg_ce_id6.value = mg_ce_id;
					opener.document.frm.grade6.value = grade;
					opener.document.frm.belong6.value = belong;
				}
				if (seq == '7' ) {
					opener.document.frm.mg_ce7.value = mg_ce;
					opener.document.frm.mg_ce_id7.value = mg_ce_id;
					opener.document.frm.grade7.value = grade;
					opener.document.frm.belong7.value = belong;
				}
				if (seq == '8' ) {
					opener.document.frm.mg_ce8.value = mg_ce;
					opener.document.frm.mg_ce_id8.value = mg_ce_id;
					opener.document.frm.grade8.value = grade;
					opener.document.frm.belong8.value = belong;
				}
				if (seq == '9' ) {
					opener.document.frm.mg_ce9.value = mg_ce;
					opener.document.frm.mg_ce_id9.value = mg_ce_id;
					opener.document.frm.grade9.value = grade;
					opener.document.frm.belong9.value = belong;
				}
				if (seq == '10' ) {
					opener.document.frm.mg_ce10.value = mg_ce;
					opener.document.frm.mg_ce_id10.value = mg_ce_id;
					opener.document.frm.grade10.value = grade;
					opener.document.frm.belong10.value = belong;
				}
				window.close();
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.mg_ce.value =="") {
					alert('CE명을 입력하세요');
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
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>CE명을 입력하세요 </strong>
								<label>
        						<input name="mg_ce" type="text" id="mg_ce" value="<%=mg_ce%>" style="width:150px; text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
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
								<th class="first" scope="col">CE명</th>
								<th scope="col">아이디</th>
								<th scope="col">직급</th>
								<th scope="col">소속 / 상주처</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
						%>
							<tr>
								<td class="first"><a href="#" onClick="ce_list('<%=rs("user_name")%>','<%=rs("user_id")%>','<%=rs("user_grade")%>','<%=rs("belong")%>','<%=seq%>');"><%=rs("user_name")%></a>
                                </td>
								<td><%=rs("user_id")%></td>
								<td><%=rs("user_grade")%></td>
								<td><%=rs("belong")%> / <%=rs("reside_place")%></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="4">내역이 없습니다</td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
			</div>				
	</div>        				
	</form>
	</body>
</html>

