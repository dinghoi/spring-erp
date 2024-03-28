<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

company=Request("company")
dept=Request("dept")
acpt_user=Request("acpt_user")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from as_acpt where company='"+company+"' and dept='"+dept+"' and acpt_user='"+acpt_user+"' order by acpt_date desc"
Rs.Open Sql, Dbconn, 1

title_line = "A/S History"
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
		<script src="/java/ui.js" type="text/javascript"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.condi.value == "") {
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
			<div id="container">
				<h3 class="tit"><%=company%>&nbsp;<%=dept%>&nbsp;<%=acpt_user%>님&nbsp;<%=title_line%></h3>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="7%" >
							<col width="4%" >
							<col width="5%" >
							<col width="5%" >
							<col width="7%" >
							<col width="27%" >
							<col width="7%" >
							<col width="7%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">순번</th>
								<th scope="col">접수일자</th>
								<th scope="col">상태</th>
								<th scope="col">접수자</th>
								<th scope="col">담당CE</th>
								<th scope="col">장애장비</th>
								<th scope="col">장애내용</th>
								<th scope="col">처리일자</th>
								<th scope="col">처리유형</th>
								<th scope="col">처리내용</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof	
							i = i + 1											
							as_memo = replace(rs("as_memo"),chr(34),chr(39))
							view_memo = as_memo
							if len(as_memo) > 30 then
								view_memo = mid(as_memo,1,30) + ".."
							end if
							if isnull(rs("as_history")) then
								view_history = ""
							  else
								as_history = replace(rs("as_history"),chr(34),chr(39))
								view_history = as_history
							end if
							if len(as_history) > 30 then
								view_history = mid(as_history,1,30) + ".."
							end if
							if as_process = "입고" then
								pro_date = rs("in_date")
							  elseif as_process = "완료" or as_process = "취소" then
							  	pro_date = rs("visit_date")
							  else
							  	pro_date = "미처리"
							end if							  	
							%>
							<tr>
								<td class="first"><%=i%></td>
								<td><%=mid(rs("acpt_date"),1,10)%></td>
								<td><%=rs("as_process")%></td>
								<td><%=rs("acpt_man")%></td>
								<td><%=rs("mg_ce")%></td>
								<td><%=rs("as_device")%></td>
							  	<td class="left"><p style="cursor:pointer"><span title="<%=as_memo%>"><%=view_memo%></span></p></td>
								<td><%=pro_date%></td>
								<td><%=rs("as_type")%></td>
							  	<td class="left"><p style="cursor:pointer"><span title="<%=as_history%>"><%=view_history%></span></p></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
		</div>				
	</div>        				
	</body>
</html>

