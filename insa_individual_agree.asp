<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_individual_agree.asp"

curr_date = datevalue(mid(cstr(now()),1,10))

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

Sql = "select * from emp_agree where agree_empno = '"&emp_no&"' ORDER BY agree_year,agree_seq"
Rs.Open Sql, Dbconn, 1

title_line = "근로계약 및 서약 현황(K-won Information Portal 시스템 개발 진행중입니다.)"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
			}
		</script>
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
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_pagree_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_agree.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="12%" >
							<col width="10%" >
							<col width="10%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="8%" >
							<col width="8%" >
							<col width="14%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">년도</th>
								<th scope="col">서약유형</th>
								<th scope="col">회사</th>
								<th scope="col">소속</th>
								<th scope="col">직위</th>
								<th scope="col">직책</th>
                                <th scope="col">직무</th>
                                <th scope="col">입사일</th>
								<th scope="col">서약일자</th>
								<th scope="col">서약기간</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
						
						%>
							<tr>
								<td class="first"><%=rs("agree_year")%></td>
                                <td>
                                <a href="insa_card00.asp?emp_no=<%=rs("agree_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>"><%=rs("agree_id")%></a>
								</td>
                                <td><%=rs("agree_company")%>&nbsp;</td>
                                <td><%=rs("agree_org_name")%>&nbsp;</td>
                                <td><%=rs("agree_job")%>&nbsp;</td>
                                <td><%=rs("agree_position")%>&nbsp;</td>
                                <td><%=rs("agree_jikmu")%>&nbsp;</td>
                                <td><%=rs("agree_in_date")%>&nbsp;</td>
                                <td><%=rs("agree_date")%>&nbsp;</td>
                                <td class="left"><%=rs("agree_from_date")%>&nbsp;∼&nbsp;<%=rs("agree_to_date")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

