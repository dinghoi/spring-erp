<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
gubun = Request("gubun")

trade_name = Request.form("trade_name")
if gubun = "" or isnull(gubun) then
	gubun = Request.form("gubun")
end if
'response.write(gubun)
Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if gubun = "1" or gubun = "5" then
   if trade_name = "" then
	   SQL = "select * from trade where (trade_id = '일반' or trade_id = '매출') and trade_name = '%" + trade_name + "%' ORDER BY trade_name ASC"
    else
	   SQL = "select * from trade where (trade_id = '일반' or trade_id = '매출') and trade_name like '%" + trade_name + "%' ORDER BY trade_name ASC"
   end if
   Rs.open SQL, Dbconn, 1
end if

if gubun = "2" or gubun = "3" or gubun = "4"then
   if trade_name = "" then
	   SQL = "select * from trade where trade_name = '" + trade_name + "' ORDER BY trade_name ASC"
    else
	   SQL = "select * from trade where trade_name like '%" + trade_name + "%' ORDER BY trade_name ASC"
   end if
   Rs.open SQL, Dbconn, 1
end if
'Response.write SQL
title_line = "거래처 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>거래처 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function trade_list(trade_code,trade_name,trade_no,trade_person,trade_email,group_name)
			{
				if(document.frm.gubun.value =="1") {
					opener.document.frm.org_reside_company.value = trade_name;
					opener.document.frm.org_cost_group.value = group_name;
//					opener.document.frm.trade_no.value = trade_no;
//					opener.document.frm.trade_person.value = trade_person;
//					opener.document.frm.trade_email.value = trade_email;
					window.close();
				}
				if(document.frm.gubun.value =="2") {
					opener.document.frm.cost_company.value = trade_name;
					window.close();
				}
				if(document.frm.gubun.value =="3") {
					opener.document.frm.customer.value = trade_name;
					opener.document.frm.customer_no.value = trade_no;
					window.close();
				}
				if(document.frm.gubun.value =="4") {
					opener.document.frm.company.value = trade_name;
					window.close();
				}
				if(document.frm.gubun.value =="5") {
					opener.document.frm.emp_reside_company.value = trade_name;
					opener.document.frm.cost_group.value = group_name;
//					opener.document.frm.trade_no.value = trade_no;
//					opener.document.frm.trade_person.value = trade_person;
//					opener.document.frm.trade_email.value = trade_email;
					window.close();
				}
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.trade_name.value =="") {
					alert('거래처명을 입력하세요');
					frm.trade_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_trade_search.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>거래처명을 입력하세요 </strong>
								<label>
        						<input name="trade_name" type="text" id="trade_name" value="<%=trade_name%>" style="width:150px;text-align:left;ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="25%" >
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">거래처명</th>
								<th scope="col">그룹명</th>
								<th scope="col">사업자번호</th>
								<th scope="col">담당자</th>
								<th scope="col">이메일</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
							trade_code = rs("trade_code")
							trade_name = rs("trade_name")
							trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + mid(rs("trade_no"),6)
							trade_person = rs("trade_person")
							trade_email = rs("trade_email")
							group_name = rs("group_name")
							if Trim(group_name) = "" or isnull(group_name) then
							    group_name = trade_name
							end if
						%>
							<tr>
								<td class="first">
									<a href="#" onClick="trade_list('<%=trade_code%>','<%=trade_name%>','<%=rs("trade_no")%>','<%=trade_person %>','<%=trade_email%>','<%=group_name%>');"><%=rs("trade_name")%></a>
                </td>
								<td><%=group_name%>&nbsp;</td>
                                <td><%=trade_no%>&nbsp;</td>
								<td><%=trade_person%>&nbsp;</td>
								<td><%=trade_email%>&nbsp;</td>
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
				<input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
			</form>
		</div>
	</body>
</html>

