<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
gubun = request("gubun")
slip_gubun = Request("slip_gubun")
if slip_gubun = "" then
	slip_gubun = Request.form("slip_gubun")
end if

if slip_gubun = "비용" then
	SQL = "select concat(account_name,'-',account_item) as etc_name from account_item where cost_yn = 'Y' ORDER BY account_name ASC"
 else
	SQL = "select etc_name from etc_code where type_name = '"&slip_gubun&"' ORDER BY etc_name ASC"
end if
Rs.open SQL, Dbconn, 1
Response.write  SQL
title_line = "비용 유형 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>비용 유형 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function account_list(gubun,slip_gubun,account_view,account,account_item)
			{
				if(gubun =="영업")
					{
					opener.document.frm.slip_gubun.value = slip_gubun;
					opener.document.frm.account.value = account;
					opener.document.frm.account_item.value = account_item;
					opener.document.frm.account_view.value = account_view;
					window.close();
					}
				else
					{
					opener.document.frm.slip_gubun.value = slip_gubun;
					opener.document.frm.account.value = account;
					opener.document.frm.account_item.value = account_item;
					opener.document.frm.account_view.value = account_view;
					window.close();
					}
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if(document.frm.slip_gubun.value =="") {
					alert('비용유형을 선택하세요!!!');
					frm.slip_gubun.focus();
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
				<form action="tax_bill_account_search.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>비용유형 : </strong>
							<label>
                            <select name="slip_gubun" id="slip_gubun" style="width:120px">
                              <option value=''>선택</option>
                              <%
                                Sql="select * from type_code where etc_seq = '4' and etc_id = 'T' order by type_name asc"
                                rs_etc.Open Sql, Dbconn, 1
                                do until rs_etc.eof
                                %>
                              <option value='<%=rs_etc("type_name")%>' <%If slip_gubun = rs_etc("type_name") then %>selected<% end if %>><%=rs_etc("type_name")%></option>
                              <%
                                    rs_etc.movenext()
                                loop
                                rs_etc.close()
                                %>
                              <option value='비용' <%If slip_gubun = "비용" then %>selected<% end if %>>비용</option>
                            </select>
							</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">비용유형</th>
								<th scope="col">세부 비용</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
							if slip_gubun = "비용" then
								accountitem = rs("etc_name")
								i=instr(1,accountitem,"-")
								account = mid(accountitem,1,i-1)
								account_item = mid(accountitem,i+1)
							  else
							  	account = rs("etc_name")
								account_item = rs("etc_name")
							end if
						%>
							<tr>
								<td class="first"><%=slip_gubun%></td>
								<td class="left"><a href="#" onClick="account_list('<%=gubun%>','<%=slip_gubun%>','<%=rs("etc_name")%>','<%=account%>','<%=account_item%>');"><%=rs("etc_name")%></a></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="2">내역이 없습니다</td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
				<br>
			</form>
		</div>
	</body>
</html>

