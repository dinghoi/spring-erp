<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
account_name = request("account_name")
account_group = request("account_group")
account_seq = request("account_seq")

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "SELECT * FROM ACCOUNT ORDER BY account_group, account_name ASC"
Rs.Open Sql, Dbconn, 1

if account_name = "" then
	account_name = rs("account_name")
	account_group = rs("account_group")
	account_seq = rs("account_seq")
end if

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
			function getPageCode(){
				return "3 1";
			}
		</script>
	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/account_header.asp" -->
			<!--#include virtual = "/include/account_code_menu.asp" -->
			<div id="container">
				<h3 class="tit">계정과목 관리</h3>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="30%" height="356" valign="top"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="25%" >
				          <col width="*" >
				          <col width="25%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">계정그룹</th>
				            <th scope="col">계정과목</th>
				            <th scope="col">비용코드</th>
			              </tr>
			            </thead>
			            <tbody>
				        <%
						do until rs.eof
						%>
				        <tr>
				          <td class="first"><%=rs("account_group")%></td>
				          <td><a href="account_code_mg.asp?account_name=<%=rs("account_name")%>&account_group=<%=rs("account_group")%>&account_seq=<%=rs("account_seq")%>&u_type=<%="U"%>"><%=rs("account_name")%></a></td>
				          <td><%=rs("account_code")%>&nbsp;</td>
			            </tr>
				        <%
							rs.movenext()
						loop
						rs.close()
						%>
			            </tbody>
			          </table>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="68%" valign="top" height="356"><table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="15%" >
				          <col width="*" >
				          <col width="10%" >
				          <col width="10%" >
				          <col width="10%" >
				          <col width="20%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">계정과목</th>
				            <th scope="col">적요</th>
				            <th scope="col">비용사용</th>
				            <th scope="col">등록인</th>
				            <th scope="col">수정인</th>
				            <th scope="col">수정일</th>
			              </tr>
			            </thead>
			            <tbody>
				        <%
						Sql = "select * from account_item where account_name = '"+ account_name +"' ORDER BY account_name ASC"
						Rs.Open Sql, Dbconn, 1

						do until rs.eof
							item_seq = rs("item_seq")
							if rs("cost_yn") = "Y" then
								cost_view = "기본사용"
							  elseif rs("cost_yn") = "C" then
								cost_view = "확장사용"
							  else
							  	cost_view = "미사용"
							end if
						%>
				        <tr>
				          <td class="first"><%=rs("account_name")%></td>
				          <td>
						<% if rs("account_item") = "" or isnull(rs("account_item")) then %>
							&nbsp;
                        <%   else %>
						  <a href="#" onClick="pop_Window('account_code_add.asp?account_group=<%=rs("account_group")%>&account_seq=<%=rs("account_seq")%>&item_seq=<%=rs("item_seq")%>&u_type=<%="U"%>','계정과목등록','scrollbars=yes,width=500,height=250')"><%=rs("account_item")%></a>
                        <% end if %>
                          </td>
				          <td><%=cost_view%></td>
				          <td><%=rs("reg_user")%></td>
				          <td><%=rs("mod_user")%>&nbsp;</td>
				          <td><%=rs("mod_date")%>&nbsp;</td>
			            </tr>
				        <%
							rs.movenext()
						loop
						rs.close()
						%>
			            </tbody>
			          </table>
                        <table width="100%" border="0" cellpadding="0" cellspacing="0">
                          <tr>
                            <td>
                            <div class="btnRight">
                            <a href="#" class="btnType04" onClick="pop_Window('account_code_add.asp?account_group=<%=account_group%>&account_seq=<%=account_seq%>&account_name=<%=account_name%>','계정과목등록','scrollbars=yes,width=500,height=250')">계정과목 등록</a>
                            </div>                  
                            </td>
                          </tr>
                          </table>
                     </td>
			        </tr>
			      </table>
                </div>
			</div>				
	</div>        				
	</body>
</html>

