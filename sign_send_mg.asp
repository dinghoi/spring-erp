<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
dim pg_cnt

Page=Request("page")
pg_cnt=cint(Request("pg_cnt"))

pgsize = 10 ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY reg_date ASC"

Sql = "select count(*) from sign_process where reg_id = '"&user_id&"'"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from sign_process where reg_id = '"&user_id&"' "&order_sql&" limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "보낸 전자결재 함 ( " + user_name + " " + user_grade + " )"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>전자 결재 시스템</title>
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
				return true;
			}
			function auto_submit()
			{
			 window.setTimeout("frmcheck()", 300000);
			 return true;
			}
		</script>

	</head>
	<body onload="auto_submit()">
		<div id="wrap">			
			<!--#include virtual = "/include/sign_header.asp" -->
			<!--#include virtual = "/include/sign_menu.asp" -->
			<div id="container">
				<h3 class="tit" style="color:#F60"><%=title_line%></h3>
				<br>	
				<form action="sign_send_mg.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="10%" >
							<col width="10%" >
							<col width="*" >
							<col width="15%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
						</colgroup>
						<thead>
							<tr>
								<th rowspan="2" class="first" scope="col">순번</th>
						    	<th rowspan="2" scope="col">결재상신자</th>
								<th rowspan="2" scope="col">문서번호</th>
								<th rowspan="2" scope="col">결 재 내 용</th>
								<th rowspan="2" scope="col">요청시간</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">결재진행</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">팀장</th>
							  <th scope="col">사업부장</th>
							  <th scope="col">본부장</th>
							  <th scope="col">사장</th>
					      </tr>
						</thead>
						<tbody>
						<%
    					seq = total_record - ( page - 1 ) * pgsize + 1
						do until rs.eof
							paper_no = cstr(rs("sign_date")) + "-" + cstr(rs("sign_seq")) 
							sql="select * from sign_msg where paper_no='"&paper_no&"'"
							set rs_msg=dbconn.execute(sql)
							sign_month = rs("sign_month")
							sign_yn = "Y"
							if rs("team_sign") = "E" then
								team_sign_view = "결재"
							  elseif rs("team_sign") = "C" then
								team_sign_view = "반려"
							  else
								team_sign_view = ""
							end if
							if rs("saupbu_sign") = "E" then
								saupbu_sign_view = "결재"
							  elseif rs("saupbu_sign") = "C" then
								saupbu_sign_view = "반려"
							  else
								saupbu_sign_view = ""
							end if
							if rs("bonbu_sign") = "E" then
								bonbu_sign_view = "결재"
							  elseif rs("bonbu_sign") = "C" then
								bonbu_sign_view = "반려"
							  else
								bonbu_sign_view = ""
							end if
							if rs("ceo_sign") = "E" then
								ceo_sign_view = "결재"
							  elseif rs("ceo_sign") = "C" then
								ceo_sign_view = "반려"
							  else
								ceo_sign_view = ""
							end if
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=rs("reg_user")%>&nbsp;(<%=rs("reg_id")%>)</td>
								<td><%=paper_no%></td>
								<td class="left"><a href="#" onClick="pop_Window('cost_sign.asp?sign_date=<%=rs("sign_date")%>&sign_seq=<%=rs("sign_seq")%>&sign_month=<%=sign_month%>&msg_seq=<%=rs_msg("msg_seq")%>&sign_yn=<%=sign_yn%>','cost_sign_pop','scrollbars=yes,width=1150,height=500')"><%=rs_msg("sign_head")%></a></td>
								<td><%=rs("reg_date")%></td>
								<td><%=team_sign_view%>&nbsp;</td>
								<td><%=saupbu_sign_view%>&nbsp;</td>
								<td><%=bonbu_sign_view%>&nbsp;</td>
								<td><%=ceo_sign_view%>&nbsp;</td>
							</tr>
						<%
							rs_msg.close()
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				</div>
				<%
                intstart = (int((page-1)/10)*10) + 1
                intend = intstart + 9
                first_page = 1
                
                if intend > total_page then
                    intend = total_page
                end if
                %>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="sign_send_mg.asp?page=<%=first_page%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="sign_send_mg.asp?page=<%=intstart -1%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sign_send_mg.asp?page=<%=i%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="sign_send_mg.asp?page=<%=intend+1%>">[다음]</a> <a href="sign_send_mg.asp?page=<%=total_page%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%"><strong><%=now()%></strong></td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

