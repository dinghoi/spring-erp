<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
dim pg_cnt

Page=Request("page")
pg_cnt=cint(Request("pg_cnt"))

pgsize = 10 ' ȭ�� �� ������ 

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

Sql = "select count(*) from sign_msg where recv_id = '"&user_id&"'"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

'sql = "select * from sign_msg where recv_id = '"&user_id&"'and sign_yn = 'N'"&order_sql&" limit "& stpage & "," &pgsize 
sql = "select * from sign_msg where recv_id = '"&user_id&"' "&order_sql&" limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "���� ���ڰ��� �� ( " + user_name + " " + user_grade + " )"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���� ���� �ý���</title>
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
				<form action="sign_process_mg.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
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
								<th rowspan="2" class="first" scope="col">����</th>
								<th rowspan="2" scope="col"><img src="image/close_icon.gif" width="16" height="13"></th>
						    	<th rowspan="2" scope="col">��������</th>
								<th rowspan="2" scope="col">������ȣ</th>
								<th rowspan="2" scope="col">�� �� �� ��</th>
								<th rowspan="2" scope="col">��û�ð�</th>
								<th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">��������</th>
							</tr>
							<tr>
							  <th scope="col" style=" border-left:1px solid #e3e3e3;">����</th>
							  <th scope="col">�������</th>
							  <th scope="col">������</th>
							  <th scope="col">����</th>
					      </tr>
						</thead>
						<tbody>
						<%
    					seq = total_record - ( page - 1 ) * pgsize + 1
						do until rs.eof 
							sign_date = mid(rs("paper_no"),1,10)
							sign_seq = right(rs("paper_no"),3)
							sql="select * from sign_process where sign_date='"&sign_date&"' and sign_seq = '"&sign_seq&"'"
							set rs_sign=dbconn.execute(sql)
							sign_month = rs_sign("sign_month")
							
							if rs_sign("team_sign") = "E" then
								team_sign_view = "����"
							  elseif rs_sign("team_sign") = "C" then
								team_sign_view = "�ݷ�"
							  else
								team_sign_view = ""
							end if
							if rs_sign("saupbu_sign") = "E" then
								saupbu_sign_view = "����"
							  elseif rs_sign("saupbu_sign") = "C" then
								saupbu_sign_view = "�ݷ�"
							  else
								saupbu_sign_view = ""
							end if
							if rs_sign("bonbu_sign") = "E" then
								bonbu_sign_view = "����"
							  elseif rs_sign("bonbu_sign") = "C" then
								bonbu_sign_view = "�ݷ�"
							  else
								bonbu_sign_view = ""
							end if
							if rs_sign("ceo_sign") = "E" then
								ceo_sign_view = "����"
							  elseif rs_sign("ceo_sign") = "C" then
								ceo_sign_view = "�ݷ�"
							  else
								ceo_sign_view = ""
							end if
							rs_sign.close()
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td>
						<% if rs("read_yn") = "Y" then	%>
								<img src="image/open_icon.gif" width="16" height="14">
						<%   else	%>
								<img src="image/close_icon.gif" width="16" height="13">
                        <% end if	%>
                                </td>
								<td><%=rs("send_name")%>&nbsp;(<%=rs("send_id")%>)</td>
								<td><%=rs("paper_no")%></td>
								<td class="left">
                  		<% if rs("sign_yn") <> "C" then	%>     
                                <a href="#" onClick="pop_Window('cost_sign.asp?sign_date=<%=sign_date%>&sign_seq=<%=sign_seq%>&sign_month=<%=sign_month%>&msg_seq=<%=rs("msg_seq")%>&sign_yn=<%=rs("sign_yn")%>&sign_head=<%=rs("sign_head")%>','cost_sign_pop','scrollbars=yes,width=1150,height=650')"><%=rs("sign_head")%></a>
						<%   else	%>
								<%=rs("sign_head")%>
                        <% end if %>
                                </td>
								<td><%=rs("reg_date")%></td>
								<td><%=team_sign_view%>&nbsp;</td>
								<td><%=saupbu_sign_view%>&nbsp;</td>
								<td><%=bonbu_sign_view%>&nbsp;</td>
								<td><%=ceo_sign_view%>&nbsp;</td>
							</tr>
						<%
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
                        <a href="sign_process_mg.asp?page=<%=first_page%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="sign_process_mg.asp?page=<%=intstart -1%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="sign_process_mg.asp?page=<%=i%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="sign_process_mg.asp?page=<%=intend+1%>">[����]</a> <a href="sign_process_mg.asp?page=<%=total_page%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
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

