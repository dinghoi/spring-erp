<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

board_gubun = Request("board_gubun")
condi = Request.form("condi")

if board_gubun = "" then
	board_gubun = "0"
end if

if board_gubun = "1" then
	title_line = "�λ����"
  elseif board_gubun = "2" then
  	title_line = "�λ�Խ���"
  elseif board_gubun = "3" then
  	title_line = "�޿�����"
  elseif board_gubun = "4" then
  	title_line = "�ڷ��"
  else
  	title_line = "��ü�Խ���"  
end if

ck_sw = request("ck_sw")
page = request("page")

If ck_sw ="y" Then
	condi = request("condi")
	condi_value = request("condi_value")
Else
	condi = request.form("condi")
	condi_value = request.form("condi_value")
End if

if condi = "" then
	condi = "all"
end if

If condi = "all" Then
	condi_value = ""
End If

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sel_sql = "select * from emp_board"
'sel_sql = "select * from board"

if board_gubun = "0" then
	where_sql = ""
  else
	where_sql = " where board_gubun = '" + board_gubun + "'"
end if

if condi = "all" then
	condi_sql = " "
  else
	if board_gubun = "0" then
		condi_sql = " where " + condi + " like '%" + condi_value  + "%'"
	  else	
  		condi_sql = " and " + condi + " like '%" + condi_value  + "%'"
	end if
end if

order_sql = " order by reg_date desc"

Sql = "select count(*) from emp_board " + where_sql + condi_sql
'Sql = "select count(*) from board " + where_sql + condi_sql
Set RsCount = Dbconn.Execute (sql)

total_record = cint(RsCount(0)) 'Result.RecordCount

IF total_record mod pgsize = 0 THEN
	total_page = int(total_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((total_record / pgsize) + 1)
END IF

sql = sel_sql + where_sql + condi_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

new_date = now() - 14
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ�޿� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
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
					alert ("�Ҽ��� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
            <!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_board_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���� �˻�</dt>
                        <dd>
                            <p>
                              	<input type="radio" name="board_gubun" value="0" <% if board_gubun = "0" then %>checked<% end if %> style="width:30px">�Ѱ�
                              	<input type="radio" name="board_gubun" value="1" <% if board_gubun = "1" then %>checked<% end if %> style="width:30px">�λ����
                              	<input type="radio" name="board_gubun" value="2" <% if board_gubun = "2" then %>checked<% end if %> style="width:30px">�λ�Խ���
                              	<input type="radio" name="board_gubun" value="3" <% if board_gubun = "3" then %>checked<% end if %> style="width:30px">�޿�����
                              	<input type="radio" name="board_gubun" value="4" <% if board_gubun = "4" then %>checked<% end if %> style="width:30px">�ڷ��
                                &nbsp;&nbsp;
                                <strong>���� : </strong>
                                <select name="condi" style="width:100px">
                                  <option value="all" <%If condi = "all" then %>selected<% end if %>>��ü</option>
                                  <option value="board_title" <%If condi = "board_title" then %>selected<% end if %>>����</option>
                                  <option value="board_body" <%If condi = "board_body" then %>selected<% end if %>>����</option>
                                  <option value="reg_name" <%If condi = "reg_name" then %>selected<% end if %>>�ۼ���</option>
                                </select>
								<input name="condi_value" type="text" value="<%=condi_value%>">
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="50%" >
							<col width="10%" >
							<col width="15%" >
							<col width="10%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">�ۼ���</th>
								<th scope="col">�ۼ���</th>
								<th scope="col">��ȸ��</th>
								<th scope="col">÷��</th>
							</tr>
						</thead>
						<tbody>
						<%
    					seq = total_record - ( page - 1 ) * pgsize
						do until rs.eof
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td class="left"><a href="insa_board_view.asp?board_back=<%=board_gubun%>&board_gubun=<%=rs("board_gubun")%>&board_seq=<%=rs("board_seq")%>&page=<%=page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>"><%=rs("board_title")%></a>
                                  <input name="board_seq" type="hidden" id="board_seq" value="<%=Rs("board_seq")%>">
                                  <%	if rs("reg_date") > new_date then 	%>
                                  <img src="image/new.gif" width="24" height="11" border="0">
                                  <%	end if	%>
                                </td>
								<td><%=rs("reg_name")%></td>
								<td><%=rs("reg_date")%></td>
								<td><%=rs("read_cnt")%></td>
								<td>
								<% 
                                If rs("att_file") <> "" Then 
                                    path = "/srv_upload" 
                                %>
                                  <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rs("att_file")%>"><img src="image/att_file.gif" border="0"></a>
                                  <% Else %>
				                    &nbsp;
                                <% End If %>
                                </td>
							</tr>
						<%
							rs.movenext()
  							seq = seq -1
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
				<div id="paging">
					<a href = "insa_board_list.asp?page=<%=first_page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>">[ó��]</a>
                  <% if intstart > 1 then %>
                  	<a href="insa_board_list.asp?page=<%=intstart -1%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>">[����]</a>
                  <% end if %>
                  <% for i = intstart to intend %>
              <% if i = int(page) then %>
                  	<b>[<%=i%>]</b>
                  <% else %>
                  	<a href="insa_board_list.asp?page=<%=i%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                  <% end if %>
                  <% next %>
              <% if 	intend < total_page then %>
              		<a href="insa_board_list.asp?page=<%=intend+1%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_board_list.asp?page=<%=total_page%>&condi=<%=condi%>&condi_value=<%=condi_value%>&ck_sw=<%="y"%>">&nbsp;[������]</a>
        			<%	else %>
    				����&nbsp;������
    			  <% end if %>
					&nbsp;&nbsp;&nbsp;&nbsp;
                <%	if board_gubun <> "0" then %>
                    <a href="insa_board_write.asp?board_gubun=<%=board_gubun%>" class="btnType04">�ۿø���</a>
				<%	end if	%>
				</div>
				<div class="btnRight">
				</div>
				<input type="hidden" name="board_back" value="<%=board_gubun%>">
			</form>
		</div>				
	</div>        				
	</body>
</html>

