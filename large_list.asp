<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim field_check
Dim field_view
Dim win_sw

win_sw = "close"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	field_check=Request("field_check")
	field_view=Request("field_view")
	company=Request("company")
  else
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	company=Request.form("company")
End if

If company = "" Then
	field_check = "total"
	company = "��ü"
End If

If field_check = "total" Then
	field_view = ""
End If

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
'  else
'  	page = cint(page)
'	start_page = int(page/setsize)
'	if start_page = (page/setsize) then
'		strat_page = page - setsize + 1
'	  else
'	  	start_page = int(page/setsize)*setsize + 1
'	end if
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_into = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' ���Ǻ� ��ȸ.........
base_sql = "select * from large_acpt "

where_sql = "where upload_ok = 'N' "

if field_check <> "total" then
	field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
  else
  	field_sql = " "
end if

if company = "��ü" then
	com_sql = " "
  else
  	com_sql = " and (company = '" + company + "') "
end if

order_sql = " ORDER BY paper_no, sido, gugun, dong, addr ASC"

Sql = "SELECT count(*) FROM large_acpt " + where_sql + com_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + where_sql + com_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = "�뷮�� ���ε� ��� ��Ȳ"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "1 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.field_check.value == "") {
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/large_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="large_list.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>					
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
                                <label>
								<strong>ȸ��</strong>
								<%
                                Sql="select * from trade where use_sw = 'Y' and mg_group = '"+mg_group+"' order by trade_name asc"
                                rs_etc.Open Sql, Dbconn, 1
                                %>
                                <select name="company" id="company">
 									<option value="��ü">��ü</option> 
          					<% 
								While not rs_etc.eof 
							%>
          							<option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = company  then %>selected<% end if %>><%=rs_etc("trade_name")%></option>
          					<%
									rs_etc.movenext()  
								Wend 
								rs_etc.Close()
							%>
                                </select>
								</label>
								<strong>�ʵ�����</strong>
                                <select name="field_check" id="field_check" style="width:70px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                    <option value="paper_no" <% if field_check = "paper_no" then %>selected<% end if %>>������ȣ</option>
                                    <option value="mg_ce" <% if field_check = "mg_ce" then %>selected<% end if %>>���CE</option>
                                    <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>�õ�</option>
                                    <option value="gugun" <% if field_check = "gugun" then %>selected<% end if %>>����</option>
                                    <option value="dong" <% if field_check = "dong" then %>selected<% end if %>>��</option>
                                </select>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:80px; text-align:left" >
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="�˻�"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="3%" >
							<col width="6%" >
							<col width="10%" >
							<col width="7%" >
							<col width="8%" >
							<col width="12%" >
							<col width="15%" >
							<col width="*" >
							<col width="7%" >
							<col width="7%" >
							<col width="5%" >
							<col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">����</th>
								<th scope="col">������ȣ</th>
								<th scope="col">ȸ��</th>
								<th scope="col">ó������</th>
								<th scope="col">�����</th>
								<th scope="col">�μ���</th>
								<th scope="col">��ȭ��ȣ</th>
								<th scope="col">����</th>
								<th scope="col">������</th>
								<th scope="col">������</th>
								<th scope="col">�۾�����</th>
								<th scope="col">���CE</th>
							</tr>
						</thead>
						<tbody>
						<%
    					seq = tottal_record - ( page - 1 ) * pgsize
						do until rs.eof
							tel1 = rs("tel_ddd") + "-" + rs("tel_no1") + "-" + rs("tel_no2")
							tel2 = rs("hp_ddd") + "-" + rs("hp_no1") + "-" + rs("hp_no2")
							if rs("as_type") = "������" or rs("as_type") = "����������" then
								work_cnt = rs("ran_cnt")
							  else
							  	work_cnt = rs("dev_inst_cnt")
							end if
						%>
							<tr>
								<td class="first"><%=seq%></td>
								<td><%=rs("paper_no")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("as_type")%></td>
								<td><%=rs("acpt_user")%></td>
								<td><%=rs("dept")%></td>
								<td><%=tel1%>,<%=tel2%></td>
								<td class="left"><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=rs("request_date")%></td>
								<td><%=rs("end_date")%></td>
								<td><%=work_cnt%></td>
								<td><%=rs("mg_ce")%></td>
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
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td width="15%">
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="large_list.asp?page=<%=first_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="large_list.asp?page=<%=intstart -1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="large_list.asp?page=<%=i%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="large_list.asp?page=<%=intend+1%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[����]</a> <a href="large_list.asp?page=<%=total_page%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
                    </td>
			      </tr>
				  </table>
				<input type="hidden" name="user_id">
				<input type="hidden" name="pass">
			</form>
		</div>				
	</div>        				
	</body>
</html>

