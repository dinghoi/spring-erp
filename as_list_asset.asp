<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
Dim from_date
Dim to_date
Dim as_process
Dim field_check
Dim field_view
Dim win_sw
dim company_tab(160)

win_sw = "close"
be_pg = "as_list_asset.asp"

ck_sw=Request("ck_sw")
Page=Request("page")

If ck_sw = "y" Then
	from_date=Request("from_date")
	to_date=Request("to_date")
	date_sw=Request("date_sw")
	process_sw=Request("process_sw")
	field_check=Request("field_check")
	field_view=Request("field_view")
 else
	from_date=Request.form("from_date")
	to_date=Request.form("to_date")
	date_sw=Request.form("date_sw")
	process_sw=Request.form("process_sw")
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
End if

If to_date = "" or from_date = "" Then
	to_date = mid(cstr(now()),1,10)
	from_date = mid(cstr(now()-7),1,10)
	field_check = "total"
	date_sw = "acpt"
	process_sw = "N"
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

base_sql = "select acpt_no,acpt_man,as_type,acpt_date,as_process,acpt_user,as_memo,company,dept,tel_ddd,tel_no1,tel_no2,sido,gugun,request_date,visit_date,mg_ce,asets_no from as_acpt "

if date_sw = "acpt" then
	date_sql = "where (CAST(acpt_date as date) >= '" + from_date  + "' and CAST(acpt_date as date) <= '" + to_date  + "') and (mg_group ='" + mg_group + "') and company = '" + user_name + "'"
  else
	date_sql = "where (visit_date >= '" + from_date  + "' and visit_date <= '" + to_date  + "') and (mg_group ='" + mg_group + "') and company = '" + user_name + "'"
end if

if process_sw = "Y" then
	process_sql = " and ( as_process = '�Ϸ�' or as_process = '��ü' or as_process = '���' ) "
  else
	process_sql = " and ( as_process = '����' or as_process = '����' or as_process = '�԰�' or as_process = '��ü�԰�' ) "
end if

if field_check <> "total" then
	if field_check = "asets_no" then
		field_sql = " and ( " + field_check + " = '" + field_view + "' ) "
	  else
		field_sql = " and ( " + field_check + " like '%" + field_view + "%' ) "
	end if
  else
  	field_sql = " "
end if
order_sql = " ORDER BY acpt_date DESC"

'sql = base_sql + date_sql + process_sql + field_sql + order_sql

com_sql = " "

Sql = "SELECT count(*) FROM as_acpt " + date_sql + com_sql + process_sql + field_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + date_sql + com_sql + process_sql + field_sql + order_sql + " limit "& stpage & "," &pgsize
Rs.Open Sql, Dbconn, 1

'Response.write sql

title_line = "A/S �Ѱ� ��Ȳ"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "0 1";
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
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/asset_header.asp" -->
			<!--#include virtual = "/include/asset_as_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="as_list_asset.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>���ǰ˻�</dt>
                        <dd>
                            <p>
                                <input name="process_sw" type="radio" value="N"  <% if process_sw = "N" then %>checked<% end if %> style="width:25px">��ó��
                                <input name="process_sw" type="radio" value="Y"  <% if process_sw = "Y" then %>checked<% end if %> style="width:25px">ó���Ϸ�

                              	<input type="radio" name="date_sw" value="acpt" <% if date_sw = "acpt" then %>checked<% end if %> style="width:25px">������
                              	<input type="radio" name="date_sw" value="visit" <% if date_sw = "visit" then %>checked<% end if %> style="width:25px">�Ϸ���
								<label>
								&nbsp;&nbsp;<strong>������ : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>������ : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <label>
								<strong>�ʵ�����</strong>
                                <select name="field_check" id="field_check" style="width:70px">
                              		<option value="total" <% if field_check = "total" then %>selected<% end if %>>��ü</option>
                                    <option value="mg_ce" <% if field_check = "mg_ce" then %>selected<% end if %>>���CE</option>
                                    <option value="acpt_man" <% if field_check = "acpt_man" then %>selected<% end if %>>������</option>
                                    <option value="sido" <% if field_check = "sido" then %>selected<% end if %>>�õ�</option>
                                    <option value="gugun" <% if field_check = "gugun" then %>selected<% end if %>>����</option>
                                    <option value="acpt_user" <% if field_check = "acpt_user" then %>selected<% end if %>>�����</option>
                                    <option value="dept" <% if field_check = "dept" then %>selected<% end if %>>������</option>
                                    <option value="asets_no" <% if field_check = "asets_no" then %>selected<% end if %>>�ڻ��ȣ</option>
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
							<col width="7%" >
							<col width="8%" >
							<col width="4%" >
							<col width="6%" >
							<col width="10%" >
							<col width="10%" >
							<col width="8%" >
							<col width="11%" >
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="*" >
							<col width="3%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">ó������</th>
								<th scope="col">��������</th>
								<th scope="col">����</th>
								<th scope="col">�����</th>
								<th scope="col">ȸ��</th>
								<th scope="col">������</th>
								<th scope="col">��ȭ��ȣ</th>
								<th scope="col">����</th>
								<th scope="col">��û����</th>
								<th scope="col">ó������</th>
								<th scope="col">���CE</th>
								<th scope="col">��ֳ���</th>
								<th scope="col">��ȸ</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

							dim len_date, hangle, bit01, bit02, bit03
							acpt_date = rs("acpt_date")
							len_date = len(acpt_date)
							bit01 = left(acpt_date, 10)
						' 	bit01 = Replace(bit01,"-",".")
							bit03 = left(right(acpt_date, 5), 2)
							hangle = mid(acpt_date, 12, 2)
							if len_date = 22 then
								bit02 = mid(acpt_date, 15, 2)
							  else
								bit02 = "0"&mid(acpt_date, 15, 1)
							end If

							if hangle = "����" and bit02 <> 12 then
								bit02 = bit02 + 12
							end if

							date_to_date = bit01 & " " &bit02 & ":" & bit03
							acpt_date = mid(date_to_date,3)
'							acpt_date = replace(acpt_date,"-","/")

							as_memo = replace(rs("as_memo"),chr(34),chr(39))
							view_memo = as_memo
							if len(as_memo) > 18 then
								view_memo = mid(as_memo,1,18) + ".."
							end if
							if rs("as_process") = "����" or rs("as_process") = "����" or rs("as_process") = "�԰�" then
								visit_date = "."
								if rs("as_process") = "�԰�" then
									sql_into = "select in_process, into_date from as_into where acpt_no = "&rs("acpt_no")&" and in_process = '�����Ϸ�'"
									Set rs_into=DbConn.Execute(sql_into)
									if rs_into.eof or rs_into.bof then
										visit_date = "."
									  else
										visit_date = rs_into("into_date")
									end if
									rs_into.close()
								end if
							  else
								visit_date = mid(rs("visit_date"),3)
							end if
						%>
							<tr>
								<td class="first"><%=rs("as_type")%></td>
								<td><%=acpt_date%></td>
								<td><%=rs("as_process")%></td>
								<td><%=rs("acpt_user")%></td>
								<td><%=rs("company")%></td>
								<td><%=rs("dept")%></td>
								<td><%=rs("tel_ddd")%>)<%=rs("tel_no1")%>-<%=rs("tel_no2")%></td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%></td>
								<td><%=mid(rs("request_date"),3)%></td>
								<td><%=visit_date%></td>
								<td><%=rs("mg_ce")%></td>
							  	<td class="left"><p style="cursor:pointer"><span title="<%=as_memo%>"><%=view_memo%></span></p></td>
								<td><a href="#" onClick="pop_Window('as_view.asp?acpt_no=<%=rs("acpt_no")%>','asview_pop','scrollbars=yes,width=800,height=700')">��ȸ</a></td>
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
					<div class="btnCenter">
                    <a href="excel_down_asset.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>" class="btnType04">�����ٿ�ε�</a>
					</div>
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="as_list_asset.asp?page=<%=first_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="as_list_asset.asp?page=<%=intstart -1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="as_list_asset.asp?page=<%=i%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="as_list_asset.asp?page=<%=intend+1%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[����]</a> <a href="as_list_asset.asp?page=<%=total_page%>&from_date=<%=from_date%>&to_date=<%=to_date%>&date_sw=<%=date_sw%>&process_sw=<%=process_sw%>&field_check=<%=field_check%>&field_view=<%=field_view%>&ck_sw=<%="y"%>&company=<%=company%>">[������]</a>
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

