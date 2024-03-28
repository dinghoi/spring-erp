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
	view_c=Request("view_c")
	field_view=Request("field_view")
  else
	view_c=Request.form("view_c")
	field_view=Request.form("field_view")
End if

If view_c = "" Then
	view_c = "total"
End If

if view_c = "total" then
	field_view = ""
end if

pgsize = 10 ' 화면 한 페이지

If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_sum = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

' 조건별 조회.........
base_sql = "select * from trade where use_sw ='Y' and mg_group = '"&mg_group&"'"

if view_c = "total" then
	com_sql = " "
  else
	if field_view = "" then
  		com_sql = " and (trade_name = '') "
	  else
  		com_sql = " and (trade_name like '%" + field_view + "%') "
	end if
end if

order_sql = " ORDER BY trade_name ASC"

Sql = "SELECT count(*) from trade where use_sw ='Y' and mg_group = '"&mg_group&"'"&com_sql
Set RsCount = Dbconn.Execute (sql)
tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = base_sql + com_sql + order_sql + " limit "& stpage & "," &pgsize
rs.Open Sql, Dbconn, 1

title_line = "회사별 양식 관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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
				if (chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm() {
				if (document.frm.view_c.value == "") {
					alert ("조회조건을 선택하시기 바랍니다");
					return false;
				}
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/as_sub_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="company_form_mg.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
						<dt>조건검색</dt>
                        <dd>
                            <p>
                                <label>
								<strong>조회조건</strong>
                              	<input type="radio" name="view_c" value="total" <% if view_c = "total" then %>checked<% end if %> style="width:25px">전체

                                <input type="radio" name="view_c" value="com" <% if view_c = "com" then %>checked<% end if %> style="width:25px">회사명 검색
								</label>
								<input name="field_view" type="text" value="<%=field_view%>" style="width:150px; text-align:left" >
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
							<col width="8%" >
							<col width="10%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">회사</th>
								<th colspan="2" scope="col">양식 1</th>
								<th colspan="2" scope="col">양식 2</th>
								<th colspan="2" scope="col">양식 3</th>
								<th colspan="2" scope="col">양식 4</th>
								<th colspan="2" scope="col">양식 5</th>
							</tr>
						</thead>
						<tbody>
						<%
						path = "/forms"
 						do until rs.eof
							sql = "select * from company_form where company = '"&rs("trade_name")&"'"
							set rs_etc = dbconn.execute(sql)
							if rs_etc.eof or rs_etc.bof then
								form1 = ""
								up_date1 = ""
								form2 = ""
								up_date2 = ""
								form3 = ""
								up_date3 = ""
								form4 = ""
								up_date4 = ""
								form5 = ""
								up_date5 = ""
							  else
								form1 = rs_etc("form1")
								up_date1 = rs_etc("up_date1")
								form2 = rs_etc("form2")
								up_date2 = rs_etc("up_date2")
								form3 = rs_etc("form3")
								up_date3 = rs_etc("up_date3")
								form4 = rs_etc("form4")
								up_date4 = rs_etc("up_date4")
								form5 = rs_etc("form5")
								up_date5 = rs_etc("up_date5")
							end if
						%>
							<tr>
								<td rowspan="2" class="first"><%=rs("trade_name")%></td>
								<td colspan="2">&nbsp;<a href="download.asp?path=<%=path%>&att_file=<%=form1%>"><%=form1%></a></td>
								<td colspan="2">&nbsp;<a href="download.asp?path=<%=path%>&att_file=<%=form2%>"><%=form2%></a></td>
								<td colspan="2">&nbsp;<a href="download.asp?path=<%=path%>&att_file=<%=form3%>"><%=form3%></a></td>
								<td colspan="2">&nbsp;<a href="download.asp?path=<%=path%>&att_file=<%=form4%>"><%=form4%></a></td>
								<td colspan="2">&nbsp;<a href="download.asp?path=<%=path%>&att_file=<%=form5%>"><%=form5%></a></td>
							</tr>
							<tr>
								<td style=" border-left:1px solid #e3e3e3;">&nbsp;<%=up_date1%></td>
								<td>
							<% if c_grade < "2"  or user_id = "101955" then ' 윤종윤 01.01.10 요구	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload.asp?company=<%=rs("trade_name")%>&seq=<%=1%>','form_upload_pop','scrollbars=yes,width=700,height=200')"><strong>UP</strong></a>
							<%     if form1 > "1" then	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload_del.asp?company=<%=rs("trade_name")%>&seq=<%=1%>','form_upload_del_pop','scrollbars=yes,width=10,height=10')"><strong>DEL</strong></a>
							<%	   end if	%>
							<%   else	%>
                            	&nbsp;
							<%  end if	%>
                                </td>
								<td>&nbsp;<%=up_date2%></td>
								<td>
							<% if c_grade < "2"  or user_id = "101955" then ' 윤종윤 01.01.10 요구	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload.asp?company=<%=rs("trade_name")%>&seq=<%=2%>','form_upload_pop','scrollbars=yes,width=700,height=200')"><strong>UP</strong></a>
							<%     if form2 > "1" then	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload_del.asp?company=<%=rs("trade_name")%>&seq=<%=2%>','form_upload_del_pop','scrollbars=yes,width=10,height=10')"><strong>DEL</strong></a>
							<%	   end if	%>
							<%   else	%>
                            	&nbsp;
							<%  end if	%>
                                </td>
								<td>&nbsp;<%=up_date3%></td>
								<td>
							<% if c_grade < "2"  or user_id = "101955" then ' 윤종윤 01.01.10 요구	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload.asp?company=<%=rs("trade_name")%>&seq=<%=3%>','form_upload_pop','scrollbars=yes,width=700,height=200')"><strong>UP</strong></a>
							<%     if form3 > "1" then	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload_del.asp?company=<%=rs("trade_name")%>&seq=<%=3%>','form_upload_del_pop','scrollbars=yes,width=10,height=10')"><strong>DEL</strong></a>
							<%	   end if	%>
							<%   else	%>
                            	&nbsp;
							<%  end if	%>
                                </td>
								<td>&nbsp;<%=up_date4%></td>
								<td>
							<% if c_grade < "2"  or user_id = "101955" then ' 윤종윤 01.01.10 요구	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload.asp?company=<%=rs("trade_name")%>&seq=<%=4%>','form_upload_pop','scrollbars=yes,width=700,height=200')"><strong>UP</strong></a>
							<%     if form4 > "1" then	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload_del.asp?company=<%=rs("trade_name")%>&seq=<%=4%>','form_upload_del_pop','scrollbars=yes,width=10,height=10')"><strong>DEL</strong></a>
							<%	   end if	%>
							<%   else	%>
                            	&nbsp;
							<%  end if	%>
                                </td>
								<td>&nbsp;<%=up_date5%></td>
								<td>
							<% if c_grade < "2"  or user_id = "101955" then ' 윤종윤 01.01.10 요구	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload.asp?company=<%=rs("trade_name")%>&seq=<%=5%>','form_upload_pop','scrollbars=yes,width=700,height=200')"><strong>UP</strong></a>
							<%     if form5 > "1" then	%>
                                <a href="#" class="btnType03" onClick="pop_Window('form_upload_del.asp?company=<%=rs("trade_name")%>&seq=<%=5%>','form_upload_del_pop','scrollbars=yes,width=10,height=10')"><strong>DEL</strong></a>
							<%	   end if	%>
							<%   else	%>
                            	&nbsp;
							<%  end if	%>
                                </td>
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
                        <a href="company_form_mg.asp?page=<%=first_page%>&ck_sw=<%="y"%>&view_c=<%=view_c%>&field_view=<%=field_view%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="company_form_mg.asp?page=<%=intstart -1%>&ck_sw=<%="y"%>&view_c=<%=view_c%>&field_view=<%=field_view%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	    <% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="company_form_mg.asp?page=<%=i%>&ck_sw=<%="y"%>&view_c=<%=view_c%>&field_view=<%=field_view%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	    <% if 	intend < total_page then %>
                        <a href="company_form_mg.asp?page=<%=intend+1%>&ck_sw=<%="y"%>&view_c=<%=view_c%>&field_view=<%=field_view%>">[다음]</a> <a href="company_form_mg.asp?page=<%=total_page%>&ck_sw=<%="y"%>&view_c=<%=view_c%>&field_view=<%=field_view%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
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

