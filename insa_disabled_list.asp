<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt
Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_disabled_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

if page_cnt > 0 then 
	pg_cnt = page_cnt
end if
if pg_cnt > 0 then
	page_cnt = pg_cnt
end if

if page_cnt < 10 or page_cnt > 20 then
	page_cnt = 10
end if

pgsize = page_cnt ' ȭ�� �� ������ 

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

order_Sql = " ORDER BY emp_company, emp_in_date,emp_no DESC"
where_sql = " WHERE emp_disabled_yn = 'Y'"

Sql = "SELECT count(*) FROM emp_master " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_master " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " ����� ��Ȳ "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>�λ���� �ý���</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "5 1";
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
			<!--#include virtual = "/include/insa_report_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_disabled_list.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
                            <col width="9%" >
							<col width="6%" >
							<col width="4%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">����<br>�Ի���</th>
                                <th scope="col">�Ի���</th>
                                <th scope="col">������</th>
                                <th scope="col">ȸ��</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">�������</th>
								<th scope="col">���<br>���</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof
						
						if rs("emp_end_date") = "1900-01-01" then
						   emp_end_date = ""
						   else 
						   emp_end_date = rs("emp_end_date")
						end if
						if rs("emp_org_baldate") = "1900-01-01" then
						   emp_org_baldate = ""
						   else 
						   emp_org_baldate = rs("emp_org_baldate")
						end if
						if rs("emp_grade_date") = "1900-01-01" then
						   emp_grade_date = ""
						   else 
						   emp_grade_date = rs("emp_grade_date")
						end if
						%>
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=emp_end_date%>&nbsp;</td>
                                <td><%=rs("emp_company")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_disabled")%>&nbsp;</td>
                                <td><%=rs("emp_disab_grade")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
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
                    <a href="insa_excel_disabled.asp" class="btnType04">�����ٿ�ε�</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_disabled_list.asp?page=<%=first_page%>&view_sort=<%=view_sort%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_disabled_list.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_disabled_list.asp?page=<%=i%>&view_sort=<%=view_sort%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_disabled_list.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>">[����]</a> <a href="insa_disabled_list.asp?page=<%=total_page%>&view_sort=<%=view_sort%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
                      <% end if %>
                    </div>
                    </td>
				    <td width="15%">
					<div class="btnCenter">
					<% if view_sort = "DESC" then	%>
                          <a href="insa_disabled_list.asp?view_sort=<%="ASC"%>" >&nbsp;</a>
                    <% else %>
                          <a href="insa_disabled_list.asp?view_sort=<%="DESC"%>" >&nbsp;</a>
                    <% end if %>
					</div>                  
                    </td>
			      </tr>
				  </table>
			</form>
		</div>				
	</div>        				
		<input type="hidden" name="user_id">
		<input type="hidden" name="pass">
	</body>
</html>

