<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")
user_name = request.cookies("nkpmg_user")("coo_user_name")
emp_no = request.cookies("nkpmg_user")("coo_user_id")
emp_company = request.cookies("nkpmg_user")("coo_emp_company")
bonbu = request.cookies("nkpmg_user")("coo_bonbu")
saupbu = request.cookies("nkpmg_user")("coo_saupbu")
team = request.cookies("nkpmg_user")("coo_team")
org_name = request.cookies("nkpmg_user")("coo_org_name")
position = request.cookies("nkpmg_user")("coo_position")

'response.write(position)

Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))
be_pg = "insa_manager_emp_list.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

view_condi = emp_company

pgsize = 10 ' ȭ�� �� ������ 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_hol = Server.CreateObject("ADODB.Recordset")
Set rs_org = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

order_Sql = " ORDER BY emp_company,emp_bonbu,emp_saupbu,emp_team,emp_org_code,emp_no,emp_in_date ASC"
'where_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_company = '"&view_condi&"') and (emp_bonbu = '"&bonbu&"') and (emp_no < '900000')"

if position = "�Ѱ���ǥ" or position = "������" then
        where_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_bonbu = '"&bonbu&"') and (emp_no < '900000')"
   else
        where_sql = " WHERE (isNull(emp_end_date) or emp_end_date = '1900-01-01') and (emp_saupbu = '"&saupbu&"') and (emp_no < '900000')"
end if

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

title_line = ""+ bonbu +" - ���� ��Ȳ "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>���ξ���-�λ�</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
			}
		</script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.view_condi.value == "") {
					alert ("�ʵ������� �����Ͻñ� �ٶ��ϴ�");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_plist_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_manager_emp_list.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">���</th>
								<th scope="col">��  ��</th>
								<th scope="col">����</th>
								<th scope="col">����</th>
								<th scope="col">��å</th>
								<th scope="col">�Ի���</th>
                                <th scope="col">�Ҽ�</th>
                                <th scope="col">�����Ի���</th>
								<th scope="col">�Ҽӹ߷���</th>
								<th scope="col">������</th>
                                <th scope="col">�������</th>
								<th scope="col">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof
						
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
                                <a href="#" onClick="pop_Window('insa_individual_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1300,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_job")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=emp_grade_date%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
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
				    <td>
                    <div id="paging">
                        <a href="insa_manager_emp_list.asp?page=<%=first_page%>&bonbu=<%=bonbu%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[ó��]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_manager_emp_list.asp?page=<%=intstart -1%>&bonbu=<%=bonbu%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[����]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_manager_emp_list.asp?page=<%=i%>&bonbu=<%=bonbu%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_manager_emp_list.asp?page=<%=intend+1%>&bonbu=<%=bonbu%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[����]</a> <a href="insa_manager_emp_list.asp?page=<%=total_page%>&bonbu=<%=bonbu%>&view_condi=<%=view_condi%>&ck_sw=<%="y"%>">[������]</a>
                        <%	else %>
                        [����]&nbsp;[������]
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

