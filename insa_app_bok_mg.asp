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
be_pg = "insa_app_bok_mg.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_emp = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")

dbconn.open DbConnect

view_sort = request("view_sort")

if view_sort = "" then
	view_sort = "DESC"
end if


order_Sql = " ORDER BY app_date,app_empno,app_seq " + view_sort
where_sql = " WHERE app_id = '휴직발령' and app_bokjik_id = 'N'"
'where_sql = ""

Sql = "SELECT count(*) FROM emp_appoint " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_appoint " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " 휴직발령 현황 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "2 1";
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
					alert ("소속을 선택하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_app_bok_mg.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
							<col width="7%" >
                            <col width="10%" >
                            <col width="10%" >
							<col width="7%" >
							<col width="7%" >
							<col width="14%" >
                            <col width="18%" >
                            <col width="6%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
                                <th scope="col">발령일</th>
								<th scope="col">휴직유형</th>
								<th scope="col">휴직기간</th>
								<th scope="col">휴직사유</th>
								<th scope="col">발령</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof
						
		                  app_empno = rs("app_empno")
		                  app_emp_name = rs("app_emp_name")
		
                         if app_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&app_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
							  emp_grade = Rs_emp("emp_job")
							  emp_position = Rs_emp("emp_position")
		                      emp_org_code = Rs_emp("emp_org_code")
							  emp_org_name = Rs_emp("emp_org_name")
							  emp_company = Rs_emp("emp_company")
		                   end if
	                       Rs_emp.Close()
	                	end if		
						%>
							<tr>
								<td class="first"><%=rs("app_empno")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("app_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("app_emp_name")%></a>
								</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rs("app_to_company")%>&nbsp;</td>
                                <td><%=rs("app_to_org")%>&nbsp;</td>
                                <td><%=rs("app_date")%>&nbsp;</td>
                                <td><%=rs("app_id_type")%>&nbsp;</td>
                                <td><%=rs("app_start_date")%>&nbsp;∼&nbsp;<%=rs("app_finish_date")%></td>
								<td><%=rs("app_comment")%>&nbsp;</td>
                                <td><a href="#" onClick="pop_Window('insa_app_bokadd.asp?app_empno=<%=rs("app_empno")%>&emp_name=<%=rs("app_emp_name")%>&app_seq=<%=rs("app_seq")%>&app_id=<%=rs("app_id")%>&app_date=<%=rs("app_date")%>&u_type=<%=""%>','insa_app_bokadd_pop','scrollbars=yes,width=750,height=350')">복직</a>&nbsp;</td>
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
                    <a href="insa_excel_apphujik.asp" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_app_bok_mg.asp?page=<%=first_page%>&view_sort=<%=view_sort%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_app_bok_mg.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_app_bok_mg.asp?page=<%=i%>&view_sort=<%=view_sort%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_app_bok_mg.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>">[다음]</a> <a href="insa_app_bok_mg.asp?page=<%=total_page%>&view_sort=<%=view_sort%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
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
        <input type="hidden" name="app_to_company" value="<%=emp_bonbu%>" ID="Hidden1">
        <input type="hidden" name="app_to_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
        <input type="hidden" name="app_to_team" value="<%=emp_team%>" ID="Hidden1">
        <input type="hidden" name="app_to_org" value="<%=emp_org_code%>" ID="Hidden1">
        <input type="hidden" name="app_to_org_name" value="<%=emp_org_name%>" ID="Hidden1">        
	</body>
</html>

