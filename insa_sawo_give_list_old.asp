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
be_pg = "insa_sawo_give_list.asp"
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

order_Sql = " ORDER BY give_company,give_date,give_empno DESC"
'where_sql = " WHERE (give_ask_process = '2') and (give_company = '"+view_condi+"') and (give_date > '"+from_date+"') and (give_date < '"+to_date+"')"
where_sql = " WHERE give_ask_process = '2'"


Sql = "SELECT count(*) FROM emp_sawo_give " + where_sql
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_sawo_give " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1

title_line = " 경조회 경조금 지급내역 "
ask_process = "2"

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
				return "8 1";
			}
			function goAction () {
			   window.close () ;
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
			<!--#include virtual = "/include/insa_sawo_menu.asp" -->    
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_give_list.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="9%" >
                            <col width="9%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="9%" >
							<col width="12%" >
                            <col width="12%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">현직급</th>
								<th scope="col">현직책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
								<th scope="col">지급일</th>
								<th scope="col">지급구분</th>
								<th scope="col">지급유형</th>
                                <th scope="col">발생일</th>
                                <th scope="col">지급금액</th>
                                <th scope="col">경조장소</th>
                                <th scope="col">경조내용</th>
                                <th scope="col">비  고</th>
							</tr>
						</thead>
					<tbody>
						<%
						do until rs.eof
						
		                  give_empno = rs("give_empno")
		                  give_emp_name = rs("give_emp_name")
		
                         if give_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&give_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
		                      emp_position = Rs_emp("emp_position")
		                   end if
	                       Rs_emp.Close()
	                	end if		
						%>
							<tr>
								<td class="first"><%=rs("give_empno")%></td>
                                <td><%=rs("give_emp_name")%></td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("give_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("give_emp_name")%></a>
                                </td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rs("give_company")%>&nbsp;</td>
                                <td><%=rs("give_org_name")%>&nbsp;</td>
                                <td><%=rs("give_date")%>&nbsp;</td>
                                <td><%=rs("give_id")%>&nbsp;</td>
                                <td><%=rs("give_type")%>&nbsp;</td>
                                <td><%=rs("give_sawo_date")%>&nbsp;</td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("give_pay")),0)%>&nbsp;</td>
                                <td><%=rs("give_sawo_place")%>&nbsp;</td>
                                <td class="left"><%=rs("give_sawo_comm")%>&nbsp;</td>
                                <td class="left"><%=rs("give_comment")%>&nbsp;</td>
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
                    <a href="insa_excel_sawo_give.asp?view_condi=<%=view_condi%>&ask_process=<%=ask_process%>&from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
					</div>                  
                  	</td>
				    <td>
                    <div id="paging">
                        <a href="insa_sawo_give_list.asp?page=<%=first_page%>&view_sort=<%=view_sort%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_sawo_give_list.asp?page=<%=intstart -1%>&view_sort=<%=view_sort%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_sawo_give_list.asp?page=<%=i%>&view_sort=<%=view_sort%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_sawo_give_list.asp?page=<%=intend+1%>&view_sort=<%=view_sort%>">[다음]</a> <a href="insa_sawo_give_list.asp?page=<%=total_page%>&view_sort=<%=view_sort%>">[마지막]</a>
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
	</body>
</html>

