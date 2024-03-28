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

curr_date = datevalue(mid(cstr(now()),1,10))

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

be_pg = "insa_sawo_ask_report.asp"

cfm_use =""
cfm_use_dept =""
cfm_comment =""

win_sw = "close"

if page_cnt > 0 then 
	pg_cnt = page_cnt
end if
if pg_cnt > 0 then
	page_cnt = pg_cnt
end if

if page_cnt < 10 or page_cnt > 20 then
	page_cnt = 10
end if

pgsize = page_cnt ' 화면 한 페이지 

If Page = "" Then
	Page = 1
	start_page = 1
End If
stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

order_Sql = " ORDER BY ask_date,ask_seq DESC"
where_sql = " WHERE ask_empno = '"&in_empno&"'"

Sql = "SELECT count(*) FROM emp_sawo_ask where ask_empno = '"&in_empno&"'"
Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

sql = "select * from emp_sawo_ask " + where_sql + order_sql + " limit "& stpage & "," &pgsize 
Rs.Open Sql, Dbconn, 1


title_line = " 경조금 신청 현황 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
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
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
				if (document.frm.in_empno.value == "") {
					alert ("사번을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}
            
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psawo_menu.asp" -->
            <div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_ask_report.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=in_empno%>" readonly="true" style="width:100px; text-align:left">
								</label>
                            <strong>성명 : </strong>
                                <label>
                               	<input name="in_name" type="text" id="in_name" value="<%=in_name%>" readonly="true" style="width:150px; text-align:left">
								</label>
                                
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="10%" >
							<col width="10%" >
                            <col width="14%" >
							<col width="*" >
							<col width="10%" >
						</colgroup>
						<thead>
						  <tr>
							<th class="first" scope="col">경조일시</th>
							<th scope="col">경조구분</th>
                            <th scope="col">경조유형</th>
							<th scope="col">경조장소</th>
                            <th scope="col">기타내용</th>
                            <th scope="col">첨부</th>
						  </tr>
						</thead>
						<tbody>
						<%
						do until rs.eof

		                  ask_empno = rs("ask_empno")
		
                         if ask_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&ask_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_job = Rs_emp("emp_job")
		                      emp_position = Rs_emp("emp_position")
		                   end if
	                       Rs_emp.Close()
	                	 end if		


	           			%>
							<tr>
								<td class="first"><%=rs("ask_date")%></td>
                                <td><%=rs("ask_id")%>&nbsp;</td>
                                <td><%=rs("ask_type")%>&nbsp;</td>
								<td><%=rs("ask_sawo_place")%>&nbsp;</td>
                                <td><%=rs("ask_sawo_comm")%>&nbsp;</td>
                                <td>
								<% 
                                If rs("ask_att_file") <> "" Then 
                                    path = "/emp_sawo" 
                                %>
                                  <a href="att_file_download.asp?path=<%=path%>&att_file=<%=rs("ask_att_file")%>"><img src="image/att_file.gif" border="0"></a>
                                <% Else %>
				                    &nbsp;
                                <% 
								End If %>
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
                    <td>
                    <div id="paging">
                        <a href="insa_sawo_ask_report.asp?page=<%=first_page%>&in_empno=<%=in_empno%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_sawo_ask_report.asp?page=<%=intstart -1%>&in_empno=<%=in_empno%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
                  	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_sawo_ask_report.asp?page=<%=i%>&in_empno=<%=in_empno%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
                  	<% if 	intend < total_page then %>
                        <a href="insa_sawo_ask_report.asp?page=<%=intend+1%>&in_empno=<%=in_empno%>">[다음]</a> <a href="insa_sawo_ask_report.asp?page=<%=total_page%>&in_empno=<%=in_empno%>">[마지막]</a>
                        <%	else %>
                        [다음]&nbsp;[마지막]
                      <% end if %>
                    </div>
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="emp_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

