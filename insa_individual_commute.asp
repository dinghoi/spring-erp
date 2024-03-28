<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows
dim page_cnt
dim pg_cnt

be_pg = "insa_individual_commute.asp"

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

curr_dd = cstr(datepart("d",now))
curr_time = datevalue(mid(cstr(now()),1,10))

from_date = request.form("from_date")
to_date = request.form("to_date")
Page=Request("page")
page_cnt=Request.form("page_cnt")
pg_cnt=cint(Request("pg_cnt"))

if from_date = "" then
	from_date = mid(cstr(now()-curr_dd+1),1,10)
end if

if to_date = "" then
	to_date = datevalue(mid(cstr(now()),1,10))
end if

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
Set Rs_use = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select count(*) from commute where emp_no = '"&in_empno&"' and wrkt_dt between '"&from_date&"' and '"&to_date&"'"

Set RsCount = Dbconn.Execute (sql)

tottal_record = cint(RsCount(0)) 'Result.RecordCount

IF tottal_record mod pgsize = 0 THEN
	total_page = int(tottal_record / pgsize) 'Result.PageCount
  ELSE
	total_page = int((tottal_record / pgsize) + 1)
END IF

Sql = "select * from commute where emp_no = '"&in_empno&"' and wrkt_dt between '"&from_date&"' and '"&to_date&"' ORDER BY wrkt_dt ASC limit "& stpage & "," &pgsize

Rs.Open Sql, Dbconn, 1

title_line = " 출퇴근 현황"

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
				return "4 1";
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
				if (document.frm.from_date.value == "") {
					alert ("시작일자 입력하시기 바랍니다");
					return false;
				}
				
				return true;
			}
			
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_pappo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_commute.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조회기간</dt>
                        <dd>
							<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="datepicker">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
								</label>
						</dd>
					</dl>
				</fieldset>
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
							<col width="5%" >
							<col width="6%">
						</colgroup>
						<thead>
              <tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
                <th scope="col">회사</th>
                <th scope="col">소속</th>
								<th scope="col">출근일</th>
								<th scope="col">출근시간</th>
								<th scope="col">근무형태</th>
							</tr>              
						</thead>
						<tbody>
						<%
						do until rs.eof

                         emp_no = rs("emp_no")
                         if emp_no <> "" then
		                    Sql="select * from emp_master where emp_no = '"&emp_no&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                          emp_name = Rs_emp("emp_name")
												  emp_grade = Rs_emp("emp_grade")
												  emp_job = Rs_emp("emp_job")
							            emp_position = Rs_emp("emp_position")
												  emp_org_code = Rs_emp("emp_org_code")
												  emp_org_name = Rs_emp("emp_org_name")
												  emp_company = Rs_emp("emp_company")
		                   end if
	                       Rs_emp.Close()
	                	  end if	

	           			%>
							<tr>
								<td class="first"><%=rs("emp_no")%>&nbsp;</td>
                                <td><%=emp_name%>&nbsp;</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=emp_company%>&nbsp;</td>
                                <td><%=emp_org_name%>&nbsp;</td>
                                <td><%=rs("wrkt_dt")%>&nbsp;</td>
                                <td><%=rs("wrk_start_time")%>&nbsp;</td>
                                <td><%=rs("wrk_type")%>&nbsp;</td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="#" onClick="pop_Window('insa_individual_commute_add.asp?in_empno=<%=in_empno%>&in_name=<%=in_name%>&u_type=<%="S"%>','insa_individual_commute_add','scrollbars=yes,width=600,height=340')" class="btnType04">출근 등록</a>
                    <!-- a href="#" onClick="pop_Window('insa_individual_commute_add.asp?in_empno=<%=in_empno%>&in_name=<%=in_name%>&u_type=<%="E"%>','insa_individual_commute_add','scrollbars=yes,width=600,height=340')" class="btnType04">퇴근 등록</a -->
					</div> 
                    </td>
			      </tr>
				  </table>                    
				</div>
        <input type="hidden" name="emp_empno" value="<%=in_empno%>" ID="Hidden1">
        
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
                        <a href = "insa_individual_commute.asp?page=<%=first_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[처음]</a>
                  	<% if intstart > 1 then %>
                        <a href="insa_individual_commute.asp?page=<%=intstart -1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[이전]</a>
                      <% end if %>
                      <% for i = intstart to intend %>
           	<% if i = int(page) then %>
                        <b>[<%=i%>]</b>
                      <% else %>
                        <a href="insa_individual_commute.asp?page=<%=i%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[<%=i%>]</a>
                      <% end if %>
                      <% next %>
           	<% if 	intend < total_page then %>
                        <a href="insa_individual_commute.asp?page=<%=intend+1%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[다음]</a> <a href="insa_gun_mg.asp?page=<%=total_page%>&view_condi=<%=view_condi%>&from_date=<%=from_date%>&to_date=<%=to_date%>&ck_sw=<%="y"%>">[마지막]</a>
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
	</body>
</html>

