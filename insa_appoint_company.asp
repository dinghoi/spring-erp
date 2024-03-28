<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

be_pg = "insa_appoint_company.asp"
curr_date = datevalue(mid(cstr(now()),1,10))

in_empno =""
in_name = ""
If Request.Form("in_empno")  <> "" Then 
  in_empno = Request.Form("in_empno") 
End If

win_sw = "close"
Page=Request("page")

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	field_check=Request("field_check")
	field_view=Request("field_view")
	page_cnt=Request("page_cnt")

Else
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	page_cnt=Request.form("page_cnt")
End if


Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If Request.Form("in_empno")  <> "" Then 
   Sql = "SELECT * FROM emp_master where emp_no = '"&in_empno&"'"
   Set rs_emp = DbConn.Execute(SQL)
  
   if not Rs_emp.eof then
      in_name = rs_emp("emp_name")
	  else
      response.write"<script language=javascript>"
	  response.write"alert('등록된 직원이 아닙니다....');"		
	  response.write"</script>"
	  Response.End	
   end if
   
   if isNull(rs_emp("emp_end_date")) or  rs_emp("emp_end_date") = "1900-01-01" then
      response.write"<script language=javascript>"
	  response.write"alert('발령전 회사에서 퇴직발령을 먼저 하셔야 합니다....');"		
	  response.write"</script>"
	  Response.End	
   end if
   
   rs_emp.close()
End If

sql = "select * from emp_master where emp_no = '" + in_empno + "' ORDER BY emp_no,emp_name ASC"
Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " 계열전적 인사발령 처리  "
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
			function goAction () {
			   window.close () ;
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
				if (document.frm.in_empno.value == "") {
					alert ("사번을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_appoint_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_appoint_company.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
							<strong>사번 : </strong>
								<label>
        						<input name="in_empno" type="text" id="in_empno" value="<%=in_empno%>" style="width:100px; text-align:left">
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
							<col width="5%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="6%" >
                            <col width="6%" >
							<col width="9%" >
							<col width="6%" >
							<col width="6%" >
							<col width="7%" >
                            <col width="6%" >
							<col width="*" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
						       <th class="first" scope="col">사번</th>
							   <th scope="col">성  명</th>
							   <th scope="col">직급</th>
							   <th scope="col">직책</th>
							   <th scope="col">입사일</th>
                               <th scope="col">퇴직일</th>
                               <th scope="col">소속</th>
                               <th scope="col">최초입사일</th>
							   <th scope="col">소속발령일</th>
							   <th scope="col">상주처</th>
                               <th scope="col">생년월일</th>
							   <th scope="col">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
                               <th>처리</th>
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
						emp_type = rs("emp_type")
						%>                        
							<tr>
								<td class="first"><%=rs("emp_no")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("emp_no")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("emp_name")%></a>
								</td>
                                <td><%=rs("emp_grade")%>&nbsp;</td>
                                <td><%=rs("emp_position")%>&nbsp;</td>
                                <td><%=rs("emp_in_date")%>&nbsp;</td>
                                <td><%=rs("emp_end_date")%>&nbsp;</td>
                                <td><%=rs("emp_org_name")%>&nbsp;</td>
                                <td><%=rs("emp_first_date")%>&nbsp;</td>
                                <td><%=emp_org_baldate%>&nbsp;</td>
                                <td><%=rs("emp_reside_place")%>&nbsp;</td>
                                <td><%=rs("emp_birthday")%>&nbsp;</td>
                                <td class="left"><%=rs("emp_company")%>-<%=rs("emp_bonbu")%>-<%=rs("emp_saupbu")%>-<%=rs("emp_team")%></td>
							    <td>
                                <a href="insa_appo_company_add.asp?emp_no=<%=rs("emp_no")%>&emp_name=<%=in_name%>&be_pg=<%=be_pg%>&u_type=<%="U"%>">발령</a>
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
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<% if end_view = "Y" then %>
                    <div class="btnRight">
					<a href="#" onClick="pop_Window('insa_appoint_add.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">발령등록</a>
                    <% end if %>
					<% if end_view = "Y" then %>
					<a href="payment_slip_end.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&over_cash=<%=over_cash%>&use_cash=<%=use_cash%>" class="btnType04">전표마감</a>
					<% end if %>
					<% if user_id = "jinhs" then %>
					<a href="payment_slip_end_cancle.asp?from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">마감취소</a>
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

