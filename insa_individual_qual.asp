<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

win_sw = "close"

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
   rs_emp.close()
End If

sql = "select * from emp_qual where qual_empno = '" + in_empno + "' ORDER BY qual_empno,qual_seq ASC"
Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " 자격 사항 "
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>개인업무-인사</title>
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
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_psub_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_qual.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
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
							<col width="15%" >
							<col width="8%" >
							<col width="9%" >
							<col width="15%" >
							<col width="*" >
							<col width="15%" >
                            <col width="5%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                            <th>자격증 종목</th>
                            <th>등급</th>
                            <th>합격년월일</th>
                            <th>발급 기관명</th>
                            <th>자격 등록번호</th>
                            <th>경력수첩No.</th>
                            <th>순번</th>
                            <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						do until rs.eof
						%>
							<tr>
                              <td ><%=rs("qual_type")%>&nbsp;</td>
                              <td ><%=rs("qual_grade")%>&nbsp;</td>
                              <td ><%=rs("qual_pass_date")%>&nbsp;</td>
                              <td ><%=rs("qual_org")%>&nbsp;</td>
                              <td ><%=rs("qual_no")%>&nbsp;</td>
                              <td ><%=rs("qual_passport")%>&nbsp;</td>
                              <td ><%=rs("qual_seq")%>&nbsp;</td>
							  <td><a href="#" onClick="pop_Window('insa_individual_qual_add.asp?qual_empno=<%=rs("qual_empno")%>&qual_seq=<%=rs("qual_seq")%>&emp_name=<%=in_name%>&u_type=<%="U"%>','insa_qual_add_pop','scrollbars=yes,width=750,height=300')">수정</a></td>
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
					<div class="btnRight">
					<a href="#" onClick="pop_Window('insa_individual_qual_add.asp?qual_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_qual_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">자격 등록</a>
					</div>                  
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="qual_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

