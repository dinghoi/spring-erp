<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

'If Request.Form("in_empno")  <> "" Then 
'  in_empno = Request.Form("in_empno") 
'End If

win_sw = "close"

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If Request.Form("in_empno")  <> "" Then 
   Sql = "SELECT * FROM emp_master where emp_no = '"&in_empno&"'"
   Set rs_emp = DbConn.Execute(SQL)
   in_name = rs_emp("emp_name")
   rs_emp.close()
End If

sql = "select * from emp_family where family_empno = '" + in_empno + "' ORDER BY family_empno,family_seq ASC"
Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " 가족 사항 "
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
				<form action="insa_individual_family.asp" method="post" name="frm">
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
							<col width="9%" >
							<col width="1%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
							<col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="9%" >
                            <col width="4%" >
                            <col width="5%" >
						</colgroup>
						<thead>
                            <tr>
                                <th colspan="2">관계</th>
                                <th>성명</th>
                                <th>생년월일</th>
                                <th colspan="2">직업</th>
                                <th colspan="2">전화번호</th>
                                <th colspan="2">주민번호</th>
                                <th>동거여부</th>
                                <th>No.</th>
                                <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						do until rs.eof
						      if rs("family_person2") = "" or isnull(rs("family_person2")) then 
							          family_person2 = rs("family_person2")
								 else
									  family_person2 = "*******"
							  end if
						%>
							<tr>
                              <td colspan="2" ><%=rs("family_rel")%>&nbsp;</td>
                              <td ><%=rs("family_name")%>&nbsp;</td>
                              <td class="left"><%=rs("family_birthday")%>&nbsp;(<%=rs("family_birthday_id")%>)&nbsp;</td>
                              <td colspan="2" class="left"><%=rs("family_job")%>&nbsp;</td>
                              <td colspan="2" ><%=rs("family_tel_ddd")%>-<%=rs("family_tel_no1")%>-<%=rs("family_tel_no2")%>&nbsp;</td>
                              <td colspan="2" ><%=rs("family_person1")%>-<%=rs("family_person2")%>&nbsp;</td>
                              <td ><%=rs("family_live")%>&nbsp;</td>
                              <td ><%=rs("family_seq")%></td>
							  <td><a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=rs("family_empno")%>&family_seq=<%=rs("family_seq")%>&emp_name=<%=in_name%>&u_type=<%="U"%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')">수정</a></td>
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
					<a href="#" onClick="pop_Window('insa_family_add.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_family_add_pop','scrollbars=yes,width=750,height=400')" class="btnType04">가족등록</a>
                    <% if in_empno = "900002" then %>
                    <a href="#" onClick="pop_Window('insa_aiax_test.asp?family_empno=<%=in_empno%>&emp_name=<%=in_name%>','insa_aiax_test_pop','scrollbars=yes,width=750,height=600')" class="btnType04">aiax test</a>
                    <% end if %>
					</div>                  
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="family_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

