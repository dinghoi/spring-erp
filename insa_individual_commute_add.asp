<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
empno = request("in_empno")
emp_name = request("in_name")

commute_get_date = datevalue(mid(cstr(now()),1,10))
commute_get_time = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect


if u_type = "S" then
title_line = " 출근 시간 등록 "
commute_date_title_line = "출근일"
commute_time_title_line = "출근시간"
else
title_line = " 퇴근 시간 등록 "
commute_date_title_line = "퇴근일"
commute_time_title_line = "퇴근시간"
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>

    <link href="/include/jquery.ui.timepicker.css" type="text/css" rel="stylesheet">
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
    <script src='/java/jquery.ui.timepicker.js'></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=commute_get_date%>" );
			});	 
			</script>
			<style>
			.ui-timepicker { font-size: 12px; width: 80px; }
			</style>    
			<script type="text/javascript">
				
				var date = new Date();
				var hour = date.getHours();
				var minute = date.getMinutes();

			$(function() {
			    $('.timepicker').timepicker();
			});
      </script>
			<script type="text/javascript">
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
					return true;
			}
      </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_commute_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
				    <tbody>
				    	<colgroup>
							<col width="10%" >
							<col width="25%" >
							<col width="10%" >
							<col width="55%" >
						</colgroup>
                    <tr>
                      <th scope="col" style="background:#FFFFE6">사번</th>
                      <td scope="col" class="left" bgcolor="#FFFFE6">
					  <input name="empno" type="text" id="empno" size="14" value="<%=empno%>" readonly="true">
                      <th scope="col" style="background:#FFFFE6">성명</th>
                      <td scope="col" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>
                      <th scope="col"><%=commute_date_title_line%></th>
                      <td scope="col" class="left">
					  <input name="commute_date" type="text" value="<%=commute_get_date%>" style="width:80px;text-align:center" id="datepicker">&nbsp;
                      </td>

                      <th scope="col"><%=commute_time_title_line%></th>
                      <td scope="col" class="left">
					  <input name="commute_time" type="text" style="width:80px;text-align:center" id="timepicker" class="timepicker">&nbsp;
                      </td>
                    </tr>  
			    	</tr>
			        </tbody>
			      </table>
				</div>
                <br>
                <div align=center>
				<%	
				'if end_sw = "N" then	%>
                    <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
				'end if	%>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

