<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
app_empno = request("app_empno")
app_seq = request("app_seq")
app_id = request("app_id")
app_date = request("app_date")
emp_name = request("emp_name")

'response.write (app_empno)
'response.write (app_seq)
'response.write (app_id)
'response.write (emp_name)

app_emp_name = ""
app_id_type = ""
app_to_company = ""
app_to_orgcode = ""
app_to_org = ""
app_to_grade = ""
app_to_job = ""
app_to_position = ""
app_to_enddate = ""
app_be_company = ""
app_be_orgcode = ""
app_be_org = ""
app_be_grade = ""
app_be_job = ""
app_be_position = ""
app_be_enddate = ""
app_start_date = ""
app_finish_date = ""
app_reward = ""
app_comment = ""
app_bokjik_id = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_app = Server.CreateObject("ADODB.Recordset")
Set Rs_max = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 복직발령 등록 "

Sql="select * from emp_appoint where app_empno = '"&app_empno&"' and app_seq = '"&app_seq&"' and app_id = '"&app_id&"' and app_date = '"&app_date&"'"
Set rs_app=DbConn.Execute(Sql)

apphu_seq = rs_app("app_seq")
apphu_id_type = rs_app("app_id_type")
apphu_date = rs_app("app_date")
apphu_start_date = rs_app("app_start_date")
apphu_finish_date = rs_app("app_finish_date")
apphu_comment = rs_app("app_comment")

rs_app.close()

    app_bok_id = "복직발령"
    app_bok_date = ""
    sql="select max(app_seq) as max_seq from emp_appoint where app_empno = '"&app_empno&"' and app_id = '"&app_bok_id&"'"
	set rs_max=dbconn.execute(sql)
	
	if	isnull(rs_max("max_seq"))  then
		code_last = "001"
	  else
		max_seq = "00" + cstr((int(rs_max("max_seq")) + 1))
		code_last = right(max_seq,3)
	end if
    rs_max.close()
	
app_bok_seq = code_last

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=app_bok_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.app_bok_date.value =="") {
					alert('복직발령일을 입력하세요');
					frm.app_bok_date.focus();
					return false;}
				if(document.frm.apphu_finish_date.value > document.frm.app_bok_date.value) {
						alert('복직발령일이 휴직기간보다 빠름니다');
						frm.app_bok_date.focus();
						return false;}
				if(document.frm.apphu_start_date.value > document.frm.app_bok_date.value) {
						alert('복직발령일이 휴직기간보다 빠름니다');
						frm.app_bok_date.focus();
						return false;}
				
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_app_bokadd_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="14" >
						<col width="22%" >
						<col width="10%" >
						<col width="22%" >
						<col width="10%" >
						<col width="22" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="app_empno" type="text" id="app_empno" size="14" value="<%=app_empno%>" readonly="true"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="app_emp_name" type="text" id="app_emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    <%
                         if app_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&app_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
							  emp_job = Rs_emp("emp_job")
		                      emp_position = Rs_emp("emp_position")
							  emp_org_code = Rs_emp("emp_org_code")
							  emp_org_name = Rs_emp("emp_org_name")
							  emp_company = Rs_emp("emp_company")
							  emp_bonbu = Rs_emp("emp_bonbu")
							  emp_saupbu = Rs_emp("emp_saupbu")
							  emp_team = Rs_emp("emp_team")
							  emp_reside_place = Rs_emp("emp_reside_place")
		                   end if
	                       Rs_emp.Close()
	                	  end if	
				    %>	
                      <th style="background:#FFFFE6">직급/직책</th>                      
                      <td class="left" bgcolor="#FFFFE6"><%=emp_grade%>&nbsp;-&nbsp;<%=emp_position%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>현소속</th>                      
                      <td class="left"><%=emp_org_code%>&nbsp;-&nbsp;<%=emp_org_name%>&nbsp;</td>
                      <th>현조직</th>                      
                      <td colspan="3" class="left"><%=emp_company%>&nbsp;-&nbsp;<%=emp_bonbu%>&nbsp;-&nbsp;<%=emp_saupbu%>&nbsp;-&nbsp;<%=emp_team%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>휴직발령일</th>                      
                      <td class="left"><%=apphu_date%>&nbsp;</td>                      
                      <th>휴직유형</th>                      
                      <td class="left"><%=apphu_id_type%>&nbsp;</td>
                      <th>휴직기간</th>                      
                      <td class="left"><%=apphu_start_date%>&nbsp;∼&nbsp;<%=apphu_finish_date%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>휴직사유</th>
                      <td colspan="5" class="left"><%=apphu_comment%>&nbsp;</td>
                    </tr>
                    <tr>
                      <th>복직발령일</th>
                      <td colspan="5" class="left">
					  <input name="app_bok_date" type="text" value="<%=app_bok_date%>" style="width:80px;text-align:center" id="datepicker"></td>
                    </tr>
                    <tr>
                      <th>복직내용</th>
                      <td colspan="5" class="left">
					  <input name="app_comment" type="text" id="app_comment" style="width:300px; ime-mode:active" onKeyUp="checklength(this,50);" value="<%=app_comment%>">
                      </td>
                    </tr>
                    <tr>
                      <th>No.</th>  
					  <td colspan="5" class="left"><%=app_bok_seq%><input name="app_bok_seq" type="hidden" value="<%=app_bok_seq%>"></td>
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
                <input type="hidden" name="app_grade" value="<%=emp_grade%>" ID="Hidden1">
                <input type="hidden" name="app_position" value="<%=emp_position%>" ID="Hidden1">
                <input type="hidden" name="app_job" value="<%=emp_job%>" ID="Hidden1">
                <input type="hidden" name="app_company" value="<%=emp_company%>" ID="Hidden1">
                <input type="hidden" name="app_bonbu" value="<%=emp_bonbu%>" ID="Hidden1">
                <input type="hidden" name="app_saupbu" value="<%=emp_saupbu%>" ID="Hidden1">
                <input type="hidden" name="app_team" value="<%=emp_team%>" ID="Hidden1">
                <input type="hidden" name="app_org" value="<%=emp_org_code%>" ID="Hidden1">
                <input type="hidden" name="app_org_name" value="<%=emp_org_name%>" ID="Hidden1">
                <input type="hidden" name="apphu_seq" value="<%=apphu_seq%>" ID="Hidden1">
                <input type="hidden" name="apphu_id_type" value="<%=apphu_id_type%>" ID="Hidden1">
                <input type="hidden" name="apphu_date" value="<%=apphu_date%>" ID="Hidden1">
                <input type="hidden" name="apphu_start_date" value="<%=apphu_start_date%>" ID="Hidden1">
                <input type="hidden" name="apphu_finish_date" value="<%=apphu_finish_date%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

