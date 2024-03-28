<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")
lang_empno = request("lang_empno")
lang_seq = request("lang_seq")
emp_name = request("emp_name")

lang_id = ""
lang_id_type = ""
lang_point = ""
lang_grade = ""
lang_get_date = ""

curr_date = mid(cstr(now()),1,10)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = " 어학능력 등록 "
if u_type = "U" then

	Sql="select * from emp_language where lang_empno = '"&lang_empno&"' and lang_seq = '"&lang_seq&"'"
	Set rs=DbConn.Execute(Sql)

	lang_id = rs("lang_id")
    lang_id_type = rs("lang_id_type")
    lang_point = rs("lang_point")
    lang_grade = rs("lang_grade")
    lang_get_date = rs("lang_get_date")
	
	rs.close()

	title_line = " 어학능력 변경 "
	
end if

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=lang_get_date%>" );
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
				if(document.frm.lang_id.value =="") {
					alert('어학구분을 선택하세요');
					frm.lang_id.focus();
					return false;}
				if(document.frm.lang_id_type =="") {
					alert('어학종류을 선택하세요');
					frm.lang_id_type.focus();
					return false;}
				if(document.frm.lang_grade.value =="") {
					alert('급수를 입력하세요');
					frm.lang_grade.focus();
					return false;}
				if(document.frm.lang_point.value =="") {
					alert('점수를 입력하세요');
					frm.lang_point.focus();
					return false;}
				if(document.frm.lang_get_date.value =="") {
					alert('취득일을 입력하세요');
					frm.lang_get_date.focus();
					return false;}
				if(document.frm.lang_get_date.value > document.frm.curr_date.value) {
						alert('취득일이 현재일보다 빠릅니다');
						frm.lang_get_date.focus();
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
				<form action="insa_language_add_save.asp" method="post" name="frm">
				<div class="gView">
				  <table cellpadding="0" cellspacing="0" class="tableWrite">
                  	<colgroup>
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="22%" >
						<col width="11%" >
						<col width="*" >
					</colgroup>
				    <tbody>
                    <tr>
                      <th style="background:#FFFFE6">사번</th>
                      <td class="left" bgcolor="#FFFFE6">
					  <input name="lang_empno" type="text" id="lang_empno" size="14" value="<%=lang_empno%>" readonly="true">
                      <input type="hidden" name="lang_seq" value="<%=lang_seq%>" ID="Hidden1"></td>
                      <th style="background:#FFFFE6">성명</th>
                      <td colspan="3" class="left" bgcolor="#FFFFE6">
					  <input name="emp_name" type="text" id="emp_name" size="14" value="<%=emp_name%>" readonly="true"></td>
                    </tr>
                    <tr>  
                      <th>어학구분</th>
                      <td colspan="2" class="left">
                    <%
					  Sql="select * from emp_etc_code where emp_etc_type = '08' order by emp_etc_code asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="lang_id" id="lang_id" style="width:90px">
                         <option value="" <% if lang_id = "" then %>selected<% end if %>>선택</option>
                			  <% 
								do until rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If lang_id = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							  %>
            		  </select>       
                      </td>
                    </tr>
                    <tr>  
                      <th>어학종류</th>
                      <td class="left">
                    <%
					  Sql="select * from emp_etc_code where emp_etc_type = '09' order by emp_etc_code asc"
					  Rs_etc.Open Sql, Dbconn, 1
					%>
					  <select name="lang_id_type" id="lang_id_type" style="width:90px">
                         <option value="" <% if lang_id_type = "" then %>selected<% end if %>>선택</option>
                			  <% 
								do until rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("emp_etc_name")%>' <%If lang_id_type = rs_etc("emp_etc_name") then %>selected<% end if %>><%=rs_etc("emp_etc_name")%></option>
                			  <%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							  %>
            		  </select>                             
                      </td>
                      <th>급수</th>
                      <td class="left">
                      <select name="lang_grade" id="lang_grade" value="<%=lang_grade%>" style="width:100px">
			               <option value="" <% if lang_grade = "" then %>selected<% end if %>>선택</option>
				           <option value='급수었음' <%If lang_grade = "급수었음" then %>selected<% end if %>>급수었음</option>
                           <option value='3급' <%If lang_grade = "3급" then %>selected<% end if %>>3급</option>
                           <option value='2급' <%If lang_grade = "2급" then %>selected<% end if %>>2급</option>
                           <option value='1급' <%If lang_grade = "1급" then %>selected<% end if %>>1급</option>
                      </select>
                      </td>
                      <th>점수</th>
                      <td class="left">
                      <input name="lang_point" type="text" id="lang_point" style="width:80px; ime-mode:active" onKeyUp="checklength(this,3);" value="<%=lang_point%>"></td>
                    </tr>
                    <tr>
                      <th>취득일</th>
                      <td colspan="5" class="left">
					  <input name="lang_get_date" type="text" value="<%=lang_get_date%>" style="width:80px;text-align:center" id="datepicker">&nbsp;
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
                <input type="hidden" name="curr_date" value="<%=curr_date%>" ID="Hidden1">
				</form>
		</div>				
	</body>
</html>

