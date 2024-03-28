<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")					
DbConn.Open dbconnect

curr_date = cstr(mid(now(),1,4)) + "-01-01"
sql = "select * from holiday where holiday >= '" + curr_date + "' order by holiday desc"
Rs.Open Sql, Dbconn, 1

title_line = "휴일관리"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "5 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=holiday%>" );
			});	  
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function form_chk(){				
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.holiday.value =="") {
					alert('일자를 입력하세요');
					frm.holiday.focus();
					return false;}

				if(document.frm.holiday_memo.value =="") {
					alert('휴일명을 입력하세요');
					frm.holiday_memo.focus();
					return false;}
			
				a=confirm('등록하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/code_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="49%" height="356" valign="top">
					  <form action="holi_del_ok.asp" method="post" name="frm_del">
                      <table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="*" >
				          <col width="30%" >
				          <col width="20%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">일자</th>
				            <th scope="col">휴일명</th>
				            <th scope="col">삭제</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
                        do until rs.eof
                        %>
				        <tr>
				          <td class="first"><%=rs("holiday")%></td>
				          <td><%=rs("holiday_memo")%></td>
				          <td><input name="del_ck" type="checkbox" id="del_ck" value="<%=rs("holiday")%>"></td>
			            </tr>
				        <%
							rs.movenext()
						loop
						%>
			            </tbody>
			          </table>
					  <br>
				      <div align=right>
                      	<a href="#" onclick="javascript:delcheck();" class="btnType04">선택 항목 삭제</a>
					  </div>
                      </form>
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="49%" valign="top"><form method="post" name="frm" action="holi_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				          <tbody>
				            <tr>
				              <th width="25%">날자</th>
				              <td class="left"><input name="holiday" type="text" id="datepicker" style="width:70px" readonly="true"></td>
			                </tr>
				            <tr>
				              <th>휴일명</th>
				              <td class="left"><input name="holiday_memo" type="text" id="holiday_memo" onKeyUp="checklength(this,20)" notnull errname="휴일명" style="width:150px"></td>
			                </tr>
			              </tbody>
			            </table>
						<br>
				        <input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				        <div align=center>
                        	<span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        </div>
			          </form></td>
			        </tr>
				    <tr>
				      <td width="49%">&nbsp;</td>
				      <td width="2%">&nbsp;</td>
				      <td width="49%">&nbsp;</td>
			        </tr>
			      </table>
                </div>
			</div>				
	</div>        				
	</body>
</html>

