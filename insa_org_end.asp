<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

'emp_no = request("emp_no")
'emp_name = request("emp_name")

in_orgcode =""
in_orgname = ""
If Request.Form("in_orgcode")  <> "" Then 
  in_orgcode = Request.Form("in_orgcode") 
End If

win_sw = "close"
Page=Request("page")

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	in_orgcode=Request("in_orgcode")
Else
	in_orgcode=Request.form("in_orgcode")
End if


pgsize = 10 ' 화면 한 페이지 
'pgsize = page_cnt ' 화면 한 페이지 

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

If Request.Form("in_orgcode")  <> "" Then 
   Sql = "SELECT * FROM emp_org_mst where org_code = '"+in_orgcode+"'"
   Set Rs_org = DbConn.Execute(SQL)
   in_orgname = Rs_org("org_name")
   rs_org.close()
End If

'sql = "select * from emp_org_mst where org_code = '" + in_orgcode + "' ORDER BY org_company,org_bonbu,org_saupbu,org_team,org_name ASC"
sql = "SELECT * "&_
      "  FROM emp_org_mst "&_
      " WHERE org_code = '" + in_orgcode + "' "&_
      "   AND (org_end_date='' OR isNull(org_end_date) OR org_end_date = '1900-01-01' OR org_end_date = '0000-00-00')"

Rs.Open Sql, Dbconn, 1

'Response.write sql

title_line = "조직 폐쇄"
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
				if (document.frm.in_orgcode.value == "") {
					alert ("조직코드을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_org_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_org_end.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
							<strong>조직코드 : </strong>
								<label>
        						<input name="in_orgcode" type="text" id="in_orgcode" value="<%=in_orgcode%>" style="width:100px; text-align:left">
								</label>
                            <strong>조직명 : </strong>
                                <label>
                               	<input name="in_orgname" type="text" id="in_orgname" value="<%=in_orgname%>" readonly="true" style="width:150px; text-align:left">
								</label>
                                
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
			    	    <colgroup>
			    	      <col width="4%" >
			    	      <col width="11%" >
		    		      <col width="6%" >
		    		      <col width="8%" >
                          <col width="11%" >
		    		      <col width="11%" >
		    		      <col width="11%" >
		    		      <col width="11%" >
			    	      <col width="8%" >
                          <col width="6%" >
				          <col width="8%" >
                          <col width="3%" >
			            </colgroup>
				    <thead>
				      <tr>
				        <th colspan="2" class="first" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;직</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">조&nbsp;&nbsp;직&nbsp;&nbsp;장</th>
                        <th colspan="4" scope="col" style=" border-bottom:1px solid #e3e3e3;">소&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;속</th>
				        <th rowspan="2" scope="col">조직생성일</th>
				        <th colspan="2" scope="col" style=" border-bottom:1px solid #e3e3e3;">상위&nbsp;조직장</th>
                        <th rowspan="2" scope="col">폐쇄</th>
			          </tr>
                      <tr>
				        <th class="first"scope="col">코드</th>
				        <th scope="col">조직명</th>
				        <th scope="col">사번</th>
				        <th scope="col">성명</th>
                        <th scope="col">회&nbsp;&nbsp;사</th>
				        <th scope="col">본&nbsp;&nbsp;부</th>
				        <th scope="col">사업부</th>
				        <th scope="col">팀</th>
				        <th scope="col">사번</th>
                        <th scope="col">성명</th>
                      </tr>
                     </thead>
						<tbody>
						<%
						do until rs.eof
						%>
							<tr>
 				        <td class="first"><%=rs("org_code")%>&nbsp;</td>
                        <td><%=rs("org_name")%>&nbsp;</td>
                        <td><%=rs("org_empno")%>&nbsp;</td>
                        <td><%=rs("org_emp_name")%>&nbsp;</td>
                        <td><%=rs("org_company")%>&nbsp;</td>
				        <td><%=rs("org_bonbu")%>&nbsp;</td>
                        <td><%=rs("org_saupbu")%>&nbsp;</td>
                        <td><%=rs("org_team")%>&nbsp;</td>
                        <td><%=rs("org_date")%>&nbsp;</td>
                        <td><%=rs("org_owner_empno")%>&nbsp;</td>
                        <td><%=rs("org_owner_empname")%>&nbsp;</td>
                        <td><a href="#" onClick="pop_Window('insa_orgend_reg.asp?org_code=<%=rs("org_code")%>&org_name=<%=in_name%>&u_type=<%="U"%>','insa_orgend_reg_pop','scrollbars=yes,width=1250,height=400')">등록</a>&nbsp;</td>
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

					</div>                  
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="org_code" value="<%=in_orgcode%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

