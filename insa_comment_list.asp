<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
insa_grade = request.cookies("nkpmg_user")("coo_insa_grade")

view_condi = request("view_condi")
owner_view=request("owner_view")

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	view_condi = request.form("view_condi")
	owner_view=Request.form("owner_view")
  else
	view_condi = request("view_condi")
	owner_view=request("owner_view")
end if

if view_condi = "" then
	view_condi = ""
	owner_view = "C"
	ck_sw = "n"
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_cmt = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi <> "" then
     if owner_view = "C" then  
	     Sql= "select * " & _
	          "    from emp_master " & _
	          "    where emp_name like '%" + view_condi + "%' " & _
		      "    ORDER BY emp_no ASC" 
       else
	     sql = "select * from emp_master where emp_no = '"+view_condi+"' ORDER BY emp_no ASC"
     end if
	 Rs.Open Sql, Dbconn, 1
end if
'Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " 인사 특이사항 "
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
				return "1 1";
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
				if (document.frm.view_condi.value == "") {
					alert ("조건을 입력하시기 바랍니다");
					return false;
				}	
				return true;
			}
			function comment_del(val, val2, val3, val4) {

            if (!confirm("정말 삭제하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.cmt_empno.value = val;
			document.frm.cmt_date.value = val2;
			document.frm.cmt_empname.value = val3;
			document.frm.owner_view.value = val4;
		
            document.frm.action = "insa_comment_del.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sub_menu1.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_comment_list.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>◈조건 검색◈</dt>
                        <dd>
                            <p>
                                <label>
                                <input name="owner_view" type="radio" value="T" <% if owner_view = "T" then %>checked<% end if %> style="width:25px">사번
                                <input name="owner_view" type="radio" value="C" <% if owner_view = "C" then %>checked<% end if %> style="width:25px">성명
                                </label>
							<strong>조건 : </strong>
								<label>
        						<input name="view_condi" type="text" id="view_condi" value="<%=view_condi%>" style="width:100px; text-align:left">
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
                            <col width="11%" >
                            <col width="9%" >
                            <col width="12%" >
							<col width="27%" >
                            <col width="*" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>사번</th>
                                <th>성명</th>
                                <th>현소속</th>
                                <th>발생일</th>
                                <th>당시소속</th>
                                <th>당시조직</th>
                                <th>특이사항</th>
                                <th>등록</th>
                                <th>수정</th>
                                <th>비고</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						if  view_condi <> "" then 
						do until rs.eof
						      cmt_empno = rs("emp_no")
							  Sql = "SELECT * FROM emp_comment where cmt_empno = '"&cmt_empno&"'"
                              Set rs_cmt = DbConn.Execute(SQL)
							  if not rs_cmt.eof then
                                   emp_company = rs_cmt("cmt_company")
								   emp_name = rs_cmt("cmt_emp_name")
                                   emp_bonbu = rs_cmt("cmt_bonbu")
                                   emp_saupbu = rs_cmt("cmt_saupbu")
                                   emp_team = rs_cmt("cmt_team")
                                   emp_org_code = rs_cmt("cmt_org_code")
                                   emp_org_name = rs_cmt("cmt_org_name")
								   cmt_date = rs_cmt("cmt_date")
								   task_memo = replace(rs_cmt("cmt_comment"),chr(34),chr(39))
							       view_memo = task_memo
							       if len(task_memo) > 16 then
							    	   view_memo = mid(task_memo,1,16) + "..."
							       end if	
								 else
								   emp_company = rs("emp_company")
								   emp_name = rs("emp_name")
                                   emp_bonbu = rs("emp_bonbu")
                                   emp_saupbu = rs("emp_saupbu")
                                   emp_team = rs("emp_team")
                                   emp_org_code = rs("emp_org_code")
                                   emp_org_name = rs("emp_org_name")
								   task_memo = ""
								   view_memo = ""
							  end if
					
						%>
							<tr>
                              <td><%=rs("emp_no")%>&nbsp;</td>
                              <td><%=emp_name%>&nbsp;</td>
                              <td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
                              <td><%=cmt_date%>&nbsp;</td>
                              <td class="left"><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
                              <td class="left"><%=emp_company%>-<%=emp_bonbu%>-<%=emp_saupbu%>-<%=emp_team%>&nbsp;</td>
                              <td class="left"><p style="cursor:pointer"><span title="<%=task_memo%>"><%=view_memo%></span></p></td>
                              <td ><a href="#" onClick="pop_Window('insa_comment_add.asp?cmt_empno=<%=cmt_empno%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%=""%>','insa_comment_add_pop','scrollbars=yes,width=750,height=250')">등록</a>
                              </td>
                         <% if insa_grade = "0" then %>     
                              <td ><a href="#" onClick="pop_Window('insa_comment_add.asp?cmt_empno=<%=cmt_empno%>&cmt_date=<%=cmt_date%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%="U"%>','insa_comment_add_pop','scrollbars=yes,width=750,height=250')">수정</a>
                              </td>
                              <td>
                              <a href="#" onClick="comment_del('<%=cmt_empno%>', '<%=cmt_date%>', '<%=emp_name%>', '<%=owner_view%>');return false;">삭제</a></td>
                         <%     else %>
                              <td>&nbsp;</td>
                              <td>&nbsp;</td>
                         <% end if %>          
							</tr>
						<%
							rs_cmt.close()
							rs.movenext()
						loop
						rs.close()
						
						end if
						%>
						</tbody>
					</table>
				</div>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <% if owner_view = "T" then 
                              emp_no = view_condi
							  Sql = "SELECT * FROM emp_master where emp_no = '"&emp_no&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
							  end if
							  rs_emp.close()
				    %>
					<a href="#" onClick="pop_Window('insa_comment_add.asp?cmt_empno=<%=view_condi%>&emp_name=<%=emp_name%>&u_type=<%=""%>','insa_comment_add_pop','scrollbars=yes,width=750,height=250')" class="btnType04">특이사항등록</a>
                    <% end if %>
					</div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="cmt_empno" value="<%=cmt_empno%>" ID="Hidden1">
                  <input type="hidden" name="cmt_date" value="<%=cmt_date%>" ID="Hidden1">
                  <input type="hidden" name="cmt_empname" value="<%=cmt_empname%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

