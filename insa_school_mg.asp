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
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

if view_condi <> "" then
     if owner_view = "C" then  
	      Sql= "select * " & _
	          "    from emp_school a, emp_master b " & _
	          "    where a.sch_empno = b.emp_no AND b.emp_name like '%" + view_condi + "%' " & _
		      "    ORDER BY sch_empno,sch_seq ASC" 
        else
	      sql = "select * from emp_school where sch_empno = '"+view_condi+"' ORDER BY sch_empno,sch_seq ASC"
     end if
	 Rs.Open Sql, Dbconn, 1
end if
'Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " 학력 사항 "
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
			
			function school_del(val, val2, val3, val4) {

            if (!confirm("정말 삭제하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.sch_empno.value = val;
			document.frm.sch_seq.value = val2;
			document.frm.sch_emp_name.value = val3;
			document.frm.owner_view.value = val4;
		
            document.frm.action = "insa_school_del.asp";
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
				<form action="insa_school_mg.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
                            <col width="12%" >
                            <col width="12%" >
							<col width="*" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="7%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>사번</th>
                                <th>성명</th>
                                <th>소속</th>
                                <th>기간</th>
                                <th>학교명</th>
                                <th>학과</th>
                                <th>전공</th>
                                <th>부전공</th>
                                <th>학위</th>
                                <th>졸업</th>
                                <th>학력</th>
                                <th>수정</th>
                                <th>비고</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						if  view_condi <> "" then 
						do until rs.eof
						      sch_empno = rs("sch_empno")
							  Sql = "SELECT * FROM emp_master where emp_no = '"&sch_empno&"'"
                              Set rs_emp = DbConn.Execute(SQL)
							  if not Rs_emp.eof then
                                   emp_company = rs_emp("emp_company")
								   emp_name = rs_emp("emp_name")
                                   emp_bonbu = rs_emp("emp_bonbu")
                                   emp_saupbu = rs_emp("emp_saupbu")
                                   emp_team = rs_emp("emp_team")
                                   emp_org_code = rs_emp("emp_org_code")
                                   emp_org_name = rs_emp("emp_org_name")
							  end if
							  rs_emp.close()						
						
						%>
							<tr>
                              <td><%=rs("sch_empno")%>&nbsp;</td>
                              <td><%=emp_name%>&nbsp;</td>
                              <td><%=emp_org_name%>(<%=emp_org_code%>)&nbsp;</td>
                              <td><%=rs("sch_start_date")%>∼<%=rs("sch_end_date")%>&nbsp;</td>
                              <td><%=rs("sch_school_name")%>&nbsp;</td>
                              <td><%=rs("sch_dept")%>&nbsp;</td>
                              <td><%=rs("sch_major")%>&nbsp;</td>
                              <td><%=rs("sch_sub_major")%>&nbsp;</td>
                              <td><%=rs("sch_degree")%>&nbsp;</td>
                              <td><%=rs("sch_finish")%>&nbsp;</td>
                              <td><a href="#" onClick="pop_Window('insa_school_add.asp?sch_empno=<%=rs("sch_empno")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%=""%>','insa_school_add_pop','scrollbars=yes,width=750,height=300')">등록</a></td>
							  <td><a href="#" onClick="pop_Window('insa_school_add.asp?sch_empno=<%=rs("sch_empno")%>&sch_seq=<%=rs("sch_seq")%>&emp_name=<%=emp_name%>&owner_view=<%=owner_view%>&u_type=<%="U"%>','insa_school_add_pop','scrollbars=yes,width=750,height=300')">수정</a></td>
                         <% if insa_grade = "0" then %>     
                              <td>
                              <a href="#" onClick="school_del('<%=rs("sch_empno")%>', '<%=rs("sch_seq")%>', '<%=emp_name%>', '<%=owner_view%>');return false;">삭제</a></td>
                         <%     else %>
                              <td>&nbsp;</td>
                         <% end if %>                              
                              
							</tr>
						<%
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
                    <a href="#" onClick="pop_Window('insa_school_add.asp?sch_empno=<%=view_condi%>&emp_name=<%=emp_name%>','insa_school_add_pop','scrollbars=yes,width=750,height=300')" class="btnType04">학력등록</a>
					<% end if %>
                    </div>                  
                    </td>
			      </tr>
				  </table>
                  <input type="hidden" name="sch_empno" value="<%=sch_empno%>" ID="Hidden1">
                  <input type="hidden" name="sch_seq" value="<%=sch_seq%>" ID="Hidden1">
                  <input type="hidden" name="sch_emp_name" value="<%=sch_emp_name%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

