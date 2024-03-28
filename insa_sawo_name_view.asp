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
	owner_view=Request.form("owner_view")
	view_condi = request.form("view_condi")
  else
	owner_view=request("owner_view")
	view_condi = request("view_condi")
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
	         sql = "select * from emp_sawo_mem where sawo_emp_name like '%"+view_condi+"%' ORDER BY sawo_empno ASC"
         else
	         sql = "select * from emp_sawo_mem where sawo_empno = '"+view_condi+"' ORDER BY sawo_empno ASC"
     end if
	 Rs.Open Sql, Dbconn, 1
end if

'Rs.Open Sql, Dbconn, 1


title_line = " 경조회 가입 현황 "
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
				return "8 1";
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
			
			function sawo_mem_del(val, val2) {

            if (!confirm("정말 삭제하시겠습니까 ?")) return;
            var frm = document.frm;
			document.frm.sawo_empno.value = val;
			document.frm.sawo_emp_name.value = val2;
		
            document.frm.action = "insa_sawo_mem_del.asp";
            document.frm.submit();
            }	
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_sawo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_sawo_name_view.asp?ck_sw=<%="n"%>" method="post" name="frm">
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
							<col width="6%" >
							<col width="6%" >
							<col width="4%" >
							<col width="4%" >
                            <col width="9%" >
                            <col width="9%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
							<col width="5%" >
							<col width="6%" >
                            <col width="5%" >
							<col width="6%" >
							<col width="5%" >
                            <col width="6%" >
                            <col width="4%" >
                            <col width="4%" >
                            <col width="4%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">사번</th>
								<th scope="col">성  명</th>
								<th scope="col">직급</th>
								<th scope="col">직책</th>
                                <th scope="col">회사</th>
                                <th scope="col">소속</th>
								<th scope="col">가입일</th>
								<th scope="col">가입구분</th>
								<th scope="col">탈퇴일</th>
                                <th scope="col">탈퇴구분</th>
                                <th scope="col">급여공제</th>
                                <th scope="col">납입횟수</th>
                                <th scope="col">납입금액</th>
                                <th scope="col">지급횟수</th>
                                <th scope="col">지급금액</th>
								<th colspan="3" scope="col">비고</th>
							</tr>
                        </thead>
						<tbody>
						<%
						if  view_condi <> "" then 
						do until rs.eof
		                  sawo_empno = rs("sawo_empno")
		                  sawo_emp_name = rs("sawo_emp_name")
		
                         if sawo_empno <> "" then
		                    Sql="select * from emp_master where emp_no = '"&sawo_empno&"'"
		                    Rs_emp.Open Sql, Dbconn, 1

		                   if not Rs_emp.eof then
                              emp_grade = Rs_emp("emp_grade")
		                      emp_position = Rs_emp("emp_position")
		                   end if
	                       Rs_emp.Close()
	                	 end if		
						%>
							<tr>
								<td class="first"><%=rs("sawo_empno")%></td>
                                <td>
                                <a href="#" onClick="pop_Window('insa_card00.asp?emp_no=<%=rs("sawo_empno")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&date_sw=<%=date_sw%>&page_cnt=<%=page_cnt%>','emp_card0_pop','scrollbars=yes,width=1250,height=650')"><%=rs("sawo_emp_name")%></a>
								</td>
                                <td><%=emp_grade%>&nbsp;</td>
                                <td><%=emp_position%>&nbsp;</td>
                                <td><%=rs("sawo_company")%>&nbsp;</td>
                                <td><%=rs("sawo_org_name")%>&nbsp;</td>
                                <td><%=rs("sawo_date")%>&nbsp;</td>
                                <td><%=rs("sawo_id")%>&nbsp;</td>
                                <td><%=rs("sawo_out_date")%>&nbsp;</td>
                                <td><%=rs("sawo_out")%>&nbsp;</td>
                                <% If rs("sawo_target") = "Y" then sawo_target = "공제" end if %>
                                <% If rs("sawo_target") = "N" then sawo_target = "안함" end if %>
								<td><%=sawo_target%>&nbsp;</td>
                                <td style="text-align:right">
                                <a href="#" onClick="pop_Window('insa_sawo_in_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','sawo_inview','scrollbars=yes,width=800,height=600')"><%=rs("sawo_in_count")%></a>
								</td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_in_pay")),0)%>&nbsp;</td>
                                <td style="text-align:right">
                                <a href="#" onClick="pop_Window('insa_sawo_give_view.asp?emp_no=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&be_pg=<%=be_pg%>&page=<%=page%>&view_sort=<%=view_sort%>&page_cnt=<%=page_cnt%>','sawo_inview','scrollbars=yes,width=1000,height=600')"><%=rs("sawo_give_count")%></a>
                                </td>
                                <td style="text-align:right"><%=formatnumber(clng(rs("sawo_give_pay")),0)%>&nbsp;</td>
                                <% if user_id = "sanginlee" then %>
                                <td colspan="3"><a href="#" onClick="pop_Window('insa_sawo_giveadd.asp?sawo_empno=<%=rs("sawo_empno")%>&emp_name=<%=rs("sawo_emp_name")%>&u_type=<%=""%>','insa_sawo_giveadd_pop','scrollbars=yes,width=750,height=500')">경조금지급</a>&nbsp;</td>
                                <%    else %>
                                <td colspan="3"><%=rs("sawo_out")%>&nbsp;</td>
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
                  <input type="hidden" name="sawo_empno" value="<%=sawo_empno%>" ID="Hidden1">
                  <input type="hidden" name="sawo_emp_name" value="<%=sawo_emp_name%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

