<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim win_sw

'emp_no = request("emp_no")
'emp_name = request("emp_name")

in_empno =""
in_name = ""
If Request.Form("in_empno")  <> "" Then 
  in_empno = Request.Form("in_empno") 
End If

win_sw = "close"
Page=Request("page")

ck_sw=Request("ck_sw")

If ck_sw = "y" Then
	field_check=Request("field_check")
	field_view=Request("field_view")
	page_cnt=Request("page_cnt")

Else
	field_check=Request.form("field_check")
	field_view=Request.form("field_view")
	page_cnt=Request.form("page_cnt")
End if


pgsize = 10 ' 화면 한 페이지 
If Page = "" Then
	Page = 1
	start_page = 1
End If

stpage = int((page - 1) * pgsize)

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

sql = "select * from emp_family_event ORDER BY fm_id,fm_type ASC"
Rs.Open Sql, Dbconn, 1

'response.write sql

title_line = " 경조금 지급 규정 "
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
				return "4 1";
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
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_welfare_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_family_event_mg.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="15%" >
							<col width="10%" >
							<col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="10%" >
                            <col width="5%" >
						</colgroup>
						<thead>
                            <tr>
                                <th>경조구분</th>
                                <th>경조유형</th>
                                <th>경조회<br>경조금</th>
                                <th style="background:#F5FFFA">축(부)의금</th>
                                <th style="background:#F5FFFA">휴가일수<br>유급</th>
                                <th style="background:#F5FFFA">휴가일수<br>무급</th>
                                <th style="background:#F5FFFA">화환<br>조화</th>
                                <th style="background:#F5FFFA">꽃다발</th>
                                <th>수정</th>
                            </tr>
                        </thead>
						<tbody>
						<%
						do until rs.eof
						fm_wreath = ""
						fm_flowers= ""
						%>
							<tr>
                              <td><%=rs("fm_id")%>&nbsp;</td>
                              <td><%=rs("fm_type")%>&nbsp;</td>
                              <td style="text-align:right"><%=formatnumber(clng(rs("fm_sawo_pay")),0)%>&nbsp;</td>
                              <td style="text-align:right"><%=formatnumber(clng(rs("fm_company_pay")),0)%>&nbsp;</td>
                              <td><%=rs("fm_holiday1")%>&nbsp;</td>
                              <td><%=rs("fm_holiday2")%>&nbsp;</td>
                              <% If rs("fm_wreath_yn") = "Y" then fm_wreath = "지급" end if %>
                              <% If rs("fm_wreath_yn") = "N" then fm_wreath = "" end if %>
                              <td><%=fm_wreath%>&nbsp;</td>
                              <% If rs("fm_flowers_yn") = "Y" then fm_flowers = "지급" end if %>
                              <% If rs("fm_flowers_yn") = "N" then fm_flowers = "" end if %>
                              <td><%=fm_flowers%>&nbsp;</td>
							  <td><a href="#" onClick="pop_Window('insa_fm_event_add.asp?fm_id=<%=rs("fm_id")%>&fm_type=<%=rs("fm_type")%>&u_type=<%="U"%>','insa_fm_event_add_pop','scrollbars=yes,width=750,height=350')">수정</a></td>
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
					<a href="#" onClick="pop_Window('insa_fm_event_add.asp?fm_id=<%=fm_id%>&fm_type=<%=fm_type%>','insa_fm_event_add_pop','scrollbars=yes,width=750,height=350')" class="btnType04">경조금지급 규정 등록</a>
					<% if end_view = "Y" then %>
					<a href="payment_slip_end.asp?from_date=<%=from_date%>&to_date=<%=to_date%>&over_cash=<%=over_cash%>&use_cash=<%=use_cash%>" class="btnType04">전표마감</a>
					<% end if %>
					<% if user_id = "jinhs" then %>
					<a href="payment_slip_end_cancle.asp?from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">마감취소</a>
					<% end if %>
					</div>                  
                    </td>
			      </tr>
				  </table>
                <input type="hidden" name="fm_id" value="<%=fm_id%>" ID="Hidden1">
                <input type="hidden" name="fm_type" value="<%=fm_type%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

