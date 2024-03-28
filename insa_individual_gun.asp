<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim Rs
Dim Repeat_Rows

be_pg = "insa_individual_gun.asp"

in_name = request.cookies("nkpmg_user")("coo_user_name")
in_empno = request.cookies("nkpmg_user")("coo_user_id")

curr_date = datevalue(mid(cstr(now()),1,10))
rever_yyyymm = mid(cstr(curr_date),1,7) '귀속년월

ck_sw=Request("ck_sw")

if ck_sw = "n" then
	in_year = request.form("in_year")
  else
	in_year = request("in_year")
end if

if in_year = "" then
	'inc_yyyy = mid(cstr(now()),1,4)
	in_year = cint(mid(now(),1,4)) - 1
	ck_sw = "n"
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_emp = Server.CreateObject("ADODB.Recordset")
Set Rs_use = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open DbConnect

Sql = "select * from emp_year_leave where year_empno = '"&in_empno&"' and year_year = '"&in_year&"'"
Rs.Open Sql, Dbconn, 1

If Not(Rs.bof Or Rs.eof) Then
	year_yuncha_date = rs("year_yuncha_date")
	year_continu_year = rs("year_continu_year")
	year_continu_month = rs("year_continu_month")
	' year_before_count = rs("year_before_count")
	year_basic_count = rs("year_basic_count")
	year_add_count = rs("year_add_count")
	year_leave_count = rs("year_leave_count")
	year_use_count = rs("year_use_count")
	' year_holi_count = rs("year_holi_count")
	' year_use_holi = rs("year_use_holi")
	before_remain_cnt = 0
	before_use_count = 0
	' if year_before_count = 0 then
	'         remain_cnt = year_leave_count - year_use_count
	'	else  
	'	     before_remain_cnt = year_before_count - year_use_count
	'		 before_use_count = year_use_count
	'  end if		 
	remain_holi = year_holi_count - year_use_holi
End If
Rs.close()

tot_yuncha = 0
tot_other = 0
sql = "select * from emp_holiday_use where holi_emp_no = '"&in_empno&"' and holi_year = '"&in_year&"'"
Rs_use.Open Sql, Dbconn, 1
do until Rs_use.eof
     if Rs_use("holi_type") = "반차" or Rs_use("holi_type") = "연차" then
	         tot_yuncha = tot_yuncha + int(Rs_use("holi_count"))
		else   
             tot_other = tot_other + int(Rs_use("holi_count"))
	 end if		 
   Rs_use.movenext()
loop
Rs_use.close()	

sql = "select * from emp_holiday_use where holi_emp_no = '"&in_empno&"' and holi_year = '"&in_year&"' ORDER BY holi_emp_no,holi_year,holi_start_date ASC"
Rs.Open Sql, Dbconn, 1

title_line = " 근태/휴·공가 현황"

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
				return "4 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=end_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function chkfrm() {
				if (document.frm.in_year.value == "") {
					alert ("년도 입력하시기 바랍니다");
					return false;
				}
				
				return true;
			}
			
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
		<div id="wrap">			
			<!--#include virtual = "/include/insa_pheader.asp" -->
			<!--#include virtual = "/include/insa_pappo_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_individual_gun.asp?ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>연차휴가 일수</dt>
                        <dd>
							<strong>년도: </strong>
								<label>
        						<input name="in_year" type="text" id="in_year" value="<%=in_year%>" style="width:40px; text-align:center">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
                            <col width="20%" >
						</colgroup>
						<thead>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">연차기산일</th>
                              <td style=" border-bottom:1px solid #e3e3e3;"><%=year_yuncha_date%>&nbsp;</td>
                              <th style=" border-bottom:1px solid #e3e3e3;">근속년수</th>
                              <td colspan="2" style=" border-bottom:1px solid #e3e3e3;"><%=year_continu_year%>&nbsp;년&nbsp;(<%=year_continu_month%>&nbsp;개월)</td>
						    </tr>
                            <tr>
							  <th style=" border-bottom:1px solid #e3e3e3;">&nbsp;</th>
                              <th >선연차</th>
                              <th >연차</th>
                              <th >정기휴가</th>
                              <th >기타휴가 및 공가</th>
						    </tr>
                            <tr>
                              <th style="border-bottom:1px solid #e3e3e3;">권한휴가</th>
                              <td ><%=year_before_count%>&nbsp;</td>
                              <td ><%=year_leave_count%>&nbsp;</td>
                              <td ><%=year_holi_count%>&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
                            <tr>
                              <th style="border-bottom:1px solid #e3e3e3;">사용한휴가</th>
                              <td ><%=before_use_count%>&nbsp;</td>
                              <td ><%=year_use_count%>&nbsp;</td>
                              <td ><%=year_use_holi%>&nbsp;</td>
                              <td ><%=tot_other%>&nbsp;</td>
						    </tr>
                            <tr>
                              <th style="border-bottom:1px solid #e3e3e3;">미사용휴가</th>
                              <td ><%=before_remain_cnt%>&nbsp;</td>
                              <td ><%=remain_cnt%>&nbsp;</td>
                              <td ><%=remain_holi%>&nbsp;</td>
                              <td >&nbsp;</td>
						    </tr>
						</thead>
						<tbody>
					</table>
                    <br>					
                    <table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
							<col width="8%" >
							<col width="6%" >
							<col width="6%" >
							<col width="6%" >
                            <col width="18%" >
                            <col width="*" >
							<col width="8%" >
                            <col width="6%" >
                            <col width="6%" >
                            <col width="12%" >
							<col width="4%" >
						</colgroup>
						<thead>
                            <tr>
								<th scope="col">신청일자</th>
								<th scope="col">휴가유형</th>
								<th scope="col">휴가시작일</th>
								<th scope="col">휴가종료일</th>
                                <th scope="col">휴가일수</th>
                                <th scope="col">소속</th>
                                <th scope="col">휴가사유</th>
                                <th scope="col">결재자</th>
                                <th scope="col">승인여부</th>
                                <th scope="col">결재일</th>
                                <th scope="col">결재의견</th>
                                <th scope="col">비고</th>
							</tr>
						</thead>
						<tbody>
						<%
						do until rs.eof
                            if rs("holi_sign_status") = "1" then
							        holi_sign_status = "진행"
							   elseif rs("holi_sign_status") = "2" then
							               holi_sign_status = "결재"
									  elseif rs("holi_sign_status") = "3" then
							               holi_sign_status = "반송"
						    end if
							holi_sign_date = rs("holi_sign_date")
	                        if holi_sign_date = "1900-01-01" then
	                              holi_sign_date = ""
	                        end if
	           			%>
							<tr>
                                <td><%=rs("holi_date")%>&nbsp;</td>
                                <td><%=rs("holi_type")%>&nbsp;</td>
                                <td><%=rs("holi_start_date")%>&nbsp;</td>
                                <td><%=rs("holi_end_date")%>&nbsp;</td>
                                <td><%=rs("holi_count")%>&nbsp;</td>
                                <td><%=rs("holi_org_name")%>&nbsp;(<%=rs("holi_saupnu")%>)</td>
                                <td class="left"><%=rs("holi_memo")%>&nbsp;</td>
                                <td><%=rs("holi_sing_empname")%>&nbsp;(<%=rs("holi_sign_empno")%>)</td>
                                <td><%=holi_sign_status%>&nbsp;</td>
                                <td><%=holi_sign_date%>&nbsp;</td>
                                <td class="left"><%=rs("holi_sign_memo")%>&nbsp;</td>
                        <% if  rs("holi_sign_status") = "1" or rs("holi_sign_status") = "3" then %>
                                <td><a href="#" onClick="pop_Window('insa_holiday_add.asp?in_empno=<%=rs("holi_emp_no")%>&in_year=<%=rs("holi_year")%>&holi_date=<%=rs("holi_date")%>&emp_name=<%=rs("holi_emp_name")%>&u_type=<%="U"%>','insa_holiday_add_pop','scrollbars=yes,width=850,height=370')">수정</a></td>
						<%    else  %>         
                                <td>&nbsp;</td>
                        <% end if %>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						%>
						</tbody>
					</table>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
				  <tr>
				    <td>
					<div class="btnRight">
                    <a href="#" onClick="pop_Window('insa_holiday_add.asp?in_year=<%=in_year%>&in_empno=<%=in_empno%>&u_type=<%=""%>','insa_holiday_add_pop','scrollbars=yes,width=850,height=370')" class="btnType04">휴가신청 등록</a>
					</div> 
                    </td>
			      </tr>
				  </table>                    
				</div>
                  <input type="hidden" name="emp_empno" value="<%=in_empno%>" ID="Hidden1">
			</form>
		</div>				
	</div>        				
	</body>
</html>

