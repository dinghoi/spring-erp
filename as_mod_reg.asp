<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
curr_date = mid(cstr(now()),1,10)

acpt_no = request("acpt_no")
be_pg = request("be_pg")

page = request("page")
from_date = request("from_date")
to_date = request("to_date")
date_sw = request("date_sw")
process_sw = request("process_sw")
field_check = request("field_check")
field_view = request("field_view")
condi_com = request("company")


Set DbConn = Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_ddd = Server.CreateObject("ADODB.Recordset")
DbConn.Open dbconnect

Sql = "select * from as_acpt where acpt_no = "&int(acpt_no)
Set rs = DbConn.Execute(SQL)

acpt_date = mid(cstr(rs("acpt_date")),1,10)
acpt_hh = int(datepart("h",rs("acpt_date")))
acpt_mm = int(datepart("n",rs("acpt_date")))
acpt_ss = datepart("s",rs("acpt_date"))

if acpt_hh < 10 then
	acpt_hh = "0" + cstr(acpt_hh)
end if

if acpt_mm < 10 then
	acpt_mm = "0" + cstr(acpt_mm)
end if

if arrival_time = "0000" or arrival_time = null or arrival_time = "" then
	arrival_date = curr_date
	arrival_time = "0000"
end if

title_line = "A/S 내역 수정"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 내역 수정</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}
			function goBefore () {
				window.close () ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.acpt_user.value =="") {
					alert('사용자를 입력하세요');
					frm.acpt_user.focus();
					return false;}
				if(document.frm.tel_ddd.value =="") {
					alert('전화번호를 입력하세요');
					frm.tel_ddd.focus();
					return false;}
				if(document.frm.tel_no1.value =="") {
					alert('전화번호를 입력하세요');
					frm.tel_no1.focus();
					return false;}
				if(document.frm.tel_no2.value =="") {
					alert('전화번호를 입력하세요');
					frm.tel_no2.focus();
					return false;}
				if(document.frm.dept.value =="") {
					alert('조직명을 입력하세요');
					frm.dept.focus();
					return false;}
				if(document.frm.sido.value =="") {
					alert('주소록을 등록하세요');
					frm.area_view.focus();
					return false;}
				if(document.frm.gugun.value =="") {
					alert('주소록을 등록하세요');
					frm.area_view.focus();
					return false;}
				if(document.frm.dong.value =="") {
					alert('주소록을 등록하세요');
					frm.area_view.focus();
					return false;}
				if(document.frm.addr.value =="") {
					alert('나머지 주소를 입력하세요');
					frm.addr.focus();
					return false;}
				if(document.frm.as_memo.value =="") {
					alert('장애내용을 입력하세요');
					frm.as_memo.focus();
					return false;}
			
				if(document.frm.request_date.value =="") {
					alert('요청일을 입력하세요');
					frm.request_date.focus();
					return false;}
				/**/	
				if(document.frm.request_date.value < document.frm.acpt_date.value) {
					alert('요청일이 접수일보다 빠름니다');
					frm.request_date.focus();
					return false;}
				if(document.frm.request_hh.value >"23"||document.frm.request_hh.value <"00") {
					alert('요청시간이 잘못되었습니다');
					frm.request_hh.focus();
					return false;}
				if(document.frm.request_mm.value >"59"||document.frm.request_mm.value <"00") {
					alert('요청분이 잘못되었습니다');
					frm.request_mm.focus();
					return false;}
				/**/
				if(document.frm.request_date.value == document.frm.acpt_date.value) {
					if(document.frm.request_hh.value < document.frm.acpt_hh.value) {
						alert('요청시간이 접수시간 보다 빠름니다');
						frm.request_hh.focus();
						return false;}}
				if(document.frm.request_date.value == document.frm.acpt_date.value) {
					if(document.frm.request_hh.value == document.frm.acpt_hh.value) {
						if(document.frm.request_mm.value <= document.frm.acpt_mm.value) {
							alert('요청분이 접수분 보다 빠름니다');
							frm.request_mm.focus();
							return false;}}}
							
				{
				a=confirm('등록하시겠습니까?');
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function ce_mod_view() 
			{
				if (document.frm.ce_mod_ck.checked == true) {
					document.getElementById('s_ce').style.display = ''; 
					document.getElementById('ce_mod').style.display = ''; }
				if (document.frm.ce_mod_ck.checked == false) {
					document.getElementById('ce_mod').style.display = 'none'; 
					document.getElementById('s_ce').style.display = 'none'; }
			}
			$(function() {    $( "#datepicker" ).datepicker();
											$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
											$( "#datepicker" ).datepicker("setDate", "<%=rs("request_date")%>" );
			});	  
        </script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="as_mod_reg_ok.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
						<colgroup>
							<col width="12%" >
							<col width="20%" >
							<col width="11%" >
							<col width="*" >
							<col width="11%" >
							<col width="19%" >
						</colgroup>
						<tbody>
							<tr>
							  <th>접수번호</th>
							  <td class="left"><%=rs("acpt_no")%></td>
							  <th>접수일자</th>
							  <td class="left"><%=rs("acpt_date")%></td>
							  <th>접수자</th>
							  <td class="left"><%=rs("acpt_man")%></td>
					    </tr>
							<tr>
								<th>사용자/직급</th>
							  <td class="left">
							  	<input name="acpt_user" type="text" id="acpt_user" value="<%=rs("acpt_user")%>" size="10">
                	<input name="user_grade" type="text" id="user_grade" value="<%=rs("user_grade")%>" size="6">
                </td>
							  <th>전화번호</th>
							  <td class="left">
								<% 
									Sql="select * from etc_code where etc_type = '71' and used_sw = 'Y' order by etc_code asc"
                  Rs_ddd.Open Sql, Dbconn, 1
                %>
                	<select name="tel_ddd" id="select3">
                  <% 
                  	do until rs_ddd.eof 
                  %>
                  	<option value='<%=rs_ddd("etc_name")%>' <%If rs_ddd("etc_name") = rs("tel_ddd") then %>selected<% end if %>><%=rs_ddd("etc_name")%></option>
                  <%
                  		rs_ddd.movenext()
                  		loop
                  		rs_ddd.close()						
                  %>
                  </select>
                  -
                  <input name="tel_no1" type="text" id="tel_no1" value="<%=rs("tel_no1")%>" size="4" maxlength="4">
                  -
                  <input name="tel_no2" type="text" id="tel_no2" value="<%=rs("tel_no2")%>" size="4" maxlength="4">
                </td>
							  <th>핸드폰</th>
							  <td class="left">
                	<input name="hp_ddd" type="text" id="hp_ddd" value="<%=rs("hp_ddd")%>" size="3" maxlength="3"> 
                	-
                	<input name="hp_no1" type="text" id="hp_no1" value="<%=rs("hp_no1")%>" size="4" maxlength="4">
                	-
                	<input name="hp_no2" type="text" id="hp_no2" value="<%=rs("hp_no2")%>" size="4" maxlength="4">
                </td>
              </tr>
							<tr>
							  <th>회사명</th>
							  <td class="left">
								<%
									sql="select * from trade where use_sw = 'Y' and mg_group = '" + mg_group + "' order by trade_name asc"
									Rs_etc.Open Sql, Dbconn, 1
                %>
                	<select name="company" id="company">
                  <% 
                  	do until rs_etc.eof 
                  %>
                  	<option value='<%=rs_etc("trade_name")%>' <%If rs_etc("trade_name") = rs("company")  then %>selected<% end if %>><%=rs_etc("trade_name")%></option>
                  <%
                  		rs_etc.movenext()  
                      loop 
                      rs_etc.Close()
                  %>
                  </select>
                </td>
							  <th>조직명</th>
							  <td class="left" colspan="3"><input name="dept" type="text" id="dept" onKeyUp="checklength(this,50)" value="<%=rs("dept")%>" size="30"></td>
					    </tr>
							<tr>
							  <th>주소</th>
							  <td class="left" colspan="5">
							  	<input name="sido" type="text" id="sido3" value="<%=rs("sido")%>" size="6" readonly="true">
                  <input name="gugun" type="text" id="gugun4" value="<%=rs("gugun")%>" size="20" readonly="true">
                  <input name="dong" type="text" value="<%=rs("dong")%>" size="18" readonly="true">
                  <input name="addr" type="text" id="addr" value="<%=rs("addr")%>" size="40" onKeyUp="checklength(this,50)">
                  <input name="view_ok" type="hidden" id="view_ok" value="">
              		<a href="#" class="btnType03" onclick="javascript:pop_area()" >지역조회</a>
                </td>
					    </tr>
							<tr>
							  <th>기존CE</th>
							  <td class="left">
                	<input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=rs("mg_ce_id")%>">
                  <input name="mg_ce" type="text" id="mg_ce" value="<%=rs("mg_ce")%>" size="10" readonly="true">                  
                  <input name="reside_place" type="hidden" id="reside_place" value="">
                  <input name="team" type="hidden" id="team">
                </td>
								<th>CE변경</th>
							  <td class="left">
                  <input name="ce_mod_ck" type="checkbox" id="ce_mod_ck" value="1"  onClick="ce_mod_view()">
                  <input name="s_ce_id" type="hidden" value="<%=user_id%>">
                  <input name="s_ce" type="text" value="<%=user_name%>" size="8" readonly="true" style="display:none">
                  <input name="s_reside_place" type="hidden" id="s_reside_place2" value="">
                  <input name="s_team" type="hidden" id="s_reside2">
             			<a href="#" class="btnType03" onClick="pop_Window('ce_select.asp?gubun=<%="수정"%>&mg_group=<%=mg_group%>','ceselect','scrollbars=yes,width=500,height=400')">CE변경</a>
                </td>
							  <th>문자발송</th>
							  <td class="left">
                  <input type="radio" name="sms_yn" value="Y">발송
                	<input name="sms_yn" type="radio" value="N" checked>발송안함
                </td>
					    </tr>
							<tr>
							  <th>장애내용</th>
							  <td class="left" colspan="3">
							  	<textarea name="as_memo" cols="115" rows="5" class="style12"><%=rs("as_memo")%></textarea>
                </td>
                <th>협업여부</th>
								<td class="left">
                				<input type="radio" name="cowork_yn" value="N" <% if rs("cowork_yn") = "N" then %>checked<% end if %>>일반 
              					<input type="radio" name="cowork_yn" value="Y" <% if rs("cowork_yn") = "Y" then %>checked<% end if %>>협업 
                </td>
					    </tr>
							<tr>
							  <th>요청일자</th>
							  <td class="left" colspan="3">
              					<input name="request_date" type="text" id="datepicker" style="width:70px;" readonly="true">&nbsp;
                                <input name="request_hh" type="text" id="request_hh" value="<%=mid(rs("request_time"),1,2)%>" size="2" maxlength="2">시
                                <input name="request_mm" type="text" id="request_mm" value="<%=mid(rs("request_time"),3,2)%>" size="2" maxlength="2">분
	                          </td>
							  <th>처리유형</th>
							  <td class="left">
                                <select name="as_type" id="select2">
                                  <option value="원격처리" <%If Rs("as_type") = "원격처리" then %>selected<% end if %>>원격처리</option>
                                  <option value="방문처리" <%If Rs("as_type") = "방문처리" then %>selected<% end if %>>방문처리</option>
                                  <option value="신규설치" <%If Rs("as_type") = "신규설치" then %>selected<% end if %>>신규설치</option>
                                  <option value="신규설치공사" <%If Rs("as_type") = "신규설치공사" then %>selected<% end if %>>신규설치공사</option>
                                  <option value="이전설치" <%If Rs("as_type") = "이전설치" then %>selected<% end if %>>이전설치</option>
                                  <option value="이전설치공사" <%If Rs("as_type") = "이전설치공사" then %>selected<% end if %>>이전설치공사</option>
                                  <option value="랜공사" <%If Rs("as_type") = "랜공사" then %>selected<% end if %>>랜공사</option>
                                  <option value="이전랜공사" <%If Rs("as_type") = "이전랜공사" then %>selected<% end if %>>이전랜공사</option>
                                  <option value="장비회수" <%If Rs("as_type") = "장비회수" then %>selected<% end if %>>장비회수</option>
                                  <option value="예방점검" <%If Rs("as_type") = "예방점검" then %>selected<% end if %>>예방점검</option>
                                  <option value="기타" <%If Rs("as_type") = "기타" then %>selected<% end if %>>기타</option>
                                </select>
                              </td>
					      	</tr>
						</tbody>
					</table>
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="등록" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        <span class="btnType01"><input type="button" value="취소" onclick="javascript:goBefore();"></span>
                    </div>
                    <input name="acpt_no" type="hidden" id="acpt_no" value="<%=rs("acpt_no")%>">
                    <input name="acpt_date" type="hidden" id="acpt_date" value="<%=acpt_date%>">
                    <input name="acpt_hh" type="hidden" id="acpt_hh" value="<%=acpt_hh%>">
                    <input name="acpt_mm" type="hidden" id="acpt_mm2" value="<%=acpt_mm%>">
                    <input name="as_type_old" type="hidden" id="as_type_old2" value="<%=rs("as_type")%>">
                    <input name="sms_old" type="hidden" id="sms_old" value="<%=rs("sms")%>">
                    <input name="be_pg" type="hidden" id="be_pg" value="<%=be_pg%>">
                    <input name="page" type="hidden" id="page" value="<%=page%>">
                    <input name="from_date" type="hidden" id="from_date" value="<%=from_date%>">
                    <input name="to_date" type="hidden" id="to_date" value="<%=to_date%>">
                    <input name="date_sw" type="hidden" id="date_sw" value="<%=date_sw%>">
                    <input name="process_sw" type="hidden" id="process_sw" value="<%=process_sw%>">
                    <input name="field_check" type="hidden" id="field_check" value="<%=field_check%>">
                    <input name="field_view" type="hidden" id="field_view" value="<%=field_view%>">
                    <input name="condi_com" type="hidden" id="condi_com" value="<%=condi_com%>">
				</form>
				</div>
			</div>
	</body>
</html>

