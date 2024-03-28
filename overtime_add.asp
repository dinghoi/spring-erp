<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
u_type = request("u_type")

mg_ce_id = user_id
mg_ce = user_name
apct_no = 0
company = ""
dept = ""
work_item = ""
from_hh = ""
from_mm = ""
to_hh = ""
to_mm = ""
work_gubun = ""
overtime_amt = 0
work_memo = ""
cancel_yn = "N"
end_yn = "N"
reg_id = user_id
reg_date = now()

curr_date = mid(cstr(now()),1,10)
work_date = mid(cstr(now()),1,10)

company = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "야특근 등록"
if u_type = "U" then

	work_date = request("work_date")
	mg_ce_id = request("mg_ce_id")

	sql = "select * from overtime where work_date = '" + work_date + "' and mg_ce_id = '" + mg_ce_id + "'"
	set rs = dbconn.execute(sql)

	sql="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
	set rs_memb=dbconn.execute(sql)

	if	rs_memb.eof or rs_memb.bof then
		mg_ce = "ERROR"
	  else
		mg_ce = rs_memb("user_name")
	end if
	rs_memb.close()						

	if isnull(rs("acpt_no")) then
		acpt_no = 0
	  else
		acpt_no = rs("acpt_no")
	end if
	mg_ce_id = rs("mg_ce_id")
	company = rs("company")
	dept = rs("dept")
	work_item = rs("work_item")
	from_hh = mid(rs("from_time"),1,2)
	from_mm = mid(rs("from_time"),3,2)
	to_hh = mid(rs("to_time"),1,2)
	to_mm = mid(rs("to_time"),3,2)
	work_gubun = rs("work_gubun") + "-" + cstr(rs("overtime_amt"))
	overtime_amt = int(rs("overtime_amt"))
	work_memo = rs("work_memo")
	sign_no = rs("sign_no")
	you_yn = rs("you_yn")
	cancel_yn = rs("cancel_yn")
	reg_id = rs("reg_id")
	reg_user = rs("reg_user")
	reg_date = rs("reg_date")
	mod_id = rs("mod_id")
	mod_user = rs("mod_user")
	mod_date = rs("mod_date")
	rs.close()

	sql = "select count(*) from overtime where acpt_no = "&int(acpt_no)
	Set RsCount = Dbconn.Execute (sql)
	tot_record = cint(RsCount(0)) 'Result.RecordCount

	title_line = "야특근 변경"
end if
if end_yn = "Y" then
	end_view = "마감"
  else
  	end_view = "진행"
end if

strNowWeek = WeekDay(work_date)
Select Case (strNowWeek)
   Case 1
       week = "일요일"
   Case 2
       week = "월요일"
   Case 3
       week = "화요일"
   Case 4
       week = "수요일"
   Case 5
       week = "목요일"
   Case 6
       week = "금요일"
   Case 7
       week = "토요일"
End Select

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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=work_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}			
			function chkfrm() {
				if(document.frm.work_gubun.value =="") {
					alert('작업구분을 선택하세요');
					frm.work_gubun.focus();
					return false;}
				if(document.frm.from_hh.value >"23"||document.frm.from_hh.value <"00") {
					alert('시작 시간이 잘못되었습니다');
					frm.from_hh.focus();
					return false;}
				if(document.frm.from_mm.value >"59"||document.frm.from_mm.value <"00") {
					alert('시작 분이 잘못되었습니다');
					frm.from_mm.focus();
					return false;}
				if(document.frm.to_hh.value >"23"||document.frm.to_hh.value <"00") {
					alert('종료 시간이 잘못되었습니다');
					frm.to_hh.focus();
					return false;}
				if(document.frm.to_mm.value >"59"||document.frm.to_mm.value <"00") {
					alert('종료 분이 잘못되었습니다');
					frm.to_mm.focus();
					return false;}
			
				if(document.frm.to_hh.value < document.frm.from_hh.value) {
					alert('종료시간이 시작시간 보다 빠름니다');
					frm.to_hh.focus();
					return false;}
			
				if(document.frm.from_hh.value == document.frm.to_hh.value) {
					if(document.frm.to_mm.value <= document.frm.from_mm.value) {
						alert('종료시간이 시작시간 보다 빠름니다');
						frm.to_mm.focus();
						return false;}}
				
			
				if(document.frm.work_memo.value =="" ) {
					alert('작업내용을 입력하세요');
					frm.work_memo.focus();
					return false;}
			
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function week_check() {
			
			a = document.frm.work_date.value.substring(0,4);
			b = document.frm.work_date.value.substring(5,7);
			c = document.frm.work_date.value.substring(8,10);
			
			var newDate = new Date(a,b-1,c); 
			var s = newDate.getDay(); 
			
			switch(s) {
				case 0: str = "일요일" ; break;
				case 1: str = "월요일" ; break;
				case 2: str = "화요일" ; break;
				case 3: str = "수요일" ; break;
				case 4: str = "목요일" ; break;
				case 5: str = "금요일" ; break;
				case 6: str = "토요일" ; break;
				}
			
				document.frm.week.value = str;			
			}
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
			function delcheck() 
				{
				acpt_no = document.frm.acpt_no.value;
				tot_record = document.frm.tot_record.value;
				del_msg = "접수번호 " + acpt_no + "번 전체 " + tot_record + "건이 삭제가 됩니다. 정말 삭제하시겠습니까?";
				a=confirm(del_msg)
				if (a==true) {
					document.frm.action = "overtime_del_ok.asp";
					document.frm.submit();
				return true;
				}
				return false;
				}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="overtime_add_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">작업일</th>
								<td class="left">
                                <%=work_date%>&nbsp;<%=week%>
                                <input name="work_date" type="hidden" value="<%=work_date%>">
                                </td>
								<th>작업자</th>
								<td class="left"><%=mg_ce%> (<%=mg_ce_id%>)
                                <input name="curr_date" type="hidden" id="curr_date" value="<%=now()%>">
                                <input name="mg_ce_id" type="hidden" id="mg_ce_id" value="<%=mg_ce_id%>">
                                </td>
							</tr>
							<tr>
								<th class="first">회사명</th>
								<td class="left"><%=company%></td>
								<th>부서명</th>
								<td class="left"><%=dept%></td>
							</tr>
							<tr>
								<th class="first">작업항목</th>
								<td class="left">
								<%=work_item%><input name="work_item" type="hidden" id="work_item" value="<%=work_item%>">
            					</td>
								<th>작업시간</th>
								<td class="left">
                                <input name="from_hh" type="text" id="from_hh" size="2" maxlength="2" value="<%=from_hh%>">시
								<input name="from_mm" type="text" id="from_mm" size="2" maxlength="2" value="<%=from_mm%>">분 ~
								<input name="to_hh" type="text" id="to_hh" size="2" maxlength="2" value="<%=to_hh%>">시
								<input name="to_mm" type="text" id="to_mm" size="2" maxlength="2" value="<%=to_mm%>">분
                                </td>
							</tr>
							<tr>
								<th class="first">야근항목</th>
								<td class="left"><select name="work_gubun" id="work_gubun" style="width:150px">
								  <option value="">선택</option>
								  <%
                                  Sql="select * from etc_code where etc_type = '41' and used_sw = 'Y' order by etc_code asc"
                                  rs_etc.Open Sql, Dbconn, 1
                                  do until rs_etc.eof
                                  		work_gubun_amt = rs_etc("etc_name") + "-" + cstr(rs_etc("etc_amt"))
                                  %>
								  <option value='<%=work_gubun_amt%>' <%If work_gubun_amt = work_gubun then %>selected<% end if %>><%=work_gubun_amt%></option>
								  <%
                                        rs_etc.movenext()
                                  loop
                                  rs_etc.close()						
                                  %>
						      </select></td>
								<th>작업내용</th>
								<td class="left"><input name="work_memo" type="text" id="work_memo" onKeyUp="checklength(this,50)"  style="width:200px" value="<%=work_memo%>"></td>
							</tr>
							<tr>
							  <th class="first">전자결재NO</th>
							  <td class="left"><input name="sign_no" type="text" id="sign_no" style="width:40px" onKeyUp="checkNum(this);" value="<%=sign_no%>" maxlength="4">
							    &nbsp;*숫자4자리만 입력 가능
							    <input type="hidden" name="reg_sw" value="<%=reg_sw%>" ID="reg_sw"></td>
							  <th><span class="first">유무상구분</span></th>
							  <td class="left"><input type="radio" name="you_yn" value="N" <% if you_yn = "N" then %>checked<% end if %> style="width:40px" id="Radio4">
							    무상
							    <input type="radio" name="you_yn" value="Y" <% if you_yn = "Y" then %>checked<% end if %> style="width:40px" id="Radio3">
							    유상 </td>
						  </tr>
							<tr style="display:none">
							  <th class="first">취소여부</th>
							  <td class="left"><input type="radio" name="cancel_yn" value="Y" <% if cancel_yn = "Y" then %>checked<% end if %> style="width:40px" id="Radio1">
							    취소
                                  <input type="radio" name="cancel_yn" value="N" <% if cancel_yn = "N" then %>checked<% end if %> style="width:40px" id="Radio2">
                              지급 </td>
							  <th>마감여부</th>
							  <td class="left"><%=end_view%>
						      <input name="end_yn" type="hidden" id="end_yn" value="<%=end_yn%>"></td>
			              </tr>
							<tr id="info_col" style="display:none">
								<th class="first">등록정보</th>
								<td class="left"><%=reg_user%>&nbsp;<%=reg_id%>(<%=reg_date%>)</td>
                                <th>변경정보</th>
								<td class="left"><%=mod_user%>&nbsp;<%=mod_id%>(<%=mod_date%>)</td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
				<%	
					if u_type = "U" and (user_id = mg_ce_id or user_id = reg_id) then
						if end_yn = "N" or end_yn = "C" then	
				%>
                    <span class="btnType01"><input type="button" value="삭제" onclick="javascript:delcheck();" ID="Button1" NAME="Button1"></span>
        		<%	
						end if
					end if	
				%>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
				<input type="hidden" name="acpt_no" value="<%=acpt_no%>" ID="Hidden1">
				<input type="hidden" name="tot_record" value="<%=tot_record%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

