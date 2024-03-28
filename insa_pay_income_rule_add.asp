<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim month_tab(24,2)
dim quarter_tab(8,2)
dim year_tab(3,2)

u_type = request("u_type")
rule_id = request("rule_id")
view_condi = request("view_condi")
rule_yyyy = request("rule_yyyy")
rule_cl = request("rule_cl")

' 최근3개년도 테이블로 생성
year_tab(3,1) = mid(now(),1,4)
year_tab(3,2) = cstr(year_tab(3,1)) + "년"
year_tab(2,1) = cint(mid(now(),1,4)) - 1
year_tab(2,2) = cstr(year_tab(2,1)) + "년"
year_tab(1,1) = cint(mid(now(),1,4)) - 2
year_tab(1,2) = cstr(year_tab(1,1)) + "년"

' 분기 테이블 생성
curr_mm = mid(now(),6,2)
if curr_mm > 0 and curr_mm < 4 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "1"
end if
if curr_mm > 3 and curr_mm < 7 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "2"
end if
if curr_mm > 6 and curr_mm < 10 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "3"
end if
if curr_mm > 9 and curr_mm < 13 then
	quarter_tab(8,1) = cstr(mid(now(),1,4)) + "4"
end if

quarter_tab(8,2) = cstr(mid(quarter_tab(8,1),1,4)) + "년 " + cstr(mid(quarter_tab(8,1),5,1)) + "/4분기"

for i = 7 to 1 step -1
	cal_quarter = cint(quarter_tab(i+1,1)) - 1
	if cstr(mid(cal_quarter,5,1)) = "0" then
		quarter_tab(i,1) = cstr(cint(mid(cal_quarter,1,4))-1) + "4"
	  else
		quarter_tab(i,1) = cal_quarter
	end if	 
	quarter_tab(i,2) = cstr(mid(quarter_tab(i,1),1,4)) + "년 " + cstr(mid(quarter_tab(i,1),5,1)) + "/4분기"
next

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))	
cal_month = mid(cstr(now()),1,4) + mid(cstr(now()),6,2)
month_tab(24,1) = cal_month
view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
month_tab(24,2) = view_month
for i = 1 to 23
	cal_month = cstr(int(cal_month) - 1)
	if mid(cal_month,5) = "00" then
		cal_year = cstr(int(mid(cal_month,1,4)) - 1)
		cal_month = cal_year + "12"
	end if	 
	view_month = mid(cal_month,1,4) + "년 " + mid(cal_month,5,2) + "월"
	j = 24 - i
	month_tab(j,1) = cal_month
	month_tab(j,2) = view_month
next

rule_id_name = request("view_condi")
rule_year_pay = 0
rule_st_deduct = 0
rule_exceed = 0
rule_exceed_rate = 0
rule_add = 0
rule_add_rate = 0
rule_comment = ""

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "근로소득세율 등록"

if u_type = "U" then

	sql = "select * from pay_income_rule where rule_yyyy = '" + rule_yyyy + "' and rule_id = '" + rule_id + "' and rule_cl = '" + rule_cl + "'"
	set rs = dbconn.execute(sql)

    rule_yyyy = rs("rule_yyyy")
    rule_id = rs("rule_id")
	rule_cl = rs("rule_cl")
    rule_id_name = rs("rule_id_name")
	rule_year_pay =rs("rule_year_pay")
    rule_st_deduct = rs("rule_st_deduct")
    rule_exceed = rs("rule_exceed")
    rule_exceed_rate = rs("rule_exceed_rate")
    rule_add = rs("rule_add")
    rule_add_rate = rs("rule_add_rate")
    rule_comment = rs("rule_comment")
	rs.close()

	title_line = "근로소득세율 변경"
end if
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사급여 시스템</title>
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
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
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
				if(document.frm.rule_yyyy.value =="" && document.frm.rule_yyyy.value =="") {
					alert('적용년도를 입력하세요');
					frm.rule_yyyy.focus();
					return false;}
				if(document.frm.rule_cl.value =="") {
					alert('분류를 입력하세요');
					frm.rule_cl.focus();
					return false;}
				if(document.frm.rule_year_pay.value =="") {
					alert('총급여액을 입력하세요');
					frm.rule_year_pay.focus();
					return false;}			
							
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			
			function num_chk(txtObj){
				year_pay = parseInt(document.frm.rule_year_pay.value.replace(/,/g,""));		
				year_pay = String(year_pay);
				num_len = year_pay.length;
				sil_len = num_len;
				year_pay = String(year_pay);
				if (year_pay.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) year_pay = year_pay.substr(0,num_len -3) + "," + year_pay.substr(num_len -3,3);
				if (sil_len > 6) year_pay = year_pay.substr(0,num_len -6) + "," + year_pay.substr(num_len -6,3) + "," + year_pay.substr(num_len -2,3);
				document.frm.rule_year_pay.value = year_pay; 
				
				st_deduct = parseInt(document.frm.rule_st_deduct.value.replace(/,/g,""));		
				st_deduct = String(st_deduct);
				num_len = st_deduct.length;
				sil_len = num_len;
				st_deduct = String(st_deduct);
				if (st_deduct.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) st_deduct = st_deduct.substr(0,num_len -3) + "," + st_deduct.substr(num_len -3,3);
				if (sil_len > 6) st_deduct = st_deduct.substr(0,num_len -6) + "," + st_deduct.substr(num_len -6,3) + "," + st_deduct.substr(num_len -2,3);
				document.frm.rule_st_deduct.value = st_deduct; 
				
				r_exceed = parseInt(document.frm.rule_exceed.value.replace(/,/g,""));		
				r_exceed = String(r_exceed);
				num_len = r_exceed.length;
				sil_len = num_len;
				r_exceed = String(r_exceed);
				if (r_exceed.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) r_exceed = r_exceed.substr(0,num_len -3) + "," + r_exceed.substr(num_len -3,3);
				if (sil_len > 6) r_exceed = r_exceed.substr(0,num_len -6) + "," + r_exceed.substr(num_len -6,3) + "," + r_exceed.substr(num_len -2,3);
				document.frm.rule_exceed.value = r_exceed; 		

                r_add = parseInt(document.frm.rule_add.value.replace(/,/g,""));		
				r_add = String(r_add);
				num_len = r_add.length;
				sil_len = num_len;
				r_add = String(r_add);
				if (r_add.substr(0,1) == "-") sil_len = num_len - 1;
				if (sil_len > 3) r_add = r_add.substr(0,num_len -3) + "," + r_add.substr(num_len -3,3);
				if (sil_len > 6) r_add = r_add.substr(0,num_len -6) + "," + r_add.substr(num_len -6,3) + "," + r_add.substr(num_len -2,3);
				document.frm.rule_add.value = r_add; 		
			
			}						
			
			function update_view() {
			var c = document.frm.u_type.value;
				if (c == 'U') 
				{
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
        </script>
	</head>
	<body onload="update_view()">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_pay_income_rule_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="10%" >
							<col width="23%" >
							<col width="10%" >
							<col width="23%" >
							<col width="10%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">세율구분</th>
								<td class="left">
                                <input name="rule_id_name" type="text" value="<%=rule_id_name%>" style="width:120px" readonly="true"></td>
								<th>적용년월</th>
								<td class="left">
                                <select name="rule_yyyy" id="rule_yyyy" type="text" value="<%=rule_yyyy%>" style="width:90px">
                                    <%	for i = 3 to 1 step -1	%>
                                    <option value="<%=year_tab(i,1)%>" <%If rule_yyyy = year_tab(i,1) then %>selected<% end if %>><%=year_tab(i,2)%></option>
                                    <%	next	%>
                                 </select>
                                </td>
                                <th>분류<br>01부터~</th>
								<td class="left">
                                <input name="rule_cl" type="text" value="<%=rule_cl%>" style="width:30px" onKeyUp="checklength(this,2)">
                                 -특별공제: 공제대상수
                                </td>
							</tr>
                           	<tr>
								<th class="first">총급여액</th>
								<td class="left">
                                <input name="rule_year_pay" type="text" value="<%=formatnumber(rule_year_pay,0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
								<td colspan="4" class="left">* 과세표준액/산출세액</td>
							</tr>             
							<tr>
								<th colspan="6" class="first" style="background:#F5FFFA">공&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;제</th>
							</tr>
							<tr>
								<th class="first">기준공제액</th>
								<td class="left">
                                <input name="rule_st_deduct" type="text" value="<%=formatnumber(rule_st_deduct,0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th class="first">초과공제액</th>
								<td class="left">
                                <input name="rule_exceed" type="text" value="<%=formatnumber(rule_exceed,0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th class="first">초과공제율</th>
								<td class="left">
                                <input name="rule_exceed_rate" type="text" value="<%=formatnumber(rule_exceed_rate,2)%>" style="width:90px;text-align:right">
                                </td>
							</tr>
                            <tr>
								<th colspan="6" class="first" style="background:#F5FFFA">추가&nbsp;&nbsp;&nbsp;공제</th>
							</tr>
                            <tr>
								<th class="first">추가공제액</th>
								<td colspan="3" class="left">
                                <input name="rule_add" type="text" value="<%=formatnumber(rule_add,0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th class="first">추가공제율</th>
								<td class="left">
                                <input name="rule_add_rate" type="text" value="<%=formatnumber(rule_add_rate,2)%>" style="width:90px;text-align:right">
                                </td>
							</tr>
                        	<tr>
								<th class="first">비고</th>
								<td colspan="5" class="left">
                                <textarea name="rule_comment" rows="2" id="textarea"><%=rule_comment%></textarea>
                                </td>
							</tr>                            
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
                <input type="hidden" name="rule_id" value="<%=rule_id%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

