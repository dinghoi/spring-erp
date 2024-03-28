<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim month_tab(24, 2)
Dim quarter_tab(8, 2)
Dim year_tab(3, 2)

Dim u_type, insu_id, view_condi, insu_yyyy, insu_class, insu_id_name
Dim curr_mm, i, cal_quarter, cal_month, view_month, j, cal_year
Dim from_amt, to_amt, st_amt, hap_rate, emp_rate, com_rate, insu_comment
Dim title_line, rsIns

u_type = Request("u_type")
insu_id = Request("insu_id")
view_condi = Request("view_condi")
insu_yyyy = Request("insu_yyyy")
insu_class = Request("insu_class")
insu_id_name = Request("view_condi")

' 최근3개년도 테이블로 생성
year_tab(3,1) = Mid(Now(), 1, 4)
year_tab(3,2) = CStr(year_tab(3, 1)) & "년"
year_tab(2,1) = CInt(Mid(Now(), 1, 4)) - 1
year_tab(2,2) = CStr(year_tab(2, 1)) & "년"
year_tab(1,1) = CInt(Mid(Now(), 1, 4)) - 2
year_tab(1,2) = CStr(year_tab(1, 1)) & "년"

' 분기 테이블 생성
curr_mm = Mid(Now(), 6, 2)

If curr_mm > 0 And curr_mm < 4 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) & "1"
End If

If curr_mm > 3 And curr_mm < 7 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) & "2"
End If

If curr_mm > 6 And curr_mm < 10 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) & "3"
End If

If curr_mm > 9 And curr_mm < 13 Then
	quarter_tab(8, 1) = CStr(Mid(Now(), 1, 4)) & "4"
End If

quarter_tab(8, 2) = CStr(Mid(quarter_tab(8, 1), 1, 4)) & "년 " & CStr(Mid(quarter_tab(8, 1), 5, 1)) & "/4분기"

For i = 7 To 1 Step -1
	cal_quarter = CInt(quarter_tab(i+1, 1)) - 1

	If CStr(Mid(cal_quarter, 5, 1)) = "0" Then
		quarter_tab(i, 1) = CStr(CInt(Mid(cal_quarter, 1, 4)) - 1) & "4"
	Else
		quarter_tab(i, 1) = cal_quarter
	End If
	quarter_tab(i, 2) = CStr(Mid(quarter_tab(i, 1), 1, 4)) & "년 " & CStr(Mid(quarter_tab(i, 1), 5, 1)) & "/4분기"
Next

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
month_tab(24, 1) = cal_month
view_month = Mid(cal_month, 1, 4) & "년 " & Mid(cal_month, 5, 2) & "월"
month_tab(24, 2) = view_month

For i = 1 To 23
	cal_month = CStr(Int(cal_month) - 1)

	If Mid(cal_month, 5) = "00" Then
		cal_year = CStr(Int(Mid(cal_month, 1, 4)) - 1)
		cal_month = cal_year & "12"
	End If

	view_month = Mid(cal_month, 1, 4) & "년 " & Mid(cal_month, 5, 2) & "월"
	j = 24 - i
	month_tab(j, 1) = cal_month
	month_tab(j, 2) = view_month
Next

from_amt = 0
to_amt = 0
st_amt = 0
hap_rate = 0
emp_rate = 0
com_rate = 0
insu_comment = ""

title_line = "4대보험요율 등록"

If u_type = "U" Then
	'sql = "select * from pay_insurance where insu_yyyy = '" + insu_yyyy + "' and insu_id = '" + insu_id + "' and insu_class = '" + insu_class + "'"
	objBuilder.Append "SELECT insu_yyyy, insu_id, insu_class, insu_id_name, from_amt, to_amt, "
	objBuilder.Append "	st_amt, hap_rate, emp_rate, com_rate, insu_comment "
	objBuilder.Append "FROM pay_insurance "
	objBuilder.Append "WHERE insu_yyyy = '"&insu_yyyy&"' "
	objBuilder.Append "	AND insu_id = '"&insu_id&"' "
	objBuilder.Append "	AND insu_class = '"&insu_class&"' "

	Set rsIns = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

    insu_yyyy = rsIns("insu_yyyy")
    insu_id = rsIns("insu_id")
	insu_class = rsIns("insu_class")
    insu_id_name = rsIns("insu_id_name")
	from_amt =rsIns("from_amt")
    to_amt = rsIns("to_amt")
    st_amt = rsIns("st_amt")
    hap_rate = rsIns("hap_rate")
    emp_rate = rsIns("emp_rate")
    com_rate = rsIns("com_rate")
    insu_comment = rsIns("insu_comment")

	rsIns.Close() : Set rsIns = Nothing

	title_line = "4대보험요율 변경"
End If

DBConn.Close() : Set DBConn = Nothing
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
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
			function goAction(){
			   window.close();
			}

			function goBefore(){
			   history.back();
			}

			function frmcheck(){
				if(formcheck(document.frm) && chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				if(document.frm.insu_yyyy.value =="" && document.frm.insu_yyyy.value ==""){
					alert('적용 년도를 입력하세요');
					frm.insu_yyyy.focus();
					return false;
				}

				if(document.frm.insu_class.value ==""){
					alert('등급을 입력하세요');
					frm.insu_class.focus();
					return false;
				}

				if(document.frm.emp_rate.value ==""){
					alert('근로자 요율을 입력하세요');
					frm.emp_rate.focus();
					return false;
				}

				if(document.frm.com_rate.value ==""){
					alert('사용자 요율을 입력하세요');
					frm.com_rate.focus();
					return false;
				}

				if(document.frm.to_amt.value ==""){
					alert('표준보수월액을 입력하세요');
					frm.to_amt.focus();
					return false;
				}

				{
					a=confirm('입력하시겠습니까?')
					if (a==true) {
						return true;
					}
					return false;
				}
			}

			function num_chk(txtObj){
				f_amt = parseInt(document.frm.from_amt.value.replace(/,/g,""));
				f_amt = String(f_amt);
				num_len = f_amt.length;
				sil_len = num_len;
				f_amt = String(f_amt);

				if(f_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if(sil_len > 3) f_amt = f_amt.substr(0,num_len -3) + "," + f_amt.substr(num_len -3,3);
				if(sil_len > 6) f_amt = f_amt.substr(0,num_len -6) + "," + f_amt.substr(num_len -6,3) + "," + f_amt.substr(num_len -2,3);

				document.frm.from_amt.value = f_amt;

				t_amt = parseInt(document.frm.to_amt.value.replace(/,/g,""));
				t_amt = String(t_amt);
				num_len = t_amt.length;
				sil_len = num_len;
				t_amt = String(t_amt);

				if(t_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if(sil_len > 3) t_amt = t_amt.substr(0,num_len -3) + "," + t_amt.substr(num_len -3,3);
				if(sil_len > 6) t_amt = t_amt.substr(0,num_len -6) + "," + t_amt.substr(num_len -6,3) + "," + t_amt.substr(num_len -2,3);

				document.frm.to_amt.value = t_amt;

				s_amt = parseInt(document.frm.st_amt.value.replace(/,/g,""));
				s_amt = String(s_amt);
				num_len = s_amt.length;
				sil_len = num_len;
				s_amt = String(s_amt);

				if(s_amt.substr(0,1) == "-") sil_len = num_len - 1;
				if(sil_len > 3) s_amt = s_amt.substr(0,num_len -3) + "," + s_amt.substr(num_len -3,3);
				if(sil_len > 6) s_amt = s_amt.substr(0,num_len -6) + "," + s_amt.substr(num_len -6,3) + "," + s_amt.substr(num_len -2,3);

				document.frm.st_amt.value = s_amt;

                e_rate = parseFloat((document.frm.emp_rate.value),3);
				c_rate = parseFloat((document.frm.com_rate.value),3);
				h_rate = e_rate + c_rate;
				document.frm.hap_rate.value = h_rate;
			}

			function update_view(){
				var c = document.frm.u_type.value;

				if(c == 'U'){
					document.getElementById('cancel_col').style.display = '';
					document.getElementById('info_col').style.display = '';
				}
			}
        </script>
	</head>
	<body onload="update_view();">
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="/pay/insa_pay_insurance_save.asp" method="post" name="frm">
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
								<th class="first">보험구분</th>
								<td class="left">
									<input name="insu_id_name" type="text" value="<%=insu_id_name%>" style="width:120px" readonly="true"></td>
								<th>적용년도</th>
								<td class="left">
									<select name="insu_yyyy" id="insu_yyyy" type="text" value="<%=insu_yyyy%>" style="width:90px">
									<%For i=3 To 1 Step -1	%>
										<option value="<%=year_tab(i,1)%>" <%If insu_yyyy = year_tab(i, 1) Then %>selected<%End If %>><%=year_tab(i, 2)%></option>
									<%Next	%>
									</select>
                                </td>
                                <th>등급<br>01부터~</th>
								<td class="left">
                                <input name="insu_class" type="text" value="<%=insu_class%>" style="width:30px" onKeyUp="checklength(this, 3)"></td>
							</tr>
                           	<tr>
								<th class="first">보수월액<br>이상-미만</th>
								<td colspan="3" class="left">
                                <input name="from_amt" type="text" value="<%=FormatNumber(from_amt, 0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                -
                                <input name="to_amt" type="text" value="<%=FormatNumber(to_amt, 0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th>표준<br>보수월액</th>
								<td class="left">
                                <input name="st_amt" type="text" value="<%=FormatNumber(st_amt, 0)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
							</tr>
							<tr>
								<th colspan="6" class="first" style="background:#F5FFFA">요율(사업장)</th>
							</tr>
							<tr>
								<th class="first">근로자</th>
								<td class="left">
                                <input name="emp_rate" type="text" value="<%=FormatNumber(emp_rate, 3)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th class="first">사용자</th>
								<td class="left">
                                <input name="com_rate" type="text" value="<%=FormatNumber(com_rate, 3)%>" style="width:90px;text-align:right" onKeyUp="num_chk(this);">
                                </td>
                                <th class="first">합계</th>
								<td class="left">
                                <input name="hap_rate" type="text" value="<%=FormatNumber(hap_rate, 3)%>" style="width:90px;text-align:right" readonly="true">
                                </td>
							</tr>
                        	<tr>
								<th class="first">비고</th>
								<td colspan="5" class="left">
                                <input name="insu_comment" type="text" value="<%=insu_comment%>" style="width:570px" onKeyUp="checklength(this, 50)">
                                </td>
							</tr>
                      </tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" /></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();" /></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" />
                <input type="hidden" name="insu_id" value="<%=insu_id%>" />
			</form>
		</div>
	</body>
</html>