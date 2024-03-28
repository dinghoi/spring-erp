<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'On Error Resume Next

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

Dim curr_date, be_date, be_month, inc_yyyy, inc_yyyy1
Dim title_line, cal_month, view_month, i, j, cal_year, pre_month

title_line = " 조직 및 인사 월 마감처리 "
curr_date = Now()
be_date = DateAdd("m", -1, curr_date)
be_month = CStr(Mid(be_date, 1, 4)) & CStr(Mid(be_date, 6, 2))
inc_yyyy = CStr(Mid(be_date, 1, 4)) & " 년 " & CStr(Mid(be_date, 6, 2))
inc_yyyy1 = CStr(Mid(be_date, 1, 4)) & "" & CStr(Mid(be_date, 6, 2))

'전월 날짜
pre_month = CStr(Mid(DateAdd("m", -2, curr_date), 1, 4)) & CStr(Mid(DateAdd("m", -2, curr_date), 6, 2))

' 년월 테이블생성
'cal_month = cstr(mid(dateadd("m",-1,now()),1,4)) + cstr(mid(dateadd("m",-1,now()),6,2))
cal_month = Mid(CStr(Now()), 1, 4) & Mid(CStr(Now()), 6, 2)
month_tab(24, 1) = cal_month
view_month = Mid(cal_month,1, 4) & "년 " & Mid(cal_month, 5, 2) & "월"
month_tab(24, 2) = view_month

For i = 1 To 23
	cal_month = CStr(Int(cal_month) - 1)

	If Mid(cal_month,5) = "00" Then
		cal_year = cstr(int(Mid(cal_month,1, 4)) - 1)
		cal_month = cal_year + "12"
	End If

	view_month = Mid(cal_month, 1, 4) & "년 " & Mid(cal_month, 5, 2) & "월"
	j = 24 - i
	month_tab(j, 1) = cal_month
	month_tab(j, 2) = view_month
Next
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
	<title>인사관리 시스템</title>
	<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
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

		function frmcheck(){
			if(formcheck(document.frm)){
				document.frm.submit ();
			}
		}

		function insa_month_final(val){
			/*if(!confirm("조직 및 인사 월 마감처리를 하시겠습니까 ?")) return;
			var frm = document.frm;

			document.frm.action = "/insa/insa_month_final_submit_ok.asp";
			document.frm.submit();*/

			if(!confirm("조직 및 인사 월 마감처리를 하시겠습니까 ?")){
				return false;
			}else{
				var inc_yyyy1 = $('#inc_yyyy1').val();
				var pre_month = $('#pre_month').val();

				var param = {"inc_yyyy1":inc_yyyy1, "pre_month":pre_month};

				let start_time = new Date();

				$.ajax({
					type : "GET"
					, dataType : 'html'
					, contentType: "application/x-www-form-urlencoded; charset=EUC-KR"
					, url: "/insa/insa_month_final_submit_ok.asp"
					, data: param
					, async: true
					, error: function(request, status, error){
						console.log("code = "+ request.status + " message = " + request.responseText + " error = " + error);
					}
					, success: function(data){
						let end_time = new Date();
						var elapedMin = (end_time.getTime() - start_time.getTime()) / 1000 / 60;

						console.log('진행시간(분) : ' + elapedMin);
						console.log(data);

						alert(data);
						//location.href="/insa/insa_month_final_submit.asp";
						window.close();
						return;
					}
					, beforeSend: function(){
						var width = 0;
						var height = 0;
						var left = 0;
						var top = 0;

						width = 220;
						height = 118;
						top = ( $(window).height() - height ) / 2 + $(window).scrollTop();
						left = ( $(window).width() - width ) / 2 + $(window).scrollLeft();

						if($("#div_ajax_load_image").length != 0){
							$("#div_ajax_load_image").css({
								"top": top+"px",
								"left": left+"px"
							});
							$("#div_ajax_load_image").show();
						}else{
							$('body').append('<div id="div_ajax_load_image" style="position:absolute; top:' + top + 'px; left:' + left + 'px; width:' + width + 'px; height:' + height + 'px; z-index:9999; background:#f0f0f0; filter:alpha(opacity=50); opacity:alpha*0.5; margin:auto; padding:0; "><img src="/image/wait.gif" style="width:220px; height:118px;"></div>');
						}
					}
					, complete: function(){
						$("#div_ajax_load_image").hide();
					}
				});
			}
		}
	</script>
</head>
<body>
	<div id="container">
		<h3 class="insa"><%=title_line%></h3>
		<!--<form action="/insa/insa_month_final_submit.asp" method="post" name="frm">-->
		<fieldset class="srch">
			<legend>조회영역</legend>
			<dl>
				<dd>
					<p>
					<strong>
						<select name="inc_yyyy1" id="inc_yyyy1" type="text" value="<%=inc_yyyy1%>" style="width:110px">
						<%For i = 24 To 1 Step - 1	%>
							<option value="<%=month_tab(i,1)%>" <%If inc_yyyy1 = month_tab(i,1) then %>selected<% end if %>><%=month_tab(i,2)%></option>
						<%Next	%>
						 </select>
						월 조직 및 인사 마감
					</strong>
					<label>
						<input name="emp_no" type="hidden" id="emp_no" value="<%=emp_no%>" style="width:40px" readonly="true">
					</label>
					</p>
				</dd>
			</dl>
		</fieldset>
		<h3 class="stit"> 마감처리를 클릭하시면 <%=inc_yyyy%>  월 조직 및 인사자료가 마감됩니다.<br>&nbsp;<br></h3>
		<div class="gView">
			<table cellpadding="0" cellspacing="0" class="tableWrite">
				<colgroup>
					<col width="30%" >
					<col width="*" >
				</colgroup>
				<tbody>
					<tr>
						<th colspan="2" class="left" style=" border-bottom:1px solid #ffffff;">
						마감처리는 매월 5일경 전월 조직 및 인사자료를 마감합니다.<br><br>
						<strong>※ 전월에 대한 조직 및 인사발령등을 5일전에 모두 마감 하시고 처리를 하셔야 합니다.</strong><br><br>
						<strong>※ 당월 조직 및 인사발령을 하시기 전에 전월 마감처리를 하셔야 합니다.</strong>
						</th>
					</tr>
				</tbody>
			</table>
		</div>
		<br>
		<div align="center">
			<span class="btnType01"><input type="button" value="마감처리" onclick="insa_month_final('be_month');return false;" /></span>
			<span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
		</div>
		<!--<input type="hidden" name="be_month" value="<%'=be_month%>" />-->
		<!--<input type="hidden" name="" value="<%'=be_month1%>" />-->
		<input type="hidden" name="pre_month" id="pre_month" value="<%=pre_month%>" />
		<!--</form>-->
	</div>
</body>
</html>