<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'on Error resume next

'===================================================
'### Request & Params
'===================================================
Dim min_month, now_month, title_line, from_month, to_month

min_month = "201501"
now_month = CStr(Mid(Now(), 1, 4)) & CStr(Mid(Now(), 6, 2))
from_month = now_month - 1
to_month = now_month

title_line = "��� ���� �ϰ� ���"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN""http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>��� ���� �ý���</title>
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

			function frmcheck(){
				if(chkfrm()){
					document.frm.submit();
				}
			}

			function chkfrm(){
				var from_month = $('#from_month').val();
				var to_month = $('#to_month').val();

				if(document.frm.from_month.value ==""){
					alert('from����� �Է��ϼ���');
					frm.from_month.focus();
					return false;
				}

				if(document.frm.to_month.value ==""){
					alert('to����� �Է��ϼ���');
					frm.to_month.focus();
					return false;
				}

				if(document.frm.from_month.value < document.frm.min_month.value){
					alert('from ����� 201501���� ũ�ų� ���ƾ� �մϴ�.');
					frm.from_month.focus();
					return false;
				}

				if(document.frm.to_month.value >= document.frm.now_month.value){
					alert('to ����� ������ ���� �۾ƾ� �մϴ�.');
					frm.to_month.focus();
					return false;
				}

				if(!confirm(from_month + ' ���� ' + to_month + '���� ��븶���� ���� ����Ͻðڽ��ϱ�?')){
					return false;
				}else{
					//return true;

					var param = {"from_month":from_month, "to_month":to_month};

					let start_time = new Date();

					$.ajax({
						type : "GET"
						, dataType : 'html'
						, contentType: "application/x-www-form-urlencoded; charset=EUC-KR"
						, url: "/cost/cost_end_condi_cancel_ok.asp"
						, data: param
						, async: true
						, error: function(request, status, error){
							console.log("code = "+ request.status + " message = " + request.responseText + " error = " + error);
						}
						, success: function(data){
							let end_time = new Date();
							var elapedMin = (end_time.getTime() - start_time.getTime()) / 1000 / 60;

							console.log('����ð�(��) : ' + elapedMin);
							console.log(data);

							alert(data);
							location.href="/cost/cost_end_mg.asp";
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
		<div id="wrap">
			<!--#include virtual = "/include/cost_header.asp" -->
			<!--#include virtual = "/include/cost_report_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="/cost/cost_end_condi_cancel_ok.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>��ȸ����</legend>
					<dl>
						<dt>ó������</dt>
                        <dd>
                            <p>
                                <label>
                                    &nbsp;&nbsp;<strong>FROM���&nbsp;</strong> :
                                    <input name="from_month" id="from_month" type="text" value="<%=from_month%>" style="width:70px" maxlength="6">
                                    &nbsp;~&nbsp;
                                    &nbsp;&nbsp;<strong>TO���&nbsp;</strong> :
                                    <input name="to_month" id="to_month" type="text" value="<%=to_month%>" style="width:70px" maxlength="6">
                                </label>
                                    &nbsp;&nbsp;����� ��)201501
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser.jpg" alt="�˻�" /></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
                    <input type="hidden" name="min_month" value="<%=min_month%>" />
                    <input type="hidden" name="now_month" value="<%=now_month%>" />
				</form>
		</div>
	</div>
	</body>
</html>