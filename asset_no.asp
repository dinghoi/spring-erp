<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
company = request("company")
gubun = request("gubun")
code_seq = request("code_seq")
asset_name = request("asset_name")
asset_name = Replace(asset_name,"'","&quot;")
asset_name = Replace(asset_name,"""","&quot;")
asset_code = company + "-" + gubun + "-" + code_seq

buy_date = mid(now(),1,10)
Set Dbconn=Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

curr_date = mid(cstr(now()),1,10)
curr_year = mid(curr_date,1,4)
curr_month = mid(curr_date,6,2)

sql="select max(asset_no) as max_no from asset"
sql="select max(asset_no) as max_no from asset where mid(asset_no,1,2) = '" + company + "'"
set rs=dbconn.execute(sql)
if isnull(rs("max_no")) then
	asset_no = company + "-" + curr_year + curr_month + "-" + "0000"
  else  	
	asset_no = mid(rs("max_no"),1,2) + "-" + mid(rs("max_no"),3,6) + "-" + mid(rs("max_no"),9,4)
end if

title_line = "자산번호 부여"
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
												$( "#datepicker" ).datepicker("setDate", "<%=buy_date%>" );
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

				if(document.frm.buy_date.value != document.frm.curr_date.value) {
					alert('부여일자와 현재일자가 다른가요?');
					}
				if(document.frm.asset_cnt.value == 0) {
					alert('부여갯수를 입력하세요');
					frm.asset_cnt.focus();
					return false;}
				if(document.frm.asset_cnt.value > 99) {
					alert('부여갯수가 100개 이상입니까?');
					}
			
				{
				a=confirm('자산번호를 부여 하겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
			function onlynum() {
			 if((event.keyCode<48)||(event.keyCode>57)) 
				{ event.returnValue=false; } 
			}//-->
        </script>
	</head>
	<body>
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="asset_no_ok.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="30%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">자산코드</th>
								<td class="left">
                                <%=asset_code%>
                                  <input name="company" type="hidden" id="company" value="<%=company%>">
                                  <input name="gubun" type="hidden" id="gubun" value="<%=gubun%>">
                                  <input name="code_seq" type="hidden" id="code_seq" value="<%=code_seq%>">
              					</td>
							</tr>
							<tr>
								<th class="first">자산명</th>
								<td class="left"><%=asset_name%><input name="asset_name" type="hidden" id="asset_name" value="<%=asset_name%>"></td>
							</tr>
							<tr>
								<th class="first">LAST번호</th>
								<td class="left"><%=asset_no%><input name="asset_no" type="hidden" id="asset_no" value="<%=rs("max_no")%>"></td>
							</tr>
							<tr>
								<th class="first">매입일자</th>
								<td class="left">
                                  <input name="buy_date" type="text" value="<%=buy_date%>" style="width:70px" id="datepicker">
                                  <input name="curr_date" type="hidden" id="curr_date" value="<%=curr_date%>">
                                </td>
							</tr>
							<tr>
								<th class="first">구입대수</th>
								<td class="left"><input name="asset_cnt" type="text" id="asset_cnt" onKeyPress="onlynum();" value="0"  style="width:70px" maxlength="4"></td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="부여" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

