<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
slip_month = request("slip_month")
slip_gubun = request("slip_gubun")
work_date = mid(slip_month,1,4) + "-" + mid(slip_month,5,2) + "-01"
work_date = datevalue(work_date)
from_date = dateadd("m",-1,work_date)
end_date = dateadd("m",1,from_date)
to_date = dateadd("d",-1,end_date)

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from general_cost where (slip_date >='"&from_date&"' and slip_date <='"&to_date&"') and (slip_gubun ='"&slip_gubun&"') and (forward_yn = 'N') and (emp_no = '"&user_id&"') ORDER BY slip_date ASC"
Rs.Open Sql, Dbconn, 1

title_line = "전월 " + slip_gubun + " 내역 (담당자 : " + user_name + ")"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>전월 외주비 내역</title>
		<link rel="stylesheet" href="http://code.jquery.com/ui/1.10.2/themes/smoothness/jquery-ui.css" />
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function frmcheck () {
					document.frm.submit ();
			}
			
//			function chkfrm() {
//				if(document.frm.srv_type.value == "" || document.frm.srv_type.value == " ") {
//					alert('서비스유형을 입력하세요');
//					frm.srv_type.focus();
//					return false;}
//				{
//					return true;
//				}
//			}

		</script>
		<script>
		$(document).ready(function(){
        	$("#all_check").on("click", function(){
            	var isChecked = $("#all_check").prop("checked");
            	$("input:checkbox[name=sel_check]").prop("checked",isChecked);
            
            	//레이블 텍스트 변경
//            	if(isChecked) {
//                	$(this).next().html("전체해제");
//            	}else{
//                	$(this).next().html("전체선택");
//            	}
        	})
    	});     
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="outside_cost_batch_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="5%" >
							<col width="10%" >
							<col width="20%" >
							<col width="10%" >
							<col width="10%" >
							<col width="9%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col"><input type="checkbox" name="all_check" id="all_check"></th>
								<th scope="col">발행일자</th>
								<th scope="col">외주업체</th>
								<th scope="col">합계</th>
								<th scope="col">공급가액</th>
								<th scope="col">부가세</th>
								<th scope="col">발행내역</th>
							</tr>
						</thead>
						<tbody>
						<%
							i = 0
							do until rs.eof or rs.bof
								i = i + 1
							%>
							<tr>
								<td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="<%=cstr(rs("slip_date"))+cstr(rs("slip_seq"))%>"></td>
								<td><%=rs("slip_date")%></td>
								<td><%=rs("customer")%></td>
								<td class="right"><%=formatnumber(rs("price"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost"),0)%></td>
								<td class="right"><%=formatnumber(rs("cost_vat"),0)%></td>
								<td class="left"><%=rs("slip_memo")%></td>
							</tr>
							<%
								rs.movenext()
							loop
							rs.close()
							%>
						<%
						  if i = 0 then
						%>
							<tr>
								<td class="first" colspan="7">내역이 없습니다</td>
							</tr>
                        <%
						end if
						%>
							<tr>
								<td class="first; left" colspan="7"><span class="btnType04"><input type="button" value="선택완료" onclick="javascript:frmcheck();"></span></td>
							</tr>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="slip_gubun" value="<%=slip_gubun%>" ID="Hidden1">
				</form>
		</div>        				
	</body>
</html>

