<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
car_no = Request.form("car_no")

Set Dbconn = Server.CreateObject("ADODB.connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if car_no = "" then
	SQL = "select * from car_info where car_owner = '회사' and car_no = '" + car_no + "' ORDER BY car_no ASC"
 else
	SQL = "select * from car_info where car_owner = '회사' and car_no like '%" + car_no + "%' ORDER BY car_no ASC"
end if
Rs.open SQL, Dbconn, 1

title_line = "차량 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>차량 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function car_list(car_owner,car_no,car_name,oil_kind,last_km)
			{
				opener.document.frm.car_owner.value = car_owner;
				opener.document.frm.car_no.value = car_no;
				opener.document.frm.car_name.value = car_name;
				opener.document.frm.oil_kind.value = oil_kind;
				opener.document.frm.last_km.value = last_km;
				opener.document.frm.start_km.value = last_km;
				opener.document.frm.end_km.value = last_km;
				window.close();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.car_no.value =="") {
					alert('차량번호를 입력하세요');
					frm.car_no.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="car_search.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>차량번호를 입력하세요 </strong>
								<label>
        						<input name="car_no" type="text" id="car_no" value="<%=car_no%>" style="width:150px;text-align:left">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="*" >
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">차량번호</th>
								<th scope="col">소유</th>
								<th scope="col">차종</th>
								<th scope="col">유종</th>
								<th scope="col">최종KM</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof or rs.bof
							car_owner = rs("car_owner")
							car_no = rs("car_no")
							sql = "select car_no, max(end_km) as max_km from transit_cost where car_no = '"&car_no&"'"
							set rs_tran=dbconn.execute(sql)							
							max_km = rs_tran("max_km")
							rs_tran.close()
							car_name = rs("car_name")
							oil_kind = rs("oil_kind")
							if max_km = "" or isnull(max_km) then
								last_km = rs("last_km")
							  else
								last_km = max_km
							end if
						%>
							<tr>
								<td class="first">
                                <a href="#" onClick="car_list('<%=car_owner%>','<%=car_no%>','<%=car_name%>','<%=oil_kind%>','<%=last_km%>');"><%=rs("car_no")%></a>
                                </td>
								<td><%=car_owner%></td>
								<td><%=car_name%></td>
								<td><%=oil_kind%></td>
								<td><%=last_km%></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()
						loop
						rs.close()
						if i = 0 then
						%>
							<tr>
								<td class="first" colspan="5">내역이 없습니다</td>
							</tr>
                        <%
						end if
						%>
						</tbody>
					</table>
				</div>
			</form>
		</div>        				
	</body>
</html>

