<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_dong
Dim rs
Dim rs_numRows

gubun = request("gubun")
in_dong = Request.Form("in_dong") 
in_dong = replace(in_dong," ","")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if in_dong = "" then
	sql = "select * from area_mg where dong = '" + in_dong + "'"
  else
	Sql = "select * from area_mg where dong like '%" + in_dong + "%' ORDER BY dong ASC"
end if

rs.open SQL, DbConn, 1

title_line = "동코드 검색"

' https://www.juso.go.kr/info/RoadNameDataList.do 도로명 정보 조회 참조

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>동코드 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function areacode(gubun,sido,gugun,dong,mg_ce_id,mg_ce,team,reside_place,reside_company,zipcode)
			{
				if(gubun =="1")
					{ 
					opener.document.frm.sido.value = sido;
					opener.document.frm.gugun.value = gugun;
					opener.document.frm.dong.value = dong;
					opener.document.frm.zip_code.value = zipcode;
					window.close();
					opener.document.frm.addr.focus();
					}
				else
					{ 
					opener.document.frm.sido.value = sido;
					opener.document.frm.gugun.value = gugun;
					opener.document.frm.dong.value = dong;
					opener.document.frm.mg_ce_id.value = mg_ce_id;
					opener.document.frm.mg_ce.value = mg_ce;
					opener.document.frm.team.value = team;
					opener.document.frm.reside_place.value = reside_place;
	//				opener.document.frm.reside_company.value = reside_company;
					window.close();
					opener.document.frm.addr.focus();
//					opener.document.frm.as_memo.focus();
					}
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.in_dong.value == "" || document.frm.in_dong.value == " ") {
					alert('동명을 입력하세요');
					frm.in_dong.focus();
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
				https://www.juso.go.kr/info/RoadNameDataList.do 도로명 정보 조회 참조
				<form action="area_search.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>동명을 입력하세요 </strong>
								<label>
        						<input name="in_dong" type="text" id="in_dong" value="<%=in_dong%>" style="text-align:left; width:150px; ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="25%" >
							<col width="25%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">시도</th>
								<th scope="col">구군</th>
								<th scope="col">동</th>
								<th scope="col">
							<% if gubun = "1" then	%>
								우편번호
                            <%   else	%>
                                담당자
                            <% end if	%>
                                </th>
							</tr>
						</thead>
						<tbody>
						<%
							i = 0
							do until rs.eof or rs.bof
								i = i + 1
								sql_area="select * from ce_area where sido = '"+ rs("sido")+"' and gugun = '"+rs("gugun")+"' and mg_group = '"+mg_group+"'"
								Set rs_area=dbconn.execute(sql_area)
								if rs_area.eof or rs_area.bof then
									mg_ce_id = "미등록"
								  else
									if rs_area("mg_ce_id") = "" or isnull(rs_area("mg_ce_id")) then
										mg_ce_id = "미등록"
									  else								
										mg_ce_id = rs_area("mg_ce_id")
									end if
								end if
																
								sql_memb="select * from memb where user_id = '"&mg_ce_id&"'"
								Set rs_memb=dbconn.execute(sql_memb)
								
								if rs_memb.eof then
									mg_ce = ""
									team = "ERROR"
									reside_place = "ERROR"
									reside_company = "ERROR"
									user_name = ""
								  else
									mg_ce = rs_memb("user_name")
									team = rs_memb("team")
									reside_place = rs_memb("reside_place")
									reside_company = rs_memb("reside_company")
									user_name = rs_memb("user_name")
								end if

							%>
							<tr>
								<td class="first"><%=rs("sido")%></td>
								<td><%=rs("gugun")%></td>
								<td>
                                <a href="#" onClick="areacode('<%=gubun%>','<%=rs("sido")%>','<%=rs("gugun")%>','<%=rs("dong")%>','<%=mg_ce_id%>','<%=mg_ce%>','<%=team%>','<%=reside_place%>','<%=reside_company%>','<%=rs("zipcode")%>');"><%=rs("dong")%></a>
                                </td>
								<td>
							<% if gubun = "1" then	%>
								<%=rs("zipcode")%>
                            <%   else	%>
								<%=mg_ce_id%> / <%=user_name%>
                            <% end if	%>
                                </td>
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
								<td class="first" colspan="4">내역이 없습니다</td>
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

