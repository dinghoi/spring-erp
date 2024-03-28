<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim dept_name
Dim rs
Dim rs_numRows

dept_name = ""
If (Request.Form("dept_name")  <> "") Then 
  dept_name = Request.Form("dept_name") 
End If
company = request("company")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs_memb = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if dept_name = "" then
	sql = "select * from asset_dept where company = 'null' and dept_name  = '" + dept_name + "'"
  else
	Sql = "select * , concat(ifnull(org_first,' '),' ',ifnull(org_second,' '),' ',ifnull(dept_name,' ')) as org_name from asset_dept where company = '" + company + "' and concat(ifnull(org_second,' '),' ',ifnull(dept_name,' ')) like '%" + dept_name + "%' ORDER BY dept_name ASC"
end if

rs.open SQL, DbConn, 1

title_line = "조직코드 조회"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>조직코드 조회</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function deptcode(org_first,org_second,dept_name,dept_code,sido,gugun,dong,addr,person,tel_ddd,tel_no1,tel_no2,mg_ce_id,mg_ce,team,reside_place,internet_no)
			{
				opener.document.frm.org_first.value = org_first;
				opener.document.frm.org_second.value = org_second;
				opener.document.frm.dept_name.value = dept_name;
				opener.document.frm.dept_code.value = dept_code;
				opener.document.frm.old_sido.value = sido;
				opener.document.frm.old_gugun.value = gugun;
				opener.document.frm.old_dong.value = dong;
				opener.document.frm.old_addr.value = addr;
				opener.document.frm.acpt_user.value = person;
				opener.document.frm.tel_ddd.value = tel_ddd;
				opener.document.frm.tel_no1.value = tel_no1;
				opener.document.frm.tel_no2.value = tel_no2;
				opener.document.frm.old_mg_ce_id.value = mg_ce_id;
				opener.document.frm.old_mg_ce.value = mg_ce;
				opener.document.frm.old_team.value = team;
				opener.document.frm.old_reside_place.value = reside_place;
				opener.document.frm.internet_no.value = internet_no;
				window.close();
				opener.document.frm.user_grade.focus();
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.dept_name.value =="") {
					alert('지점명을 입력하세요');
					frm.dept_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="dept_search_nh.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>지사 또는 지점명을 입력하세요</strong>
								<label>
        						<input name="dept_name" type="text" id="dept_name" value="<%=dept_name%>" style="text-align:left; width:150px">
								<input name="company" type="hidden" id="company" value="<%=company%>">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="45%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">코드</th>
								<th scope="col">조직명</th>
								<th scope="col">주소</th>
							</tr>
						</thead>
						<tbody>
						<%
						i = 0
						do until rs.eof
							i = i + 1
				
							sql_area="select * from ce_area where sido = '"+ rs("sido")+"' and gugun = '"+rs("gugun")+"' and mg_group = '"+mg_group+"'"
							Set rs_area=dbconn.execute(sql_area)
							if rs_area.eof or rs_area.bof then
								mg_ce_id = "미등록"
							  else
								mg_ce_id = rs_area("mg_ce_id")
							end if
							
							sql_memb="select * from memb where user_id = '"&mg_ce_id&"'"
							Set rs_memb=dbconn.execute(sql_memb)
							
							if rs_memb.eof then
								mg_ce = ""
								team = "ERROR"
								reside_place = "ERROR"
								user_name = ""
							  else
								mg_ce = rs_memb("user_name")
								team = rs_memb("team")
								reside_place = rs_memb("reside_place")
								user_name = rs_memb("user_name")
							end if
							%>
							<tr>
								<td class="first"><%=rs("dept_code")%></td>
								<td>
                                <a href="#" onClick="deptcode('<%=rs("org_first")%>','<%=rs("org_second")%>','<%=rs("dept_name")%>','<%=rs("dept_code")%>','<%=rs("sido")%>','<%=rs("gugun")%>','<%=rs("dong")%>','<%=rs("addr")%>','<%=rs("person")%>','<%=rs("tel_ddd")%>','<%=rs("tel_no1")%>','<%=rs("tel_no2")%>','<%=mg_ce_id%>','<%=mg_ce%>','<%=team%>','<%=reside_place%>','<%=rs("internet_no")%>');"><%=rs("org_name")%></a>
                                </td>
								<td><%=rs("sido")%>&nbsp;<%=rs("gugun")%>&nbsp;<%=rs("dong")%></td>
							</tr>
						<%
							rs.movenext()
						loop
						rs.close()
						if  i = 0 and dept_name <> "" then
							msg = "내역이 없습니다 !!!"
						  else
							msg = ""
						end if
						%>
							<tr>
								<td class="first" colspan="3"><%=msg%></td>
							</tr>
						</tbody>
					</table>
				</div>
				</form>
		</div>        				
	</body>
</html>

