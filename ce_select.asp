<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim ce_name
gubun = request("gubun")
ce_name = Request.Form("ce_name")
ce_name = replace(ce_name," ","")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if ce_name = "" then
	sql = "select * from memb where grade < '5' and user_name = '" + ce_name + "'"
  else
	sql = "select * from memb where grade < '5' and user_name like '%" + ce_name + "%' ORDER BY user_name ASC"
end if
Rs.Open Sql, Dbconn, 1

title_line = "CE 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>CE 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function cecode(mg_ce,mg_ce_id,team,reside_place,reside_company,gubun)
			{
				if(gubun =="입력")
					{ 
					opener.document.frm.s_ce.value = mg_ce;
					opener.document.frm.s_ce_id.value = mg_ce_id;
					window.close();
					opener.document.frm.as_memo.focus();
					}
				else
					{ 
					opener.document.frm.s_ce.value = mg_ce;
					opener.document.frm.s_ce_id.value = mg_ce_id;
					window.close();
					opener.document.frm.as_memo.focus();
					}
			}
			function frmcheck () {
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.ce_name.value == "" || document.frm.ce_name.value == " ") {
					alert('CE명을 입력하세요');
					frm.ce_name.focus();
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
				<form action="ce_select.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>CE명을 입력하세요 </strong>
								<label>
        						<input name="ce_name" type="text" id="ce_name" value="<%=ce_name%>" style="width:150px; text-align:left; ime-mode:active">
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
								<th class="first" scope="col">CE명</th>
								<th scope="col">아이디</th>
								<th scope="col">핸드폰 번호</th>
								<th scope="col">소속 / 상주처</th>
							</tr>
						</thead>
						<tbody>
						<%
							i = 0
							do until rs.eof or rs.bof
								i = i + 1
								sql_type="select * from type_code where etc_type='91' and etc_seq ='"+mg_group+"'"
								set rs_type=dbconn.execute(sql_type)
								if rs_type.eof then
									mg_group_name = "ERROR"
								  else  	
									mg_group_name = rs_type("type_name")
								end if
								rs_type.Close()		
							%>
							<tr>
								<td class="first"><a href="#" onClick="cecode('<%=rs("user_name")%>','<%=rs("user_id")%>','<%=rs("team")%>','<%=rs("reside_place")%>','<%=rs("reside_company")%>','<%=gubun%>');"><%=rs("user_name")%></a>
                                </td>
								<td><%=rs("user_id")%></td>
								<td><%=rs("hp")%></td>
								<td><%=rs("team")%> / <%=rs("reside_place")%></td>
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

