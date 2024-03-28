<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim acpt_gname
Dim acpt_standard
Dim rs
Dim rs_numRows

code_ary = request("code_ary")
goods_type = Request("goods_type")

ck_sw=Request("ck_sw")

'If ck_sw = "y" Then
If goods_type <> "" Then
	code_ary = request("code_ary")
    goods_type = Request("goods_type") 
  else
	code_ary = Request.form("code_ary")
    goods_type = Request.form("goods_type") 
End if

'response.write(code_ary)


acpt_standard = Request.form("acpt_standard")
acpt_gname = Request.form("acpt_gname")
view_c = Request.form("view_c")

if view_c = "" then
	    acpt_standard = ""
	    acpt_gname = ""
	    view_c = "acpt"
end if

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

'if goods_type = "" then
'	sql = "select * from met_goods_code where goods_type = '" + goods_type + "'"
'  else
'	sql = "select * from met_goods_code where goods_type like '%" + goods_type + "%' ORDER BY goods_name,goods_code ASC"
'end if


if view_c = "acpt" then
	sql = "select * from met_goods_code where  goods_used_sw = 'Y' and goods_name like '%"&acpt_gname&"%' ORDER BY goods_gubun,goods_name,goods_code ASC"
  else
	sql = "select * from met_goods_code where  goods_used_sw = 'Y' and goods_gubun like '%"&acpt_standard&"%' ORDER BY goods_gubun,goods_name,goods_code ASC"
end if

Rs.Open Sql, Dbconn, 1

title_line = "품목 검색"

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>품목 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			function frmcheck1 () {
//				if (chkfrm1()) {
				document.frm1.submit ();
//				}
			}			
			
			function chkfrm() {
				if(document.frm.goods_type.value == "" || document.frm.goods_type.value == " ") {
					alert('용도구분을 입력하세요');
					frm.goods_type.focus();
					return false;}
				{
					return true;
				}
			}
			
			function condi_view() {

				if (eval("document.frm.view_c[0].checked")) {
					document.getElementById('p_standard').style.display = 'none';
					document.getElementById('p_gname').style.display = '';
				}	
				if (eval("document.frm.view_c[1].checked")) {
					document.getElementById('p_standard').style.display = '';
					document.getElementById('p_gname').style.display = 'none';
				}	
			}
		</script>

	</head>
	<body onLoad="condi_view()">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="met_import_goods_select.asp?code_ary=<%=code_ary%>&ck_sw=<%="n"%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
								<label>
                                <strong>용도구분</strong>
                                <input name="goods_type" type="text" value="<%=goods_type%>" style="width:90px" id="goods_type" readonly="true">

								</label>
                                <label>
                              	<input type="radio" name="view_c" value="acpt" <% if view_c = "acpt" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목명
                                <input type="radio" name="view_c" value="work" <% if view_c = "work" then %>checked<% end if %> style="width:25px" onClick="condi_view()">
                                품목구분
								</label>
                                <label id="p_gname">
								<strong>품목명</strong>
                                <input name="acpt_gname" type="text" id="acpt_gname" value="<%=acpt_gname%>" style="width:130px; text-align:left; ime-mode:active">
								</label>
                              <%
								Sql="select * from met_etc_code where etc_type = '04' order by etc_code asc"
					            Rs_etc.Open Sql, Dbconn, 1
							  %>                                       
								<label id="p_standard">
								<strong>품목구분</strong>
                                <select name="acpt_standard" id="acpt_standard" type="text" style="width:130px">
                                    <option value="">선택</option>
                			  <% 
								do until Rs_etc.eof 
			  				  %>
                					<option value='<%=rs_etc("etc_name")%>' <%If acpt_standard = rs_etc("etc_name") then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			  <%
									Rs_etc.movenext()  
								loop 
								Rs_etc.Close()
							  %>
            					</select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<form action="met_import_goods_select_ok.asp" method="post" name="frm1">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="6%" >
                            <col width="8%" >
                            <col width="14%" >
							<col width="*" >
							<col width="20%" >
							<col width="20%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">선택</th>
                                <th scope="col">상태</th>
                                <th scope="col">품목코드</th>
                                <th scope="col">품목명</th>
								<th scope="col">품목구분</th>
								<th scope="col">Part_Number</th>
							</tr>
						</thead>
						<tbody>
						<%
							i = 0
							do until rs.eof or rs.bof
								i = i + 1
						%>
							<tr>
								<td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="<%=rs("goods_code")%>"></td>
								<td><%=rs("goods_grade")%>&nbsp;</td>
                                <td><%=rs("goods_code")%>&nbsp;</td>
								<td><%=rs("goods_name")%>&nbsp;</td>
                                <td><%=rs("goods_gubun")%>&nbsp;</td>
								<td><%=rs("part_number")%>&nbsp;</td>
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
								<td class="first" colspan="6">내역이 없습니다</td>
							</tr>
                        <%
						end if
						%>
							<tr>
								<td class="first; left" colspan="6"><span class="btnType04"><input type="button" value="선택" onclick="javascript:frmcheck1();"></span></td>
							</tr>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="code_ary" value="<%=code_ary%>">
                <input type="hidden" name="goods_type1" value="<%=goods_type%>">
				</form>
		</div>        				
	</body>
</html>

