<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/include/db_create.asp" -->
<%
code_ary = request("code_ary")
slip_id = request("slip_id")
srv_type = Request.Form("srv_type")
Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from etc_code where etc_type = '51' and type_name = '" + srv_type + "'"
if srv_type = "상품" or srv_type = "토너" then
	sql = "select goods_code as etc_code, goods_type as type_name, goods_gubun as etc_name, concat(goods_name,' ',goods_standard) as group_name from met_goods_code where goods_type = '"&srv_type&"' order by goods_gubun"
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
				if (chkfrm()) {
					document.frm.submit ();
				}
			}
			function frmcheck1 () {
//				if (chkfrm1()) {
				document.frm1.submit ();
//				}
			}			
			
			function chkfrm() {
				if(document.frm.srv_type.value == "" || document.frm.srv_type.value == " ") {
					alert('서비스유형을 입력하세요');
					frm.srv_type.focus();
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
				<form action="sales_goods_select.asp?code_ary=<%=code_ary%>&slip_id=<%=slip_id%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>서비스유형을 선택하세요 </strong>
								<label>
							<%
                                Sql="select * from type_code where etc_type = '51' order by type_name asc"
                                rs_type.Open Sql, Dbconn, 1
                            %>
                                <select name="srv_type" id="srv_type" style="width:150px">
                                    <option value=''>선택</option> 
                            <% 
                                do until rs_type.eof 
                            %>
                                    <option value='<%=rs_type("type_name")%>' <%If rs_type("type_name") = srv_type  then %>selected<% end if %>><%=rs_type("type_name")%></option>
                            <%
                                    rs_type.movenext()  
                                loop 
                                rs_type.Close()
                            %>
                                </select>
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				</form>
				<form action="sales_goods_select_ok.asp" method="post" name="frm1">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="8%" >
							<col width="10%" >
							<col width="23%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">선택</th>
								<th scope="col">유형</th>
								<th scope="col">품목명</th>
								<th scope="col">규격</th>
							</tr>
						</thead>
						<tbody>
						<%
							i = 0
							do until rs.eof or rs.bof
								i = i + 1
							%>
							<tr>
								<td class="first"><input type="checkbox" name="sel_check" id="sel_check" value="<%=rs("etc_code")%>"></td>
								<td><%=rs("type_name")%></td>
								<td><%=rs("etc_name")%></td>
								<td class="left"><%=rs("group_name")%>&nbsp;</td>
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
							<tr>
								<td class="first; left" colspan="4"><span class="btnType04"><input type="button" value="선택" onclick="javascript:frmcheck1();"></span></td>
							</tr>
						</tbody>
					</table>
				</div>
				<input type="hidden" name="code_ary" value="<%=code_ary%>">
				<input type="hidden" name="slip_id" value="<%=slip_id%>">
				</form>
		</div>        				
	</body>
</html>

