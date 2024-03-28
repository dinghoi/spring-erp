<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
ce_id = request("ce_id")
team = request("team")

if ce_id = "" then
	ce_id = request.form("ce_id")
	team = request.form("team")
end if

mod_ce_id = request.form("mod_ce_id")

if mod_ce_id = "" then
	mod_ce_id = ce_id
end if

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_mumb = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

Sql_memb="select * from memb where user_id = '" + ce_id + "'"

Set rs_memb=DbConn.Execute(SQL_memb)

user_name = rs_memb("user_name")
rs_memb.close()

Sql="select * from ce_area where mg_ce_id = '" + ce_id + "' order by sido, gugun asc"
Rs.Open Sql, Dbconn, 1

title_line = "담당 CE 변경 (휴가/대체)"

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>미처리 현황</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">

			function goAction () {
		  		 window.close () ;
			}
			function frmcheck(){
				var url;
				url = document.frm.url.value;
				if(document.frm.ce_id.value == document.frm.mod_ce_id.value) {
					alert('변경 내역이 없습니다.');
					frm.mod_ce_id.focus();
					return false;}
							
				{
				a=confirm('정말 변경 하시겠습니까?')
				if (a==true) {
					location.replace(url);
					return true;
				}
				return false;
				}
			}
			function form_submit(){
			document.frm.submit();
			}

        </script>

	</head>
	<body>
		<div id="container">				
			<div class="gView">
			<h3 class="tit"><%=title_line%></h3>
				<form method="post" name="frm" action="ce_exchange.asp">
					<table cellpadding="0" cellspacing="0" summary="" class="tableView">
						<colgroup>
							<col width="15%" >
							<col width="25%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
							  <th>담당 CE</th>
							  <td class="left"><%=ce_id%>&nbsp;(<%=user_name%>)
                				<input name="ce_id" type="hidden" id="ce_id" value="<%=ce_id%>">
                				<input name="team" type="hidden" id="team" value="<%=team%>">
                			  </td>
							  <th>변경 CE</th>
							  <td class="left">
							  <%
								Sql_memb="select * from memb where team = '"+team+"' order by user_name asc"
								Rs_memb.Open Sql_memb, Dbconn, 1
							  %>
                				<select name="mod_ce_id" id="select4" onChange="form_submit()">
                  			  <% 
								do until rs_memb.eof 
			  				  %>
                  					<option value='<%=rs_memb("user_id")%>' <%If rs_memb("user_id") = mod_ce_id then %>selected<% end if %>><%=rs_memb("user_name")%></option>
                  			  <%
									rs_memb.movenext()  
								loop 
								rs_memb.Close()
							  %>
							 	</select>
</td>
					      	</tr>
						</tbody>
					</table>
					<br>
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="10%" >
							<col width="*" >
							<col width="15%" >
							<col width="20%" >
							<col width="15%" >
							<col width="20%" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">시도</th>
								<th scope="col">구군</th>
								<th scope="col" colspan="2">기존 CE</th>
								<th scope="col" colspan="2">변경 CE</th>
							</tr>
						</thead>
						<tbody>
							<%
                            i = 0
                            do until rs.eof 
								if rs("mg_ce_id") <> "" then
										Sql_memb="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
										Set Rs_memb=dbconn.execute(Sql_memb)
										if rs_memb.eof then
											user_name = "미등록"
											else
											user_name = rs_memb("user_name")					
										end if
									else
										user_name = "미등록"
								end if
                            %>
							<tr>
								<td class="first"><%=rs("sido")%><input name="sido" type="hidden" id="sido" value="<%=rs("sido")%>"></td>
								<td><%=rs("gugun")%><input name="gugun" type="hidden" id="gugun" value="<%=rs("gugun")%>"></td>
								<td><%=rs("mg_ce_id")%></td>
								<td><%=user_name%></td>
								<td><%=mod_ce_id%></td>
					  		<%
								Sql_memb="select * from memb where user_id = '" + mod_ce_id + "'"
								Set Rs_memb=dbconn.execute(Sql_memb)
								if rs_memb.eof then
									user_name = "미등록"
								  else
									user_name = rs_memb("user_name")					
								end if
                          	%>
								<td><%=user_name%></td>
							</tr>
							<%
                                rs.movenext()
                            loop
							url = "ce_exchange_ok.asp?ce_id="+ce_id+"&mod_ce_id="+mod_ce_id
                            %>
						</tbody>
					</table>                    
					<br>
                    <div align=center>
                        <span class="btnType01"><input type="button" value="변경" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                        <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                    </div>
        			<input type="hidden" name="url" value="<%=url%>" ID="Hidden1">
				</form>
				</div>
			</div>
	</body>
</html>

