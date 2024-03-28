<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim ary_ce_id(50)
dim ary_ce_name(50)

i_sw = request.form("i_sw")

sido = request("sido")
if sido = "" then
	sido=Request.form("sido")
end if
mod_ce_id = request.form("mod_ce_id")
gugun = request.form("gugun")
act_ce_id = mod_ce_id + ","

i=1
j= 1
jj=0
k=0

do until i=0
	i=0
	i=instr(j,act_ce_id,",")'
	
	if	i=0 then
		exit do
	end if
	jj=i-1
	if j=i then
		ary_ce_id(k)=""
	  else	  
		ary_ce_id(k)=trim(mid(act_ce_id,j,jj-j+1))
	end if
	j=i+1
	k=k+1

loop

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_type = Server.CreateObject("ADODB.Recordset")

Dbconn.open dbconnect

title_line = "CE 지역 배정"
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>A/S 관리 시스템</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript">
			function getPageCode(){
				return "4 1";
			}
		</script>
		<script type="text/javascript">
			function viewcheck () {
					document.frm.submit ();
			}			
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}			

			function chkfrm(){				
				a=confirm('정말 변경 하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/ce_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="ce_area_setting.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<strong>시도 : </strong>
                                <select name="sido" id="sido" style="width:100px">
                					<option>선택</option>
                            <%
								Sql="select * from etc_code where etc_type = '81' order by etc_code asc"
								Rs_etc.Open Sql, Dbconn, 1
								do until rs_etc.eof
							%>
                					<option value='<%=rs_etc("etc_name")%>' <% if rs_etc("etc_name") = sido then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                			<%
									rs_etc.movenext()  
								loop 
								rs_etc.Close()
							%>
            					</select>
                                <a href="#" onclick="javascript:viewcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
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
								<th class="first" scope="col">시도/구군</th>
								<th scope="col" colspan="2">기존 CE</th>
								<th scope="col" colspan="2">변경 CE</th>
							</tr>
						</thead>
						<tbody>
							<%
                            Sql="select * from ce_area where sido = '" + sido + "' and mg_group = '" + mg_group + "' order by sido asc"
                            Rs.Open Sql, Dbconn, 1
                            i = 0
                            do until rs.eof 
								if rs("mg_ce_id") <> "" then
										Sql_memb="select * from memb where user_id = '" + rs("mg_ce_id") + "'"
										Set Rs_memb=dbconn.execute(Sql_memb)
                                    if rs_memb.eof or rs_memb.bof then
                                        user_name = "ERROR"
                                        else
										if rs_memb("mg_group") <> mg_group then
											user_name = "ERROR"
										  else
                                        	user_name = rs_memb("user_name")					
										end if
                                    end if
									else
										user_name = "미등록"
								end if
                            %>
							<tr>
								<td class="first"><%=sido%> / <%=rs("gugun")%>
                          		<input name="gugun" type="hidden" id="gugun" value="<%=rs("gugun")%>"></td>
								<td><%=rs("mg_ce_id")%></td>
                      			<td><%=user_name%></td>
								<td><input name="mod_ce_id" type="text" id="mod_ce_id" value="<%=ary_ce_id(i)%>" size="15"></td>
							<%
                            if ary_ce_id(i) <> "" then
                                    Sql_memb="select * from memb where user_id = '" + ary_ce_id(i) + "'"
                                    Set Rs_memb=dbconn.execute(Sql_memb)
                                    if rs_memb.eof or rs_memb.bof then
                                        user_name = "ERROR"
                                        else
										if rs_memb("mg_group") <> mg_group then
											user_name = "ERROR"
										  else
                                        	user_name = rs_memb("user_name")					
										end if
                                    end if
                                else
                                    user_name = "미등록"				
                            end if
                            ary_ce_name(i) = user_name
                            %>
								<td><%=user_name%></td>
							</tr>
						<%
							i = i + 1
							rs.movenext()  
						loop
						i_sw = i
						rs.Close()
						%>
						</tbody>
					</table>
				</div>
				<%
                check_sw = "N"
                for j = 0 to i - 1
	                if	ary_ce_id(j) <> "" then
	                    check_sw = "Y"
                    end if
                next
                for j = 0 to i - 1
                    if	ary_ce_name(j) = "ERROR" then
                        check_sw = "N"
                    end if
                next
				%>                
				<br>
                <div align=center>
                    <a href="#" onclick="javascript:viewcheck();" class="btnType01">확인</a>
                <% if gugun <> "" and check_sw = "Y" then %>
                    <a href="ce_area_setting_ok.asp?sido=<%=sido%>&gugun=<%=gugun%>&mod_ce_id=<%=mod_ce_id%>" onclick="javascript:form_chk();" class="btnType01">변경</a>
                <% end if %>
                    <a href="ce_area_setting.asp?sido=<%=sido%>" onclick="" class="btnType01">취소</a>
                </div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

