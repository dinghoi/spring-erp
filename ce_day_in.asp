<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%

dim in_cnt_tab(31)
dim in_tot_tab(31)
dim in_date_tab(31)

to_date=Request.form("to_date")
team = request.form("team")

If to_date = "" Then
	curr_dd = cstr(datepart("d",now))
	to_date = mid(cstr(now()),1,10)
	team = "전체"
End If
from_date = mid(cstr(dateadd("d",-30,to_date)),1,10)

in_cnt_tab(0) = 0
in_tot_tab(0) = 0
for i = 0 to 30
	in_date_tab(i+1) = mid(cstr(dateadd("d",i,from_date)),1,10)
	in_cnt_tab(i+1) = 0
	in_tot_tab(i+1) = 0
next

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set Rs_in = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

if  team = "전체" then
	sql = "select memb.user_id,memb.team,memb.user_name,memb.reside from as_acpt inner join memb on as_acpt.mg_ce_id = memb.user_id "
'	sql = sql + " Where (as_acpt.mg_group='"+mg_group+"')"
	sql = sql + " where (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"')"
	sql = sql + " GROUP BY memb.user_id,memb.team,memb.user_name,memb.reside Order By memb.team, memb.user_name Asc"
 else
	sql = "select memb.user_id,memb.team,memb.user_name,memb.reside from as_acpt inner join memb on as_acpt.mg_ce_id = memb.user_id "
	sql = sql + " Where (memb.team='"+team+"')"
	sql = sql + " and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"')"
	sql = sql + " GROUP BY memb.user_id,memb.team,memb.user_name,memb.reside Order By memb.user_name Asc"
end if
Rs.Open Sql, Dbconn, 1

title_line = "CE별 일자별 입고현황"
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
			function getPageCode(){
				return "4 1";
			}
		</script>
		<script type="text/javascript">
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=from_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=to_date%>" );
			});	  
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if (document.frm.to_date.value =="") {
					alert ("종료일을 지정해야 합니다");
					return false;
				}	
				return true;
			}
		</script>

	</head>
	<body>
		<div id="wrap">			
			<!--#include virtual = "/include/header.asp" -->
			<!--#include virtual = "/include/ce_menu.asp" -->
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="waiting.asp?pg_name=ce_day_in.asp" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
						<dt>조건 검색</dt>
                        <dd>
                            <p>
								<label>
								<strong>시작일 : </strong>
                                	<input name="from_date" type="text" value="<%=from_date%>" style="width:70px" id="from_date">
								</label>
								<label>
								<strong>종료일 : </strong>
                                	<input name="to_date" type="text" value="<%=to_date%>" style="width:70px" id="datepicker1">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser.jpg" alt="검색"></a>
                    			<a href="ce_day_in_excel.asp?from_date=<%=from_date%>&to_date=<%=to_date%>" class="btnType04">엑셀다운로드</a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="100" >
							<col width="90" >
							<col width="50" >
							<col width="30" >
						<%  for i = 1 to 31	%>	
							<col width="30" >
                        <%  next	%>
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col" rowspan="2">소속</th>
								<th scope="col" rowspan="2">CE명</th>
								<th scope="col" rowspan="2">상주</th>
								<th scope="col" colspan="32" style=" border-bottom:1px solid #e3e3e3;">일 자 별</th>
							</tr>
							<tr>
								<th scope="col" style=" border-left:1px solid #e3e3e3;">소계</th>
						<%  for i = 1 to 31	%>	
								<th scope="col"><%=right(in_date_tab(i),2)%></th>
                        <%  next	%>
							</tr>
						</thead>
						<tbody>
						<% 
                        do until rs.eof 

                            sql = "select count(*) as in_cnt, in_date from as_acpt where (mg_ce_id='"+rs("user_id")+"') "
                            sql = sql + " and (in_date >= '" + from_date + "' AND in_date <= '"+to_date+"') GROUP BY in_date Order By in_date Asc"
                            Rs_in.Open Sql, Dbconn, 1
                            do until rs_in.eof
                                in_cnt = clng(rs_in("in_cnt"))
                                for j = 1 to 31
                                    if cstr(rs_in("in_date")) = cstr(in_date_tab(j)) then
                                        in_cnt_tab(j) = in_cnt_tab(j) + in_cnt
                                        in_cnt_tab(0) = in_cnt_tab(0) + in_cnt				
                                        in_tot_tab(j) = in_tot_tab(j) + in_cnt
                                        in_tot_tab(0) = in_tot_tab(0) + in_cnt				
                                        exit for
                                    end if 
                                next
                                rs_in.movenext()
                            loop
                            rs_in.close()
        
                            if rs("reside") = "0" then
                                reside = "."
                              else
                                reside = "상주"
                            end if
                        %>
							<tr>
                              <td><%=rs("team")%></td>
                              <td><%=rs("user_name")%></td>
                              <td><%=reside%></td>
                              <td bgcolor="#FFFFCA" class="right"><%=formatnumber(in_cnt_tab(0),0)%></td>
						<% for j = 1 to 31 %>
                              <td class="right"><%=formatnumber(in_cnt_tab(j),0)%></td>
                        <%	next %>
							</tr>
						<%
                            for i = 0 to 31
                                in_cnt_tab(i) = 0
                            next
                            rs.movenext()
                        loop
                        rs.close()
                        %>
							<tr>
                              <th colspan="3">총계</th>
                              <th><%=formatnumber(in_tot_tab(0),0)%></th>
						<% for j = 1 to 31 %>
                              <th><%=formatnumber(in_tot_tab(j),0)%></th>
                        <%	next %>
							</tr>
 						</tbody>
					</table>
				</div>
			</form>
		</div>				
	</div>        				
	</body>
</html>

