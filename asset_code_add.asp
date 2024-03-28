<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim company_tab(50,2)

if asset_company <> "00" then
	company = asset_company
end if

u_type = request("u_type")

gubun = ""
code_seq = ""
maker = ""
asset_name = ""
cpu = ""
mem = ""
hdd = ""
os = ""
spec = ""
rental = ""
unit_price = 0

Set Dbconn=Server.CreateObject("ADODB.Connection")
Set Rs = Server.CreateObject("ADODB.Recordset")
Set rs_etc = Server.CreateObject("ADODB.Recordset")
Dbconn.open dbconnect

title_line = "자산코드 등록"
if u_type = "U" then

	company = request("company")
	gubun = request("gubun")
	code_seq = request("code_seq")

	etc_code = "75" + company
	Sql="select * from etc_code where etc_code = '" + etc_code + "'"
	Set rs_etc=DbConn.Execute(SQL)
	if rs_etc.eof or rs_etc.bof then 
		company_name = "없음"
	  else 
		company_name = rs_etc("etc_name")
	end if
	rs_etc.close()						

	sql = "select * from asset_code where company = '" + company + "' and gubun = '" + gubun + "' and code_seq = '" + code_seq + "'"
	set rs = dbconn.execute(sql)
	
	gubun = rs("gubun")
	maker = rs("maker")
	asset_name = rs("asset_name")
	cpu = rs("cpu")
	mem = rs("mem")
	hdd = rs("hdd")
	os = rs("os")
	spec = rs("spec")
	rental = rs("rental")
	unit_price = rs("unit_price")

	rs.close()

	title_line = "자산코드 변경"
end if

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
			$(function() {    $( "#datepicker" ).datepicker();
												$( "#datepicker" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker" ).datepicker("setDate", "<%=bill_date%>" );
			});	  
			$(function() {    $( "#datepicker1" ).datepicker();
												$( "#datepicker1" ).datepicker("option", "dateFormat", "yy-mm-dd" );
												$( "#datepicker1" ).datepicker("setDate", "<%=inout_date%>" );
			});	  
			function goAction () {
			   window.close () ;
			}
			function goBefore () {
			   history.back() ;
			}
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}			
			function specview() {
			var c = document.frm.gubun.options[document.frm.gubun.selectedIndex].value;
				if (c == '01' || c == '03') 
				{
					document.getElementById('spec_menu').style.display = '';
				}
				if (c == '02' || c == '04' || c == '') 
				{
					document.getElementById('spec_menu').style.display = 'none';
				}
			}
			function chkfrm() {
				if(document.frm.gubun.value =="") {
					alert('자산구분을 선택하세요 !!!');
					frm.gubun.focus();
					return false;}
				if(document.frm.maker.value =="") {
					alert('제조사를 선택하세요 !!!');
					frm.maker.focus();
					return false;}
				if(document.frm.asset_name.value =="") {
					alert('자산명을 선택하세요 !!!');
					frm.asset_name.focus();
					return false;}
				if(document.frm.rental.value =="") {
					alert('구매구분을 선택하세요 !!!');
					frm.rental.focus();
					return false;}
				{
				a=confirm('입력하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
				}
			}
        </script>
	</head>
	<body onload="specview()">
			<div id="container">
				<h3 class="tit"><%=title_line%></h3>
				<form action="asset_code_save.asp" method="post" name="frm">
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableWrite">
						<colgroup>
							<col width="15%" >
							<col width="35%" >
							<col width="15%" >
							<col width="*" >
						</colgroup>
						<tbody>
							<tr>
								<th class="first">소유회사</th>
								<td class="left">
								  <%
                                    if	asset_company = "00" then
                                        k = 0
                                        Sql="select * from etc_code where etc_type = '75' and used_sw = 'Y' order by etc_name asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                        while not rs_etc.eof
                                            k = k + 1
                                            company_tab(k,1) = rs_etc("etc_name")
                                            company_tab(k,2) = mid(rs_etc("etc_code"),3,2)
                                            rs_etc.movenext()
                                        Wend
                                        rs_etc.close()						
                                    %>
                                  <select name="company" id="company" style="width:150px">
                                    <option value="">선택</option>
                                    <% 
                                            for kk = 1 to k
                                        %>
                                    <option value='<%=company_tab(kk,2)%>' <%If company_tab(kk,2) = asset_company then %>selected<% end if %>><%=company_tab(kk,1)%></option>
                                    <%
                                            next
                                        %>
                                  </select>
                                <%		else	%>
                                    <%=user_name%>
                                    <input name="company" type="hidden" id="company" value="<%=company%>">
                                <%	end if	%>
                                </td>
								<th>자산구분</th>
								<td class="left">
                                  <select name="gubun" id="select2" style="width:150px" onChange="specview()">
								  	<%
										u_gubun = "79" + cstr(gubun)
								
										if u_type <> "U" then	
											Sql="select * from etc_code where etc_type = '79' order by etc_code asc"
									%>
                                    <option value="">선택</option>
                                    <%
										  else
											Sql="select * from etc_code where etc_code = '" + u_gubun + "' order by etc_code asc"
										 end if

                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                    <% 
                                        While not rs_etc.eof 
                                    %>
                                    <option value='<%=mid(rs_etc("etc_code"),3,2)%>' <%If mid(rs_etc("etc_code"),3,2) = gubun then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                    <%
                                        rs_etc.movenext()  
                                        Wend 
                                        rs_etc.Close()
                                    %>
                                  </select>
                                  <input name="code_seq" type="hidden" id="code_seq" value="<%=code_seq%>">
                                </td>
							</tr>
							<tr>
								<th class="first">제조사</th>
								<td class="left">
								  <%
                                        Sql="select * from etc_code where etc_type = '21' order by etc_code asc"
                                        Rs_etc.Open Sql, Dbconn, 1
                                    %>
                                  <select name="maker" id="select" style="width:150px">
                                    <option value="">선택</option>
                                    <% 
                                        While not rs_etc.eof 
                                    %>
                                    <option value='<%=rs_etc("etc_name")%>' <%If rs_etc("etc_name") = maker then %>selected<% end if %>><%=rs_etc("etc_name")%></option>
                                    <%
                                        rs_etc.movenext()  
                                        Wend 
                                        rs_etc.Close()
                                    %>
                                  </select>
                                </td>
								<th>자산명</th>
								<td class="left"><input name="asset_name" type="text" id="asset_name" style="width:150px" onKeyUp="checklength(this,30)" value="<%=asset_name%>"></td>
							</tr>
							<tr id="spec_menu" style="display:none">
								<th class="first">스펙</th>
								<td  colspan="3" class="left">
                                CPU : <input name="cpu" type="text" id="cpu" style="width:100px" onKeyUp="checklength(this,20)" value="<%=cpu%>">
                                MEM : <input name="mem" type="text" id="mem" style="width:100px" onKeyUp="checklength(this,20)" value="<%=mem%>">
                                HDD : <input name="hdd" type="text" id="hdd" style="width:100px" onKeyUp="checklength(this,20)" value="<%=hdd%>">
                                OS : <input name="os" type="text" id="os" style="width:100px" onKeyUp="checklength(this,20)" value="<%=os%>">
                                </td>
							</tr>
							<tr>
								<th class="first">세부스펙</th>
								<td  colspan="3" class="left"><input name="spec" type="text" id="spec" style="width:500px" onKeyUp="checklength(this,50)" value="<%=spec%>"></td>
							</tr>
							<tr>
								<th class="first">구매구분</th>
								<td class="left">
                                <select name="rental" id="rental" style="width:150px">
                                    <option value="">선택</option>
                                    <option value="0" <%If rental = "0" then %>selected<% end if %>>렌탈</option>
                                    <option value="1" <%If rental = "1" then %>selected<% end if %>>구매</option>
                                </select>
            					</td>
								<th>단가</th>
								<td class="left"><input name="unit_price" type="text" id="unit_price" style="width:150px;text-align:right" value="<%=formatnumber(unit_price,0)%>" onKeyUp="plusComma(this);" > [ VAT 별도 ]</td>
							</tr>
						</tbody>
					</table>
				</div>
                <br>
                <div align=center>
                    <span class="btnType01"><input type="button" value="저장" onclick="javascript:frmcheck();" ID="Button1" NAME="Button1"></span>
                    <span class="btnType01"><input type="button" value="취소" onclick="javascript:goAction();"></span>
                </div>
				<input type="hidden" name="u_type" value="<%=u_type%>" ID="Hidden1">
			</form>
		</div>				
	</body>
</html>

