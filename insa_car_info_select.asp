<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

gubun = request("gubun")

if gubun = "" then
   gubun = Request.Form("gubun")
end if

in_name = ""
If Request.Form("in_name")  <> "" Then 
  in_name = Request.Form("in_name") 
End If

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if  in_name = "" then
	first_view = "N"
	sql = "select * from car_info where (car_no = '" + in_name + "')"
end if
if  in_name <> "" then
	first_view = "Y"
	Sql = "select * from car_info where (car_no like '%" + in_name + "%') ORDER BY car_no ASC"
end if

Rs.Open Sql, Dbconn, 1

title_line = " 차량 검색 "

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
			function carsel(car_no,car_name,car_year,car_reg_date,owner_emp_name,owner_emp_no,car_use_dept,car_owner,oil_kind,gubun)
			{
				<%
				'alert(gubun);
				%>
				if(gubun =="caras")
					{ 
					opener.document.frm.car_no.value = car_no;
					opener.document.frm.car_name.value = car_name;
					opener.document.frm.car_year.value = car_year;
					opener.document.frm.car_reg_date.value = car_reg_date;
					opener.document.frm.owner_emp_name.value = owner_emp_name;
					opener.document.frm.owner_emp_no.value = owner_emp_no;
					opener.document.frm.car_use_dept.value = car_use_dept;
					opener.document.frm.oil_kind.value = oil_kind;
					opener.document.frm.car_owner.value = car_owner;
					window.close();
//					opener.document.frm.as_date.focus();
					}	
				if(gubun =="carpt")
					{ 
					opener.document.frm.car_no.value = car_no;
					opener.document.frm.car_name.value = car_name;
					opener.document.frm.car_year.value = car_year;
					opener.document.frm.car_reg_date.value = car_reg_date;
					opener.document.frm.owner_emp_name.value = owner_emp_name;
					opener.document.frm.owner_emp_no.value = owner_emp_no;
					opener.document.frm.car_use_dept.value = car_use_dept;
					opener.document.frm.car_owner.value = car_owner;
					opener.document.frm.oil_kind.value = oil_kind;
					window.close();
					opener.document.frm.pe_date.focus();
					}	
				
				<%	
				'else
				'	{ 
				'	opener.document.frm.sido.value = sido;
				'   opener.document.frm.family_gugun.value = gugun;
				'   opener.document.frm.family_dong.value = dong;
				'   opener.document.frm.family_zip.value = zip;
				'    window.close();
				'    opener.document.frm.family_addr.focus();
				'	}
				%>
			}			
			function frmcheck () {
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}
			
			function chkfrm() {
				if(document.frm.in_name.value =="") {
					alert('차량번호를 입력하세요');
					frm.in_name.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body>
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="insa_car_info_select.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>차량번호를 입력하세요 </strong>
								<label>
        						<input name="in_name" type="text" id="in_name" value="<%=in_name%>" style="width:150px; text-align:left; ime-mode:active">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><image src="/image/but_ser1.jpg" alt="검색"></a>
                            </p>
						</dd>
					</dl>
				</fieldset>
				<div class="gView">
					<table cellpadding="0" cellspacing="0" class="tableList">
						<colgroup>
							<col width="15%" >
							<col width="15%" >
                            <col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">차량번호</th>
								<th scope="col">차종</th>
								<th scope="col">소유</th>
                                <th scope="col">운행자</th>
 							</tr>
						</thead>   
						<tbody>
					<%
						if first_view = "Y" then 
						    i = 0
							do until rs.eof or rs.bof
							   i = i + 1
					%>
							<tr>
								<td class="first"><a href="#" onClick="carsel('<%=rs("car_no")%>','<%=rs("car_name")%>','<%=rs("car_year")%>','<%=rs("car_reg_date")%>','<%=rs("owner_emp_name")%>','<%=rs("owner_emp_no")%>','<%=rs("car_use_dept")%>','<%=rs("car_owner")%>','<%=rs("oil_kind")%>','<%=gubun%>');"><%=rs("car_no")%></a>
                                </td>
								<td><%=rs("car_name")%>&nbsp;</td>
                                <td><%=rs("car_owner")%>&nbsp;</td>
                                <td><%=rs("owner_emp_name")%>(<%=rs("owner_emp_no")%>)&nbsp;</td>
							</tr>
					<%
								rs.movenext()
							loop
							rs.close()
							
							if i = 0 then 
					%>
                            <tr>
								<td class="first" colspan="4">내역이 없습니다</td>
							</tr>
					<%      end if
						  else
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
			</div>				
	</div>
                <input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
	</form>
	</body>
</html>

