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

if in_name = "" then
	   SQL = "select * from trade where trade_name = '" + in_name + "' ORDER BY trade_name ASC"
    else
	   SQL = "select * from trade where trade_name like '%" + in_name + "%' ORDER BY trade_name ASC"
end if

Rs.Open Sql, Dbconn, 1

title_line = " 거래처 검색 "

%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
"http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>거래처 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>
		<script type="text/javascript">
			function tradesel(trade_name,trade_no,trade_person,trade_email,trade_person_tel,gubun)
			{
				<%
				'alert(gubun);
				%>
				if(gubun =="buy")
					{ 
					opener.document.frm.trade_no.value = trade_no;
					opener.document.frm.trade_name.value = trade_name;
					opener.document.frm.trade_person.value = trade_person;
					opener.document.frm.trade_email.value = trade_email;
					window.close();
					opener.document.frm.buy_memo.focus();
					}	
				if(gubun =="order")
					{ 
					opener.document.frm.trade_no.value = trade_no;
					opener.document.frm.trade_name.value = trade_name;
					opener.document.frm.trade_person.value = trade_person;
					window.close();
					opener.document.frm.trade_person.focus();
					}	
				if(gubun =="chulgo")
					{ 
//					opener.document.frm.trade_no.value = trade_no;
					opener.document.frm.chulgo_trade_name.value = trade_name;
//					opener.document.frm.trade_person.value = trade_person;
					window.close();
					opener.document.frm.chulgo_trade_dept.focus();
					}	
				if(gubun =="sale")
					{ 
					opener.document.frm.trade_no.value = trade_no;
					opener.document.frm.trade_name.value = trade_name;
					opener.document.frm.trade_person.value = trade_person;
					opener.document.frm.trade_email.value = trade_email;
					window.close();
					opener.document.frm.trade_person.focus();
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
					alert('거래처명을 입력하세요');
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
				<form action="insa_trade_select.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>거래처명을 입력하세요 </strong>
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
							<col width="20%" >
							<col width="20%" >
							<col width="20%" >
                            <col width="20%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">구매처명</th>
								<th scope="col">사업자번호</th>
                                <th scope="col">담당자</th>
								<th scope="col">이메일</th>
								<th scope="col">연락처</th>
 							</tr>
						</thead>
						<tbody>
					<%
						    i = 0
							do until rs.eof or rs.bof
							   i = i + 1
							   
							   trade_code = rs("trade_code")
							   trade_name = rs("trade_name")
							   trade_no = mid(rs("trade_no"),1,3) + "-" + mid(rs("trade_no"),4,2) + "-" + mid(rs("trade_no"),6)
							   trade_person = rs("trade_person")
							   trade_email = rs("trade_email")
							   trade_person_tel = rs("trade_person_tel")
					%>
							<tr>
								<td class="first"><a href="#" onClick="tradesel('<%=trade_name%>','<%=trade_no%>','<%=trade_person%>','<%=trade_email%>','<%=trade_person_tel%>','<%=gubun%>');"><%=rs("trade_name")%></a>
                                </td>
								<td><%=rs("trade_no")%>&nbsp;</td>
                                <td><%=rs("trade_person")%>&nbsp;</td>
                                <td><%=rs("trade_email")%>&nbsp;</td>
                                <td><%=rs("trade_person_tel")%>&nbsp;</td>
							</tr>
					<%
								rs.movenext()
							loop
							rs.close()
							
							if i = 0 then 
					%>
                            <tr>
								<td class="first" colspan="5">내역이 없습니다</td>
							</tr>
					<%      end if  %>
						</tbody>
					</table>
				</div>
			</div>				
	</div>
                <input type="hidden" name="gubun" value="<%=gubun%>" ID="Hidden1">
	</form>
	</body>
</html>

