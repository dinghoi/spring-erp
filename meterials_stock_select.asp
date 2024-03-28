<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
Dim in_name
Dim rs
Dim rs_numRows

gubun = request("gubun")
view_condi=Request("view_condi")

if gubun = "" then
   gubun = Request.Form("gubun")
   view_condi = Request.Form("view_condi")
end if

in_name = ""
If Request.Form("in_name")  <> "" Then 
  in_name = Request.Form("in_name") 
End If

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

if view_condi = "" and in_name = "" then
	first_view = "N"
	sql = "select * from met_stock_code where (stock_name = '" + in_name + "')"
end if
if view_condi = "" and in_name <> "" then
	first_view = "Y"
	Sql = "select * from met_stock_code where (stock_name like '%" + in_name + "%') ORDER BY stock_company,stock_level,stock_bonbu,stock_name ASC"
end if

if view_condi <> "" and in_name = "" then
	first_view = "N"
	sql = "select * from met_stock_code where  (stock_name = '" + in_name + "')"
end if
if view_condi <> "" and in_name <> "" then
	first_view = "Y"
	Sql = "select * from met_stock_code where  (stock_name like '%" + in_name + "%') ORDER BY stock_company,stock_level,stock_bonbu,stock_name ASC"
end if

Rs.Open Sql, Dbconn, 1

title_line = " 창고 검색 "

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
			function stockesel(stock_code,stock_level,stock_name,stock_manager_code,stock_manager_name,stock_company,stock_bonbu,stock_saupbu,stock_team,gubun)
			{
				<%
				'alert(gubun);
				%>
				if(gubun =="order")
					{ 
					opener.document.frm.order_stock_company.value = stock_company;
					opener.document.frm.order_stock_name.value = stock_name;
					opener.document.frm.order_stock_code.value = stock_code;
					opener.document.frm.stock_bonbu.value = stock_bonbu;
					opener.document.frm.stock_saupbu.value = stock_saupbu;
					opener.document.frm.stock_team.value = stock_team;
					opener.document.frm.stock_manager_code.value = stock_manager_code;
					opener.document.frm.stock_manager_name.value = stock_manager_name;
					window.close();
					opener.document.frm.order_in_date.focus();
					}	
				if(gubun =="stin")
					{ 
					opener.document.frm.stin_stock_company.value = stock_company;
					opener.document.frm.stin_stock_name.value = stock_name;
					opener.document.frm.stin_stock_code.value = stock_code;
					opener.document.frm.stock_bonbu.value = stock_bonbu;
					opener.document.frm.stock_saupbu.value = stock_saupbu;
					opener.document.frm.stock_team.value = stock_team;
					opener.document.frm.stock_manager_code.value = stock_manager_code;
					opener.document.frm.stock_manager_name.value = stock_manager_name;
					window.close();
					opener.document.frm.bill_collect.focus();
					}	
				if(gubun =="chulgo")
					{ 
					opener.document.frm.chulgo_stock_company.value = stock_company;
					opener.document.frm.chulgo_stock_name.value = stock_name;
					opener.document.frm.chulgo_stock.value = stock_code;
					opener.document.frm.stock_bonbu.value = stock_bonbu;
					opener.document.frm.stock_saupbu.value = stock_saupbu;
					opener.document.frm.stock_team.value = stock_team;
					opener.document.frm.stock_manager_code.value = stock_manager_code;
					opener.document.frm.stock_manager_name.value = stock_manager_name;
					window.close();
//					opener.document.frm.chulgo_date.focus();
					}	
				if(gubun =="mvreg")
					{ 
					opener.document.frm.rele_stock_company.value = stock_company;
					opener.document.frm.rele_stock_name.value = stock_name;
					opener.document.frm.rele_stock.value = stock_code;
					opener.document.frm.rele_stock_bonbu.value = stock_bonbu;
					opener.document.frm.rele_stock_saupbu.value = stock_saupbu;
					opener.document.frm.rele_stock_team.value = stock_team;
					opener.document.frm.rele_manager_code.value = stock_manager_code;
					opener.document.frm.rele_manager_name.value = stock_manager_name;
					opener.document.frm.rele_company.value = stock_company;
					opener.document.frm.rele_saupbu.value = stock_saupbu;
					window.close();
					opener.document.frm.chulgo_memo.focus();
					}
				if(gubun =="mvin")
					{ 
					opener.document.frm.chulgo_stock_company.value = stock_company;
					opener.document.frm.chulgo_stock_name.value = stock_name;
					opener.document.frm.chulgo_stock.value = stock_code;
					opener.document.frm.stock_bonbu.value = stock_bonbu;
					opener.document.frm.stock_saupbu.value = stock_saupbu;
					opener.document.frm.stock_team.value = stock_team;
					opener.document.frm.stock_manager_code.value = stock_manager_code;
					opener.document.frm.stock_manager_name.value = stock_manager_name;
					window.close();
					opener.document.frm.rele_memo.focus();
					}	
				if(gubun =="sale")
					{ 
					opener.document.frm.chulgo_stock_company.value = stock_company;
					opener.document.frm.chulgo_stock_name.value = stock_name;
					opener.document.frm.chulgo_stock.value = stock_code;
					opener.document.frm.stock_bonbu.value = stock_bonbu;
					opener.document.frm.stock_saupbu.value = stock_saupbu;
					opener.document.frm.stock_team.value = stock_team;
					opener.document.frm.stock_manager_code.value = stock_manager_code;
					opener.document.frm.stock_manager_name.value = stock_manager_name;
					window.close();
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
					alert('창고명을 입력하세요');
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
				<form action="meterials_stock_select.asp?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>					
                        <dd>
                            <p>
							<strong>창고명을 입력하세요 </strong>
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
							<col width="15%" >
                            <col width="15%" >
							<col width="*" >
						</colgroup>
						<thead>
							<tr>
								<th class="first" scope="col">창고명</th>
								<th scope="col">창고Level</th>
								<th scope="col">창고장</th>
                                <th scope="col">소속</th>
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
								<td class="first"><a href="#" onClick="stockesel('<%=rs("stock_code")%>','<%=rs("stock_level")%>','<%=rs("stock_name")%>','<%=rs("stock_manager_code")%>','<%=rs("stock_manager_name")%>','<%=rs("stock_company")%>','<%=rs("stock_bonbu")%>','<%=rs("stock_saupbu")%>','<%=rs("stock_team")%>','<%=gubun%>');"><%=rs("stock_name")%></a>
                                </td>
								<td><%=rs("stock_level")%>&nbsp;</td>
                                <td><%=rs("stock_manager_name")%>(<%=rs("stock_manager_code")%>)&nbsp;</td>
								<td class="left"><%=rs("stock_company")%> - <%=rs("stock_bonbu")%> - <%=rs("stock_saupbu")%>&nbsp;</td>
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
                <input type="hidden" name="mg_level" value="<%=mg_level%>" ID="Hidden1">
                <input type="hidden" name="view_condi" value="<%=view_condi%>" ID="Hidden1">    
            				
	</form>
	</body>
</html>

