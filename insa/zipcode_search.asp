<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/common.asp" -->
<!--#include virtual="/common/func.asp" -->
<%
'===================================================
'### DB Connection
'===================================================
Dim DBConn
Set DBConn = Server.CreateObject("ADODB.Connection")
DBConn.Open DbConnect

'===================================================
'### StringBuilder Object
'===================================================
Dim objBuilder
Set objBuilder = New StringBuilder

'===================================================
'### Request & Params
'===================================================
Dim gubun, in_dong, rsZip, arrZip, be_pg, title_line, i

gubun = f_Request("gubun")
in_dong = f_Request("in_dong")

be_pg = "/insa/zipcode_search.asp"
title_line = "◈ 주소(동코드) 검색 ◈"

objBuilder.Append "CALL USP_COMM_ZIPCODE_INFO('"&in_dong&"')"

Call Rs_Open(rsZip, DBConn, objBuilder.ToString())
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>주소(동코드) 검색</title>
		<link href="/include/jquery-ui.css" type="text/css" rel="stylesheet">
		<link href="/include/style.css" type="text/css" rel="stylesheet">
	  	<script src="/java/jquery-1.9.1.js"></script>
	  	<script src="/java/jquery-ui.js"></script>
		<script src="/java/common.js" type="text/javascript"></script>
		<script src="/java/ui.js" type="text/javascript"></script>
		<script type="text/javascript" src="/java/js_form.js"></script>
		<script type="text/javascript" src="/java/js_window.js"></script>

		<script type="text/javascript">
			function zip_Area(sido,gugun,dong,zip,gubun){
				if(gubun =="family"){
					opener.document.frm.emp_family_sido.value = sido;
				    opener.document.frm.emp_family_gugun.value = gugun;
				    opener.document.frm.emp_family_dong.value = dong;
				    opener.document.frm.emp_family_zip.value = zip;
				    window.close();
				    opener.document.frm.emp_family_addr.focus();
				}

				if(gubun =="org"){
					opener.document.frm.org_sido.value = sido;
				    opener.document.frm.org_gugun.value = gugun;
				    opener.document.frm.org_dong.value = dong;
				    opener.document.frm.org_zip.value = zip;
				    window.close();
				    opener.document.frm.org_addr.focus();
				}

				if(gubun =="juso"){
					opener.document.frm.emp_sido.value = sido;
				    opener.document.frm.emp_gugun.value = gugun;
				    opener.document.frm.emp_dong.value = dong;
				    opener.document.frm.emp_zipcode.value = zip;
				    window.close();
				    opener.document.frm.emp_addr.focus();
				}

				if(gubun =="stay"){
					opener.document.frm.stay_sido.value = sido;
				    opener.document.frm.stay_gugun.value = gugun;
				    opener.document.frm.stay_dong.value = dong;
				    opener.document.frm.stay_zip.value = zip;
				    window.close();
				    opener.document.frm.stay_addr.focus();
				}
			}

			function frmcheck(){
				if (formcheck(document.frm) && chkfrm()) {
					document.frm.submit ();
				}
			}

			function chkfrm(){
				if(document.frm.in_dong.value =="") {
					alert('동명을 입력하세요');
					frm.in_dong.focus();
					return false;}
				{
					return true;
				}
			}
		</script>

	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<form action="<%=be_pg%>?gubun=<%=gubun%>" method="post" name="frm">
				<fieldset class="srch">
					<legend>조회영역</legend>
					<dl>
                        <dd>
                            <p>
							<strong>동명을 입력하세요 </strong>
								<label>
        						<input name="in_dong" type="text" id="in_dong" value="<%=in_dong%>" style="text-align:left; width:150px">
								</label>
                                <a href="#" onclick="javascript:frmcheck();"><img src="/image/but_ser1.jpg" alt="검색"></a>
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
								<th class="first" scope="col">우편번호</th>
								<th scope="col">시도</th>
								<th scope="col">구군</th>
								<th scope="col">동</th>
							</tr>
						</thead>
						<tbody>
                    	<%
						Dim z_code, z_sido, z_gugun, z_dong, v_cnt

						v_cnt = 0
						If Not rsZip.EOF Then
							arrZip = rsZip.getRows()

							For i = LBound(arrZip) To UBound(arrZip, 2)
								z_code = arrZip(0, i)
								z_sido = arrZip(1, i)
								z_gugun = arrZip(2, i)
								z_dong = arrZip(3, i)

								v_cnt = v_cnt + 1
						%>
							<tr>
                                <td class="first"><%=z_code%></td>
								<td><%=z_sido%></td>
								<td><%=z_gugun%></td>
								<td>
                                <a href="#" onClick="zip_Area('<%=z_sido%>','<%=z_gugun%>','<%=z_dong%>','<%=z_code%>','<%=gubun%>');"><%=z_dong%></a>
                                </td>
							</tr>
						<%
							Next

							Call Rs_Close(rsZip)
						Else
						%>
						<tr>
							<td class="first" colspan="4" style=" border-top:1px solid #e3e3e3;font-weight:bold;">검색 내용이 없습니다.</td>
						</tr>
						<%
						End If
						%>
						<tr>
							<td class="first" colspan="4" style=" border-top:1px solid #e3e3e3;">(<%=v_cnt%>)&nbsp;건이 조회되었습니다.</td>
						</tr>
						</tbody>
					</table>
				</div>
			</form>
		</div>
	</body>
</html>
<!--#include virtual="/common/inc_bottom.asp"-->