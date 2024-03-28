<!--#include virtual="/common/inc_top.asp"-->
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/common/inc_nkpmg_user.asp"-->
<!--#include virtual="/include/nkpmg_user.asp" -->
<!--#include virtual="/common/func.asp" -->
<!--#include virtual="/common/common.asp" -->
<%
'===================================================
'### 작업 내역
'===================================================
' 허정호_20210722 :
'	- 신규 페이지 작성 및 코드 정리
'	- AS등록, 과태료 숨김 처리(현재 등록 기능만 있으며 별도 관리 페이지 없음, 비용 관리에서 일반 경비로 별도 등록함)

'===================================================
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
Dim u_type, oil_unit_month, title_line
Dim oil_unit_middle11, oil_unit_last11, oil_unit_middle12, oil_unit_last12, oil_unit_middle13
Dim oil_unit_last13, oil_unit_middle21, oil_unit_last21, oil_unit_middle22, oil_unit_last22
Dim oil_unit_middle23, oil_unit_last23
Dim curr_month, rsCount, total_record, stpage, pgsize
Dim rs_max, curr_date, next_date, rsOil, rs_unit, u_btn

u_type = f_Request("u_type")
oil_unit_month = f_Request("oil_unit_month")

title_line = "월별 유류비 단가 관리"

If u_type = "U" Then
	objBuilder.Append "SELECT oil_unit_id, oil_kind, oil_unit_middle, oil_unit_last "
	objBuilder.Append "FROM oil_unit "
	objBuilder.Append "WHERE oil_unit_month = '"&oil_unit_month&"' "

	Set rs_unit = Server.CreateObject("ADODB.RecordSet")
	rs_unit.Open objBuilder.ToString(), DBConn, 1
	objBuilder.Clear()

	Do Until rs_unit.EOF
		If rs_unit("oil_unit_id") = "1" And rs_unit("oil_kind") = "휘발유" Then
			oil_unit_middle11 = rs_unit("oil_unit_middle")
			oil_unit_last11 = rs_unit("oil_unit_last")
		End If

		If rs_unit("oil_unit_id") = "1" And rs_unit("oil_kind") = "디젤" Then
			oil_unit_middle12 = rs_unit("oil_unit_middle")
			oil_unit_last12 = rs_unit("oil_unit_last")
		End If

		If rs_unit("oil_unit_id") = "1" And rs_unit("oil_kind") = "가스" Then
			oil_unit_middle13 = rs_unit("oil_unit_middle")
			oil_unit_last13 = rs_unit("oil_unit_last")
		End If

		If rs_unit("oil_unit_id") = "2" And rs_unit("oil_kind") = "휘발유" Then
			oil_unit_middle21 = rs_unit("oil_unit_middle")
			oil_unit_last21 = rs_unit("oil_unit_last")
		End If

		If rs_unit("oil_unit_id") = "2" And rs_unit("oil_kind") = "디젤" Then
			oil_unit_middle22 = rs_unit("oil_unit_middle")
			oil_unit_last22 = rs_unit("oil_unit_last")
		End If

		If rs_unit("oil_unit_id") = "2" And rs_unit("oil_kind") = "가스" Then
			oil_unit_middle23 = rs_unit("oil_unit_middle")
			oil_unit_last23 = rs_unit("oil_unit_last")
		End If

		rs_unit.MoveNext()
	Loop
	rs_unit.Close() : Set rs_unit = Nothing
Else
	oil_unit_middle11 = 0
	oil_unit_last11 = 0
	oil_unit_middle12 = 0
	oil_unit_last12 = 0
	oil_unit_middle13 = 0
	oil_unit_last13 = 0
	oil_unit_middle21 = 0
	oil_unit_last21 = 0
	oil_unit_middle22 = 0
	oil_unit_last22 = 0
	oil_unit_middle23 = 0
	oil_unit_last23 = 0

	objBuilder.Append "SELECT MAX(oil_unit_month) AS max_month FROM oil_unit "

	Set rs_max = DBConn.Execute(objBuilder.ToString())
	objBuilder.Clear()

	If IsNull(rs_max("max_month")) Then
		oil_unit_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)
	Else
		curr_date = Mid(rs_max("max_month"), 1, 4) & "-" & Mid(rs_max("max_month"), 5, 2) & "-01"
		curr_date = DateValue(curr_date)
		next_date = DateAdd("m", 1, curr_date)
		oil_unit_month = Mid(next_date, 1, 4) & Mid(next_date, 6, 2)
	End If
End If

curr_month = Mid(Now(), 1, 4) & Mid(Now(), 6, 2)

objBuilder.Append "SELECT COUNT(*) FROM oil_unit "

Set rsCount = DBConn.Execute(objBuilder.ToString())
objBuilder.Clear()

total_record = CInt(rsCount(0)) 'Result.RecordCount
stpage = total_record - 12

If stpage < 0 Then
	stpage = 0
End If

pgsize = 12

objBuilder.Append "SELECT oil_unit_id, oil_unit_month, oil_kind, oil_unit_middle, "
objBuilder.Append "	oil_unit_last, oil_unit_average "
objBuilder.Append "FROM oil_unit "
objBuilder.Append "ORDER BY oil_unit_month, oil_unit_id asc, oil_kind DESC "
objBuilder.Append "LIMIT "&stpage&", "&pgsize

Set rsOil = Server.CreateObject("ADODB.RecordSet")
rsOil.Open objBuilder.ToString(), DBConn, 1
objBuilder.Clear()
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
	<head>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<title>인사 관리 시스템</title>
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
			/*
			$(function(){
				$("#datepicker").datepicker();
				$("#datepicker").datepicker("option", "dateFormat", "yy-mm-dd");
				$("#datepicker").datepicker("setDate", "<%'=holiday%>");
			});*/

			/*function delcheck () {
				if (form_chk(document.frm_del)) {
					document.frm_del.submit ();
				}
			}

			function form_chk(){
				a=confirm('삭제하시겠습니까?')
				if (a==true) {
					return true;
				}
				return false;
			}//-->*/

			function frmcheck(type){
				if(chkfrm(type)){
					document.frm.submit();
				}
			}

			function chkfrm(type){
				var message = '등록 하시겠습니까?';

				if(document.frm.oil_unit_month.value > document.frm.curr_month.value){
					alert('발생일자가 현재일보다 클 수 없습니다.');
					frm.oil_unit_month.focus();
					return false;
				}

				if(type === 'U') message = '변경 하시겠습니까?';

				if(!confirm(message)) return false;
				else return true;
			}

			//유류단가 상세 링크[허정호_20210723]
			function oilDetailView(m, type){
				var param;

				param = 'oil_unit_month='+m+'&u_type='+type;
				location.href = '/insa/oil_unit_mg.asp?'+param;
			}
		</script>
	</head>
	<body oncontextmenu="return false" ondragstart="return false">
		<div id="wrap">
			<!--#include virtual = "/include/insa_header.asp" -->
			<!--#include virtual = "/include/insa_car_menu.asp" -->
			<div id="container">
				<h3 class="insa"><%=title_line%></h3>
				<div class="gView">
				  <table width="100%" border="0" cellpadding="0" cellspacing="0">
				    <tr>
				      <td width="59%" height="356" valign="top">
					  <!--<form action="/insa/holi_del_ok.asp" method="post" name="frm_del">-->
                      <table cellpadding="0" cellspacing="0" class="tableList">
				        <colgroup>
				          <col width="*" >
				          <col width="17%" >
				          <col width="17%" >
				          <col width="17%" >
				          <col width="17%" >
				          <col width="17%" >
			            </colgroup>
				        <thead>
				          <tr>
				            <th class="first" scope="col">년월</th>
				            <th scope="col">구분</th>
				            <th scope="col">유종</th>
				            <th scope="col">월초단가</th>
				            <th scope="col">월말단가</th>
				            <th scope="col">평균단가</th>
			              </tr>
			            </thead>
			            <tbody>
						<%
						Dim unit_id_view

                        Do Until rsOil.EOF
							If rsOil("oil_unit_id") = "1" Then
								unit_id_view = "본사팀"
							Else
								unit_id_view = "지방"
							End If
                         %>
                            <tr>
                            <td class="first">
								<a href="#" onclick="oilDetailView('<%=rsOil("oil_unit_month")%>', 'U');"><%=rsOil("oil_unit_month")%></a>
							</td>
                            <td><%=unit_id_view%></td>
                            <td><%=rsOil("oil_kind")%></td>
                            <td><%=FormatNumber(rsOil("oil_unit_middle"), 0)%></td>
                            <td><%=FormatNumber(rsOil("oil_unit_last"), 0)%></td>
                            <td><%=FormatNumber(rsOil("oil_unit_average"), 0)%></td>
                            </tr>
                            <%
							rsOil.MoveNext()
						Loop
						rsOil.Close() : Set rsOil = Nothing
						%>
			            </tbody>
			          </table>
					  <br>
                      <!--</form>-->
                      </td>
				      <td width="2%" valign="top">&nbsp;</td>
				      <td width="39%" valign="top">
						<form method="post" name="frm" action="/insa/oil_unit_reg_ok.asp">
				        <table cellpadding="0" cellspacing="0" summary="" class="tableWrite">
				        <colgroup>
				          <col width="*" >
				          <col width="30%" >
				          <col width="20%" >
				          <col width="30%" >
			            </colgroup>
				          <tbody>
				            <tr>
				              <th style="background-color:#E8FFFF">년월</th>
				              <td colspan="3" class="left" style="background-color:#E8FFFF"><input name="oil_unit_month" type="text" style="width:70px; text-align:center" readonly="true" value="<%=oil_unit_month%>"></td>
			                </tr>
				            <tr>
				              <th>구분</th>
				              <td class="left">본사팀</td>
				              <th>유종</th>
				              <td class="left">휘발유</td>
			                </tr>
				            <tr>
				              <th>월초단가</th>
				              <td class="left"><input name="oil_unit_middle11" type="text" id="oil_unit_middle11" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_middle11, 0)%>" onKeyUp="plusComma(this);" ></td>
				              <th>월말단가</th>
				              <td class="left"><input name="oil_unit_last11" type="text" id="oil_unit_last11" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_last11, 0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">구분</th>
				              <td class="left" style="background-color:#E8FFFF">본사팀</td>
				              <th style="background-color:#E8FFFF">유종</th>
				              <td class="left" style="background-color:#E8FFFF">디젤</td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">월초단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_middle12" type="text" id="oil_unit_middle12" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_middle12, 0)%>" onKeyUp="plusComma(this);" ></td>
				              <th style="background-color:#E8FFFF">월말단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_last12" type="text" id="oil_unit_last12" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_last12, 0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th>구분</th>
				              <td class="left">본사팀</td>
				              <th>유종</th>
				              <td class="left">가스</td>
			                </tr>
				            <tr>
				              <th>월초단가</th>
				              <td class="left"><input name="oil_unit_middle13" type="text" id="oil_unit_middle13" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_middle13, 0)%>" onKeyUp="plusComma(this);" ></td>
				              <th>월말단가</th>
				              <td class="left"><input name="oil_unit_last13" type="text" id="oil_unit_last13" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_last13, 0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">구분</th>
				              <td class="left" style="background-color:#E8FFFF">지방</td>
				              <th style="background-color:#E8FFFF">유종</th>
				              <td class="left" style="background-color:#E8FFFF">휘발유</td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">월초단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_middle21" type="text" id="oil_unit_middle21" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_middle21, 0)%>" onKeyUp="plusComma(this);" ></td>
				              <th style="background-color:#E8FFFF">월말단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_last21" type="text" id="oil_unit_last21" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_last21, 0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th>구분</th>
				              <td class="left">지방</td>
				              <th>유종</th>
				              <td class="left">디젤</td>
			                </tr>
				            <tr>
				              <th>월초단가</th>
				              <td class="left"><input name="oil_unit_middle22" type="text" id="oil_unit_middle22" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_middle22, 0)%>" onKeyUp="plusComma(this);" ></td>
				              <th>월말단가</th>
				              <td class="left"><input name="oil_unit_last22" type="text" id="oil_unit_last22" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_last22, 0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">구분</th>
				              <td class="left" style="background-color:#E8FFFF">지방</td>
				              <th style="background-color:#E8FFFF">유종</th>
				              <td class="left" style="background-color:#E8FFFF">가스</td>
			                </tr>
				            <tr>
				              <th style="background-color:#E8FFFF">월초단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_middle23" type="text" id="oil_unit_middle23" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_middle23, 0)%>" onKeyUp="plusComma(this);" ></td>
				              <th style="background-color:#E8FFFF">월말단가</th>
				              <td class="left" style="background-color:#E8FFFF"><input name="oil_unit_last23" type="text" id="oil_unit_last23" style="width:50px;text-align:right" value="<%=FormatNumber(oil_unit_last23, 0)%>" onKeyUp="plusComma(this);" ></td>
			                </tr>
			              </tbody>
			            </table>
						<br>
                        <%
                        If u_type = "U" Then
                            u_btn = "변경"
                        Else
                            u_btn = "등록"
                        End If
                        %>
				        <input type="hidden" name="u_type" value="<%=u_type%>" />
				        <input type="hidden" name="curr_month" value="<%=curr_month%>" />
				        <div align="center">
                        	<span class="btnType01">
								<input type="button" value="<%=u_btn%>" onclick="frmcheck('<%=u_type%>');" />
							</span>
                        </div>
			          </form>
                      </td>
			        </tr>
			      </table>
                </div>
			</div>
		</div>
	</body>
</html>
<!--#include virtual="/common/inc_footer.asp" -->