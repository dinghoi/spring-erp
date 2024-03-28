<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include virtual="/include/nkpmg_dbcon.asp" -->
<!--#include virtual="/include/nkpmg_user.asp" -->
<%
dim sch_tab(10,10)

user_name = request.cookies("nkpmg_user")("coo_user_name")
user_id = request.cookies("nkpmg_user")("coo_user_id")
org_name = request.cookies("nkpmg_user")("coo_org_name")
cost_grade = request.cookies("nkpmg_user")("coo_cost_grade")
emp_company = request.cookies("nkpmg_user")("coo_emp_company")

' 창고이동 출고의뢰서

curr_date = mid(cstr(now()),1,10)
curr_year = mid(cstr(now()),1,4)
curr_month = mid(cstr(now()),6,2)
curr_day = mid(cstr(now()),9,2)

rele_date = request("rele_date")
rele_stock = request("rele_stock")
rele_seq = request("rele_seq")

Set dbconn = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.Recordset")
Set Rs_org = Server.CreateObject("ADODB.Recordset")
Set Rs_etc = Server.CreateObject("ADODB.Recordset")
Set Rs_reg = Server.CreateObject("ADODB.Recordset")
Set Rs_chul = Server.CreateObject("ADODB.Recordset")
Set Rs_good = Server.CreateObject("ADODB.Recordset")
Set RsCount = Server.CreateObject("ADODB.Recordset")
dbconn.open dbconnect

sql = "select * from met_mv_reg where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')"
Set Rs_chul = DbConn.Execute(SQL)
if not Rs_chul.eof then
    	rele_stock = Rs_chul("rele_stock")
        rele_seq = Rs_chul("rele_seq")
	    rele_date = Rs_chul("rele_date")
        rele_id = Rs_chul("rele_id")
        rele_goods_type = Rs_chul("rele_goods_type")
		rele_stock_company = Rs_chul("rele_stock_company")
        rele_stock_name = Rs_chul("rele_stock_name")
        rele_emp_no = Rs_chul("rele_emp_no")
        rele_emp_name = Rs_chul("rele_emp_name")
        rele_company = Rs_chul("rele_company")
        rele_bonbu = Rs_chul("rele_bonbu")
        rele_saupbu = Rs_chul("rele_saupbu")
        rele_team = Rs_chul("rele_team")
        rele_org_name = Rs_chul("rele_org_name")

        chulgo_rele_date = Rs_chul("chulgo_rele_date")
		chulgo_ing = Rs_chul("chulgo_ing")
        chulgo_date = Rs_chul("chulgo_date")
        chulgo_stock = Rs_chul("chulgo_stock")
        chulgo_stock_name = Rs_chul("chulgo_stock_name")
	    chulgo_stock_company = Rs_chul("chulgo_stock_company")
	    rele_att_file = Rs_chul("rele_att_file")
	    rele_memo = Rs_chul("rele_memo")
        rele_sign_yn = Rs_chul("rele_sign_yn")
	    rele_sign_no = Rs_chul("rele_sign_no")
	    rele_sign_date = Rs_chul("rele_sign_date")
	    if chulgo_date = "0000-00-00" then
	          chulgo_date = ""
	    end if
   else
		rele_stock = ""
        rele_seq = ""
	    rele_date = ""
        rele_id = ""
        rele_goods_type = ""
        rele_stock_company = ""
        rele_stock_name = ""
		rele_emp_no = ""
        rele_emp_name = ""
        rele_company = ""
        rele_bonbu = ""
        rele_saupbu = ""
        rele_team = ""
        rele_org_name = ""

        chulgo_rele_date = ""
        chulgo_ing = ""
        chulgo_date = ""
        chulgo_stock = ""
        chulgo_stock_name = ""
	    chulgo_stock_company = ""
	    rele_att_file = ""
	    rele_memo = ""
        rele_sign_yn = ""
	    rele_sign_no = ""
	    rele_sign_date = ""
end if
Rs_chul.close()

sql = "select * from met_mv_reg_goods where (rele_date = '"&rele_date&"') and (rele_stock = '"&rele_stock&"') and (rele_seq = '"&rele_seq&"')  ORDER BY rl_goods_seq,rl_goods_code ASC"

Rs.Open Sql, Dbconn, 1

title_line = "창고이동 출고의뢰서"

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<script src="/java/common.js" type="text/javascript"></script>
<script type="text/javascript">
	function printWindow(){
//		viewOff("button");   
		factory.printing.header = ""; //머리말 정의
		factory.printing.footer = ""; //꼬리말 정의
		factory.printing.portrait = true; //출력방향 설정: true - 가로, false - 세로
		factory.printing.leftMargin = 13; //외쪽 여백 설정
		factory.printing.topMargin = 25; //윗쪽 여백 설정
		factory.printing.rightMargin = 13; //오른쯕 여백 설정
		factory.printing.bottomMargin = 15; //바닦 여백 설정
//		factory.printing.SetMarginMeasure(2); //테두리 여백 사이즈 단위를 인치로 설정
//		factory.printing.printer = ""; //프린터 할 프린터 이름
//		factory.printing.paperSize = "A4"; //용지선택
//		factory.printing.pageSource = "Manusal feed"; //종이 피드 방식
//		factory.printing.collate = true; //순서대로 출력하기
//		factory.printing.copies = "1"; //인쇄할 매수
//		factory.printing.SetPageRange(true,1,1); //true로 설정하고 1,3이면 1에서 3페이지 출력
//		factory.printing.Printer(true); //출력하기
		factory.printing.Preview(); //윈도우를 통해서 출력
		factory.printing.Print(false); //윈도우를 통해서 출력
	}
	function printW() {
        window.print();
    }
	function goBefore () {
		history.back() ;
	}
	
</script>
<title>인수증</title>
<style type="text/css">
<!--
    	.style12L {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
    	.style12R {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: right; }
        .style12C {font-size: 12px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style12BC {font-size: 12px; font-weight: bold; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style14L {font-size: 18px; font-family: "굴림체", "굴림체", Seoul; text-align: left; }
		.style18L {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style18C {font-size: 18px; font-family: "바탕체", "바탕체", Seoul; text-align: center; }
        .style20L {font-size: 20px; font-family: "바탕체", "바탕체", Seoul; text-align: left; }
        .style20C {font-size: 20px; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
        .style32BC {font-size: 32px; font-weight: bold; font-family: "굴림체", "굴림체", Seoul; text-align: center; }
		.style1 {font-size:12px;color: #666666}
		.style2 {font-size:10px;color: #666666}
		.style4 {font-size:14px;color: #666666}
-->
</style>
<style media="print"> 
.noprint     { display: none }
</style>
</head>

<body oncontextmenu="return false" ondragstart="return false" onselectstart="return false">
<div class="noprint">
<p><a href="#" onClick="printWindow()"><img src="image/printer.jpg" width="39" height="36" border="0" alt="출력하기" /></a></p>
</div>
<object id="factory" style="display:none;" viewastext classid="clsid:1663ed61-23eb-11d2-b92f-008048fdd814" codebase="/smsx.cab#Version=7.0.0.8">
</object>
<table width="690" cellpadding="0" cellspacing="0">
  <tr>
    <td class="style32BC"><%=title_line%></td>
  </tr>
  <tr>
    <td height="20" class="style20C">&nbsp;</td>
  </tr>
</table>
<table width="715" cellpadding="0" cellspacing="0">
  <tr>
    <td rowspan="4" width="4%" height="20" align="center" style="border-bottom:1px solid #000000; border-top:1px solid #000000; border-right:1px solid #000000; border-left:1px solid #000000;"><span class="style1">요<br>청<br>창<br>고</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-top:1px solid #000000; border-right:1px solid #000000;"><span class="style1">담 당</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-top:1px solid #000000; border-right:1px solid #000000;"><span class="style1">팀 장</span></td>
    <td width="*" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td rowspan="4" width="4%" height="20" align="center" style="border-bottom:1px solid #000000; border-top:1px solid #000000; border-right:1px solid #000000; border-left:1px solid #000000;"><span class="style1">출<br>고<br>창<br>고</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-top:1px solid #000000; border-right:1px solid #000000;"><span class="style1">담 당</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-top:1px solid #000000; border-right:1px solid #000000;"><span class="style1">팀 장</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-top:1px solid #000000; border-right:1px solid #000000;"><span class="style1">본부장</span></td>
  </tr> 
  <tr>
    <td width="10%" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="*" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
  </tr> 
  <tr>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="*" height="20" align="center" style="border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
    <td width="10%" height="20" align="center" style="border-bottom:1px solid #000000; border-right:1px solid #000000;"><span class="style1">&nbsp;</span></td>
  </tr> 
</table>    
<table width="690" cellpadding="0" cellspacing="0">
  <tr>
    <td height="20" class="style20C">&nbsp;</td>
  </tr>
  <tr>
    <td width="100%" height="20" align="left"><span class="style1"><span>&nbsp;의뢰일자 :&nbsp;<%=mid(cstr(rele_date),1,4)%>년&nbsp;<%=mid(cstr(rele_date),6,2)%>월&nbsp;<%=mid(cstr(rele_date),9,2)%>일</span></td>
  </tr>  
</table>
<table width="690" border="1px" cellpadding="15" cellspacing="0" bordercolor="#000000">
  <tr>
    <td style="border-bottom:none; border-top:none;">
     <table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>  
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">신청일자</span></td>
        <td width="*" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=rele_date%></strong></td>
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">용도구분</span></td>
        <td width="20%" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=rele_goods_type%></strong></td>
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">신청창고</span></td>
        <td width="20%" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=rele_stock_name%></strong></td>
      </tr>
      <tr>  
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">회사</span></td>
        <td width="*" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=rele_company%></strong></td>
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">사업부</span></td>
        <td width="20%" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=rele_saupbu%></strong></td>
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">신청자</span></td>
        <td width="20%" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=rele_emp_name%>&nbsp;<%=rele_org_name%></strong></td>
      </tr>
      <tr>  
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">출고요청일</span></td>
        <td width="*" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=chulgo_rele_date%></strong></td>
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">출고창고</span></td>
        <td colspan="3" width="20%" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=chulgo_stock_name%>&nbsp;(<%=chulgo_stock_company%>)</strong></td>
      </tr>
      <tr>  
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">실출고일</span></td>
        <td width="*" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=chulgo_date%></strong></td>
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">신청창고<br>입고일</span></td>
        <td colspan="3" width="20%" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=in_stock_date%></strong></td>
      </tr>
      <tr>  
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><span class="style1">비고</span></td>
        <td colspan="5" width="20%" height="30" align="left" valign="middle" ><span class="style1"><strong>&nbsp;<%=rele_memo%></strong></td>
      </tr>
    </table>
   </td>
  </tr>  
  <tr>
    <td class="style4" style="border-bottom:none; border-top:none;"><strong>❐ 인수물품 내역</strong></td>
  </tr>
  <tr>
    <td style="border-bottom:none; border-top:none;"><table width="680" border="1px" cellpadding="0" cellspacing="0">
      <tr>
        <td width="3%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">No</strong></td>
        <td width="12%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">용도구분</strong></td>
        <td width="*" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">품목구분</strong></td>
        <td width="14%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">품목코드</strong></td>
        <td width="14%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">품목명</strong></td>
        <td width="14%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">규격</strong></td>
        <td width="6%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">수량</strong></td>
        <td width="10%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">상태</strong></td>
        <td width="10%" height="30" align="center" valign="middle" style=" background-color:#eaeaea;"><strong class="style1">확인</strong></td>
      </tr>
   <%  g_seq = 0
      do until rs.eof or rs.bof	
	     g_seq = g_seq + 1
   %>
      <tr>
        <td width="3%" height="20" align="right" valign="middle" ><span class="style1"><span>&nbsp;<%=g_seq%></span></td>
        <td width="12%" height="20" align="left" valign="middle" ><span class="style1"><span>&nbsp;<%=rs("rl_goods_type")%></span></td>
        <td width="*" height="20" align="left" valign="middle" ><span class="style1"><span>&nbsp;<%=rs("rl_goods_gubun")%></span></td>
        <td width="14%" height="20" align="left" valign="middle" ><span class="style1"><span>&nbsp;<%=rs("rl_goods_code")%></span></td>
        <td width="14%" height="20" align="left" valign="middle"><span class="style1"><span>&nbsp;<%=rs("rl_goods_name")%></span></td>
        <td width="14%" height="20" align="left" valign="middle" ><span class="style1"><span>&nbsp;<%=rs("rl_standard")%></span></td>
        <td width="6%" height="20" align="right" valign="middle" ><span class="style1"><span><%=formatnumber(rs("rl_qty"),0)%></span>&nbsp;</td>
        <td width="10%" height="20" align="left" valign="middle" ><span class="style1"><span>&nbsp<%=rs("rl_goods_grade")%>;</span></td>
        <td width="10%" height="20" align="left" valign="middle"><span class="style1"><span>&nbsp;</span></td>
      </tr>
   <%
		rs.movenext()
	loop
	rs.close()
   %>
    </table></td>
  </tr>

  <tr>
    <td class="style1" align="center" style="border-bottom:none; border-top:none;"><span>위 품목에 대해 "케이원정보통신"으로부터 인수하였음을 확인 합니다.</span></td>
  </tr>
  <tr>
	<td class="style1" align="right" style="border-bottom:none; border-top:none;"><%=mid(cstr(now()),1,4)%>년&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;월&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;일&nbsp;&nbsp;&nbsp;&nbsp;<br/><br/>소&nbsp;&nbsp;&nbsp;&nbsp;속 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br/>인수자 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;(인 또는 서명)&nbsp;&nbsp;&nbsp;&nbsp;
    <br/><br/>
    <strong>서명 날인후 FAX 부탁드립니다.</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
    </td>
  </tr>
  
</table>

</p>	

</body>
</html>
