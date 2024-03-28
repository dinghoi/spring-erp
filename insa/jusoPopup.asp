<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=EUC-KR">
<title>주소 찾기</title>
<%
	inputYn = Request.Form("inputYn")
	roadFullAddr = Request.Form("roadFullAddr")
	roadAddrPart1 = Request.Form("roadAddrPart1")
	roadAddrPart2 = Request.Form("roadAddrPart2")
	engAddr = Request.Form("engAddr")
	jibunAddr = Request.Form("jibunAddr")
	zipNo = Request.Form("zipNo")
	addrDetail = Request.Form("addrDetail")
	admCd = Request.Form("admCd")
	rnMgtSn = Request.Form("rnMgtSn")
	bdMgtSn = Request.Form("bdMgtSn")
	detBdNmList = Request.Form("detBdNmList")
	'//**2017년 2월 추가 제공 **/
	bdNm = Request.Form("bdNm")
	bdKdcd = Request.Form("bdKdcd")
	siNm = Request.Form("siNm")
	sggNm = Request.Form("sggNm")
	emdNm = Request.Form("emdNm")
	liNm = Request.Form("liNm")
	rn = Request.Form("rn")
	udrtYn = Request.Form("udrtYn")
	buldMnnm = Request.Form("buldMnnm")
	buldSlno = Request.Form("buldSlno")
	mtYn = Request.Form("mtYn")
	lnbrMnnm = Request.Form("lnbrMnnm")
	lnbrSlno = Request.Form("lnbrSlno")
	'//**2017년 3월 추가 제공 **/
	emdNo = Request.Form("emdNo")

	'NKP 사용 구분 값 추가[허정호_20220119]
	gubun = Request("gubun")
%>
</head>
<script language="javascript">
// opener관련 오류가 발생하는 경우 아래 주석을 해지하고, 사용자의 도메인정보를 입력합니다. ("주소입력화면 소스"도 동일하게 적용시켜야 합니다.)
//document.domain = "abc.go.kr";

/*
			모의 해킹 테스트 시 팝업API를 호출하시면 IP가 차단 될 수 있습니다.
			주소팝업API를 제외하시고 테스트 하시기 바랍니다.
*/
function init(){
	var url = location.href;
	var confmKey = "U01TX0FVVEgyMDIyMDExODE0MjgxMzExMjE0OTI=";
	var resultType = "4"; // 도로명주소 검색결과 화면 출력내용, 1 : 도로명, 2 : 도로명+지번+상세보기(관련지번, 관할주민센터), 3 : 도로명+상세보기(상세건물명), 4 : 도로명+지번+상세보기(관련지번, 관할주민센터, 상세건물명)
	var inputYn= "<%=inputYn%>";
	if(inputYn != "Y"){
		document.form.confmKey.value = confmKey;
		document.form.returnUrl.value = url;
		document.form.resultType.value = resultType;
		document.form.action="https://www.juso.go.kr/addrlink/addrLinkUrl.do"; //인터넷망
		//document.form.action="https://www.juso.go.kr/addrlink/addrMobileLinkUrl.do"; //모바일 웹인 경우, 인터넷망
		document.form.submit();
	}else{
		opener.jusoCallBack("<%=roadFullAddr%>","<%=roadAddrPart1%>","<%=addrDetail%>","<%=roadAddrPart2%>","<%=engAddr%>","<%=jibunAddr%>","<%=zipNo%>", "<%=admCd%>", "<%=rnMgtSn%>", "<%=bdMgtSn%>", "<%=detBdNmList%>", "<%=bdNm%>", "<%=bdKdcd%>", "<%=siNm%>", "<%=sggNm%>", "<%=emdNm%>", "<%=liNm%>", "<%=rn%>", "<%=udrtYn%>", "<%=buldMnnm%>", "<%=buldSlno%>", "<%=mtYn%>", "<%=lnbrMnnm%>", "<%=lnbrSlno%>", "<%=emdNo%>", "<%=gubun%>");
		window.close();
	}
}
</script>
<body onload="init();">
	<form id="form" name="form" method="post">
		<input type="hidden" id="confmKey" name="confmKey" value=""/>
		<input type="hidden" id="returnUrl" name="returnUrl" value=""/>
		<input type="hidden" id="resultType" name="resultType" value=""/>
		<!-- 해당시스템의 인코딩타입이 EUC-KR일경우에만 추가 START-->
		<input type="hidden" id="encodingType" name="encodingType" value="EUC-KR"/>
		<!-- 해당시스템의 인코딩타입이 EUC-KR일경우에만 추가 END-->
	</form>
</body>
</html>