<!DOCTYPE html PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=EUC-KR">
<title>�ּ� ã��</title>
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
	'//**2017�� 2�� �߰� ���� **/
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
	'//**2017�� 3�� �߰� ���� **/
	emdNo = Request.Form("emdNo")

	'NKP ��� ���� �� �߰�[����ȣ_20220119]
	gubun = Request("gubun")
%>
</head>
<script language="javascript">
// opener���� ������ �߻��ϴ� ��� �Ʒ� �ּ��� �����ϰ�, ������� ������������ �Է��մϴ�. ("�ּ��Է�ȭ�� �ҽ�"�� �����ϰ� ������Ѿ� �մϴ�.)
//document.domain = "abc.go.kr";

/*
			���� ��ŷ �׽�Ʈ �� �˾�API�� ȣ���Ͻø� IP�� ���� �� �� �ֽ��ϴ�.
			�ּ��˾�API�� �����Ͻð� �׽�Ʈ �Ͻñ� �ٶ��ϴ�.
*/
function init(){
	var url = location.href;
	var confmKey = "U01TX0FVVEgyMDIyMDExODE0MjgxMzExMjE0OTI=";
	var resultType = "4"; // ���θ��ּ� �˻���� ȭ�� ��³���, 1 : ���θ�, 2 : ���θ�+����+�󼼺���(��������, �����ֹμ���), 3 : ���θ�+�󼼺���(�󼼰ǹ���), 4 : ���θ�+����+�󼼺���(��������, �����ֹμ���, �󼼰ǹ���)
	var inputYn= "<%=inputYn%>";
	if(inputYn != "Y"){
		document.form.confmKey.value = confmKey;
		document.form.returnUrl.value = url;
		document.form.resultType.value = resultType;
		document.form.action="https://www.juso.go.kr/addrlink/addrLinkUrl.do"; //���ͳݸ�
		//document.form.action="https://www.juso.go.kr/addrlink/addrMobileLinkUrl.do"; //����� ���� ���, ���ͳݸ�
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
		<!-- �ش�ý����� ���ڵ�Ÿ���� EUC-KR�ϰ�쿡�� �߰� START-->
		<input type="hidden" id="encodingType" name="encodingType" value="EUC-KR"/>
		<!-- �ش�ý����� ���ڵ�Ÿ���� EUC-KR�ϰ�쿡�� �߰� END-->
	</form>
</body>
</html>