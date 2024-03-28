<%@LANGUAGE="VBSCRIPT" CODEPAGE="949"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<script type="text/javascript">
function form_chk()
{
		
	if(document.form1.fileData.value =="") {
		alert('업로드 할 파일을 선택하세요');
		form1.fileData.focus();
		return false;}
	
	{
	a=confirm('업로드 하시겠습니까?')
	if (a==true) {
		return true;
	}
	return false;
	}
}
function fileUploadPreview(thisObj,preViewer) {

	if(!/(\.gif|\.jpg|\.jpeg)$/i.test(thisObj.value)) {
	    alert('gif 와 jpg 파일만 지원합니다.');
    	resetImage(thisObj);
	return;
	} 

	preViewer = (typeof(preViewer) == "object") ? preViewer : document.getElementById(preViewer);
	var ua = window.navigator.userAgent; 

	if (ua.indexOf("MSIE") > -1) {
	var img_path = "";
	img_path = thisObj.value;
	preViewer.style.filter = "progid:DXImageTransform.Microsoft.AlphaImageLoader(src='fi" + "le://" + img_path + "', sizingMethod='scale')";
	} else {
	preViewer.innerHTML = "";
	var W = preViewer.offsetWidth;
	var H = preViewer.offsetHeight;
	var tmpImage = document.createElement("img");
	preViewer.appendChild(tmpImage); 

	tmpImage.onerror = function () {
	return preViewer.innerHTML = "";
	} 

	tmpImage.onload = function () {
	if (this.width > W) {
	this.height = this.height / (this.width / W);
	this.width = W;
	}
	if (this.height > H) {
	this.width = this.width / (this.height / H);
	this.height = H;
	}
	}
	if (ua.indexOf("Firefox/3") > -1) {
	var picData = thisObj.files.item(0).getAsDataURL();
	tmpImage.src = picData;
	} else {
	tmpImage.src = "file://" + thisObj.value;
	}
	}
} 
function resetImage(thisObj)
{
	thisObj.outerHTML = thisObj.outerHTML
}

</script>
<style type="text/css">
.preView { width: 400px; height: 400px; text-align: center; border:1px solid silver; }
</style>
<title>무제 문서</title>
</head>

<body>
<form name="form1" enctype="multipart/form-data" method="post" action="file_att_ok.asp" onsubmit="return form_chk(this);">
	<input name="fileData" type="file" id="fileData" onchange="fileUploadPreview(this,'preView');" size="50" />
    <input type="submit" name="Submit" value="업로드">  
  <div id="preView" class="preView" title="미리보기"></div>
</form>
</body>
</html>
