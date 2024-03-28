<script type="text/javascript">

    //◈◈ 업로드 체크 ◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈
    function fileCheck(fileValue)
    {
        //확장자 체크
        var src = getFileType(fileValue);

        if(!(src.toLowerCase() == "zip")))
        {
            alert("zip 파일로 압축하여 첨부해주세요.");
            return;
        }

        //사이즈체크
        var maxSize  = 31457280    //30MB
         var fielSize = Math.round(fileValue.fileSize);

        if(fileSize > maxSize)
        {
            alert("첨부파일 사이즈는 30MB 이내로 등록 가능합니다.    ");
            return;
        }

        form.submit();
    }

    
    //◈◈ 파일 확장자 확인 ◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈◈
    function getFileType(filePath)
    {
        var index = -1;
            index = filePath.lastIndexOf('.');

        var type = "";

        if(index != -1)
        {
            type = filePath.substring(index+1, filePath.len);
        }
        else
        {
            type = "";
        }

        return type;
    }

</script>

-------------------------------------------------------------------------------


<form name="frm">
    <input type="file" name="file1" />
    <input type="button" value="upload" onclick="fileCheck(document.frm.file1.value)">
</form>

