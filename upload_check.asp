<script type="text/javascript">

    //�¢� ���ε� üũ �¢¢¢¢¢¢¢¢¢¢¢¢¢¢¢¢¢¢¢�
    function fileCheck(fileValue)
    {
        //Ȯ���� üũ
        var src = getFileType(fileValue);

        if(!(src.toLowerCase() == "zip")))
        {
            alert("zip ���Ϸ� �����Ͽ� ÷�����ּ���.");
            return;
        }

        //������üũ
        var maxSize  = 31457280    //30MB
         var fielSize = Math.round(fileValue.fileSize);

        if(fileSize > maxSize)
        {
            alert("÷������ ������� 30MB �̳��� ��� �����մϴ�.    ");
            return;
        }

        form.submit();
    }

    
    //�¢� ���� Ȯ���� Ȯ�� �¢¢¢¢¢¢¢¢¢¢¢¢¢¢¢¢�
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

