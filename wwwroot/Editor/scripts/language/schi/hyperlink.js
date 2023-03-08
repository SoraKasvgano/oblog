function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "\u94fe\u63a5 ";
    txtLang[1].innerHTML = "\u951a\u70b9 ";
    txtLang[2].innerHTML = "\u76ee\u6807 ";
    txtLang[3].innerHTML = "\u6807\u9898";

    var optLang = document.getElementsByName("optLang");
    optLang[0].text = "\u5F53\u524D\u7A97\u53E3"
    optLang[1].text = "\u65B0\u5EFA\u7A97\u53E3"
    optLang[2].text = "\u7236\u7EA7\u7A97\u53E3"
    
    document.getElementById("btnCancel").value = "\u53d6\u6d88 ";
    document.getElementById("btnInsert").value = "\u63d2\u5165 ";
    document.getElementById("btnApply").value = "\u5e94\u7528 ";
    document.getElementById("btnOk").value = " \u786e\u8ba4  ";
    }
function writeTitle()
    {
    document.write("<title>\u8d85\u7ea7\u94fe\u63a5 </title>")
    }
