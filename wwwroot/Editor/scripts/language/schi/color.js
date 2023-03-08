function loadTxt()
    {
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "\u7f51\u9875\u8c03\u8272\u76d8 ";
    txtLang[1].innerHTML = "\u5185\u7F6E\u989c\u8272 ";
    txtLang[2].innerHTML = "\u7F51\u9875\u5B89\u5168\u8272 ";
    txtLang[3].innerHTML = "\u65b0\u7684\u989c\u8272 ";
    txtLang[4].innerHTML = "\u73b0\u5728\u7684\u989c\u8272 ";
    txtLang[5].innerHTML = "\u81EA\u5B9A\u4E49\u989c\u8272 ";
    
    document.getElementById("btnAddToCustom").value = "\u65b0\u589e\u5230\u81EA\u5B9A\u4E49\u989c\u8272 ";
    document.getElementById("btnCancel").value = " \u53d6\u6d88  ";
    document.getElementById("btnRemove").value = " \u5220\u9664\u989c\u8272  ";
    document.getElementById("btnApply").value = " \u5e94\u7528  ";
    document.getElementById("btnOk").value = " \u786e\u8ba4  ";
    }
function writeTitle()
    {
    document.write("<title>\u8272\u5f69 </title>")
    }
