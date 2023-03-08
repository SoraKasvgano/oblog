function loadTxt()
	{
    var txtLang = document.getElementsByName("txtLang");
    txtLang[0].innerHTML = "\u67E5\u627E ";
    txtLang[1].innerHTML = "\u66ff\u6362 ";
    txtLang[2].innerHTML = "\u533A\u5206\u5927\u5C0F\u5199 ";
    txtLang[3].innerHTML = "\u5168\u5B57\u5339\u914D ";
    
    document.getElementById("btnSearch").value = "\u67E5\u627E\u4e0b\u4e00\u4e2a ";;
    document.getElementById("btnReplace").value = "\u66ff\u6362 ";
    document.getElementById("btnReplaceAll").value = "\u5168\u90e8\u66ff\u6362 ";
    document.getElementById("btnClose").value = "\u5173\u95ed ";
	}
function getTxt(s)
    {
    switch(s)
        {
        case "Finished searching": return "\u6587\u6863\u641c\u7d22\u7ed3\u675f .\n\u662f\u5426\u4ece\u5934\u5f00\u59cb\u641c\u7d22?";
        default: return "";
        }
    }
function writeTitle()
	{
	document.write("<title>\u67E5\u627E\u548c\u66ff\u6362  </title>")
	}
