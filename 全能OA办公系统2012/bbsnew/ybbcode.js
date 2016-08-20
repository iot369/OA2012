var Quote = 0;
var Bold  = 0;
var Italic = 0;
var Underline = 0;
var Code = 0;
function fontchuli(){
if ((document.selection)&&(document.selection.type == "Text")) {
var range = document.selection.createRange();
var ch_text=range.text;
range.text = fontbegin + ch_text + fontend;
} 
else {
document.kbbs.body.value=fontbegin+document.kbbs.body.value+fontend;
document.kbbs.body.focus();
}
}
function AddText(text) {
	if (document.kbbs.body.createTextRange && document.kbbs.body.caretPos) {      
		var caretPos = document.kbbs.body.caretPos;      
		caretPos.text = caretPos.text.charAt(caretPos.text.length - 1) == ' ' ?
		text + ' ' : text;
	}
	else document.kbbs.body.value += text;
	document.kbbs.body.focus(caretPos);
}
function COLOR(theSmilie){
var text=prompt("请输入文字", "");
if(text){
document.kbbs.body.value += '[color=' + theSmilie + ']'+ text + '[/color]';
}
}
helpstat = false;
stprompt = true;
basic = false;
function thelp(swtch){
	if (swtch == 1){
		basic = false;
		stprompt = false;
		helpstat = true;
	} else if (swtch == 0) {
		helpstat = false;
		stprompt = false;
		basic = true;
	} else if (swtch == 2) {
		helpstat = false;
		basic = false;
		stprompt = true;
	}
}
function Cswf() {
 	if (helpstat){
		alert("Flash\nFlash 动画.\n用法: [flash=宽度, 高度]Flash 文件的地址[/flash]");
	} else if (basic) {
		AddTxt="[flash=500,350][/flash]";
		AddText(AddTxt);
	} else {                  
		txt2=prompt("flash宽度，高度","500,350"); 
		if (txt2!=null) {
                txt=prompt("Flash 文件的地址","http://");
		if (txt!=null) {
                          if (txt2=="") {             
			AddTxt="[flash=500,350]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
               } else {
		        AddTxt="[flash="+txt2+"]"+txt;
			AddText(AddTxt);
			AddTxt="[/flash]";
			AddText(AddTxt);
		 }        
	    }  
       }
    }
}

function Crm() {
	if (helpstat) {
               alert("realplay\n播放realplay文件.\n用法: [rm=宽度, 高度]文件地址[/rm]");
	} else if (basic) {
		AddTxt="[rm=500,350][/rm]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("视频的宽度，高度","500,350"); 
		if (txt2!=null) {
			txt=prompt("视频文件的地址","请输入");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[rm=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/rm]";
					AddText(AddTxt);
				} else {
					AddTxt="[rm="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/rm]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}

function Cwmv() {
	if (helpstat) {
               alert("Media Player\n播放Media Player文件.\n用法: [mp=宽度, 高度]文件地址[/mp]");
	} else if (basic) {
		AddTxt="[mp=500,350][/mp]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("视频的宽度，高度","500,350"); 
		if (txt2!=null) {
			txt=prompt("视频文件的地址","请输入");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[mp=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/mp]";
					AddText(AddTxt);
				} else {
					AddTxt="[mp="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/mp]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}
function Cdir() {
	if (helpstat) {
               alert("Shockwave\n插入Shockwave文件.\n用法: [dir=宽度, 高度]文件地址[/dir]");
	} else if (basic) {
		AddTxt="[dir=500,350][/dir]";
		AddText(AddTxt);
	} else { 
		txt2=prompt("Shockwave文件的宽度，高度","500,350"); 
		if (txt2!=null) {
			txt=prompt("Shockwave文件的地址","请输入地址");
			if (txt!=null) {
				if (txt2=="") {
					AddTxt="[dir=500,350]"+txt;
					AddText(AddTxt);
					AddTxt="[/dir]";
					AddText(AddTxt);
				} else {
					AddTxt="[dir="+txt2+"]"+txt;
					AddText(AddTxt);
					AddTxt="[/dir]";
					AddText(AddTxt);
				}         
			} 
		}
	}
}
function ybbsize(theSmilie){
var text=prompt("请输入文字", "");
if(text){
document.kbbs.body.value += '[size=' + theSmilie + ']'+ text + '[/size]';
}
}
function image() {
var FoundErrors = '';
var enterURL   = prompt("请输入图片地址","http://");
if (!enterURL) {
FoundErrors +="\n";
}
if (FoundErrors) {
return;
}
var ToAdd = "[IMG]"+enterURL+"[/IMG]";
document.kbbs.body.value+=ToAdd;
document.kbbs.body.focus();
}
function fly() {
fontbegin="[fly]";
fontend="[/fly]";
fontchuli();
}
function move() {
fontbegin="[move]";
fontend="[/move]";
fontchuli();
}
function center() {
fontbegin="[align=center]";
fontend="[/align]";
fontchuli();
}
function light() {
fontbegin="[glow=255,yellow,2]";
fontend="[/glow]";
fontchuli();
}
function grade() {
var ToAdd = "[showtograde=1]内容[/s]";
document.kbbs.body.value+=ToAdd;
document.kbbs.body.focus();
}
function name() {
var ToAdd = "[showtoname=zym]内容[/s]";
document.kbbs.body.value+=ToAdd;
document.kbbs.body.focus();
}
function reply() {
var ToAdd = "[showtoreply]内容[/s]";
document.kbbs.body.value+=ToAdd;
document.kbbs.body.focus();
}
function ying() {
fontbegin="[SHADOW=255,yellow,1]";
fontend="[/shadow]";
fontchuli();
}
function smoney() {
var ToAdd = "[smoney=1000]内容[/s]";
document.kbbs.body.value+=ToAdd;
document.kbbs.body.focus();
}
function smeili() {
var ToAdd = "[smeili=1000]内容[/s]";
document.kbbs.body.value+=ToAdd;
document.kbbs.body.focus();
}
function sjingyan() {
var ToAdd = "[sjingyan=1000]内容[/s]";
document.kbbs.body.value+=ToAdd;
document.kbbs.body.focus();
}