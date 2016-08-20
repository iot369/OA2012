<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<%if session("user")<>"" then%>
<html>
<head>
<title>聊天输入区</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
body      {  font-size: 9pt; border: 0; margin: 0px}
td        {  font-size: 9pt}
select    {  font-size: 9pt}
.normal   {  <%=m_button%> ; font-size: 9pt}
.over     {  <%=m_buttonover%> ; font-size: 9pt}
.down     {  <%=m_buttondown%> ; font-size: 9pt}
-->
</style>
<script Language="JavaScript">
<!--
 var conti=0;
 var contit=0;
 var max=20;
 var whamsg=new Array(20+1);
 var base=0;
 var p=0;
 var j;
 var tt=0;
 for (j=0;j<=max+1;j++)
  whamsg[j]='';
  
 function addOne(what){
  if (base<max+1) {
   whamsg[base]=what;
    base++;
  } else {
   for (i=0;i<max;i++)
    whamsg[i]=whamsg[i+1];
   whamsg[i]=what;
  }
  p=base;
 }
 
 function goPrev(theForm){
  if (p>0) p--;
   theForm.saystemp.value=whamsg[p]
  theForm.saystemp.focus();
 }
 function goNext(theForm){
  if (p<base) p++;
   theForm.saystemp.value=whamsg[p]
  theForm.saystemp.focus(); 
 }

function htmlstyle(list)
{ 	var listText;
	var n;
   	if((list.selectedIndex != -1)&&(list.selectedIndex != 0))
   	{ 	
   		listText = list.options[list.selectedIndex].value; 
		n=listText.indexOf('?');
		if(document.forms.formtalk.saystemp.value){
	   		if(n>=0) document.forms.formtalk.saystemp.value=listText.substring(0,n)+document.forms.formtalk.saystemp.value+listText.substring(n+1,listText.length);
	   	}else{
			document.forms.formtalk.saystemp.value=listText.substring(0,n)+'  '+listText.substring(n+1,listText.length);
		}
   	}
}

function bs() { document.forms.formtalk.saystemp.style.color = document.forms.formtalk.addwordcolor.value;
  document.forms.formtalk.saystemp.focus();
}
	function receiver(list)
	{ var listText='';
	   if((list.selectedIndex != -1)&&(list.selectedIndex != 0))
	   { listText = list.options[list.selectedIndex].value; }
	  document.forms.formtalk.saystemp.value=listText;
	}
	
function reld () {
parent.frm_content.location.href='content.asp<%if r_refresh="down" then%>#dian<%end if%>';
}

function checksays(theForm) {
if ((document.forms.formtalk.towho.options[document.forms[0].towho.selectedIndex].value.length !=0) && (document.forms.formtalk.says.value==document.forms.formtalk.saystemp.value) && (document.forms.formtalk.towho.options[document.forms[0].towho.selectedIndex].value==document.forms.formtalk.towho1.value)) {
<%if session("userlevel")<9 then%>
alert('好罗嗦啊，请不要对'+document.forms.formtalk.towho.options[document.forms[0].towho.selectedIndex].value+'总说这一句呀！');
return false;
<%else%>
return true;
<%end if%> }

  if (checkempty(document.forms.formtalk.saystemp.value)==false) {
     alert('请输入发言');
     return false;
  }

Today = new Date();
var NowHour = Today.getHours();
var NowMinute = Today.getMinutes();
var NowSecond = Today.getSeconds();
var continuum = (NowHour*3600)+(NowMinute*60)+NowSecond;
if((continuum-conti)>1)
{ conti=continuum;}
else if(contit<4){contit++;alert('有话慢慢说啊！别噎着 ^!^');return false;}
else {conti=0; contit=0; return true;}

document.forms.formtalk.job.value='says';
var hua=document.forms.formtalk.saystemp.value;
var dongzuo=hua.substring(0,1);
if(dongzuo=="/") {

	var s1=hua.indexOf(" ");
	var s2=s1+1;
	var dongzuo1=hua.substring(0,s1);
	var s3=hua.length;
	
	var un =hua.substring(s2,hua.length);
	if ((checkempty(un)==true) && (s1!=-1)) {
	hua=dongzuo1;
	document.forms[0].towho.options[document.forms[0].towho.selectedIndex].text=un;
	document.forms[0].towho.options[document.forms[0].towho.selectedIndex].value=un;
	}

	if(hua=="/ver") {alert('Info Chat 3.0'); return false;}

        else if(hua=="/") {alert('/后紧跟指令符'); return false;}
	document.forms.formtalk.saystemp.value=hua;
	document.forms.formtalk.job.value="action";	
}
	
  document.forms.formtalk.saysnoc.value++;
  document.forms.formtalk.says.value=document.forms.formtalk.saystemp.value;
  document.forms.formtalk.towho1.value=document.forms.formtalk.towho.options[document.forms[0].towho.selectedIndex].value;
  addOne(theForm.saystemp.value);
  if (document.forms.formtalk.autocls.checked==true) {
  document.forms.formtalk.saystemp.value='';
  }
  self.status='发言已送出';
  return true;
}
function checkempty(txtval){
  if (txtval.length==0)
     return false;
  var i=0;
  while (i<txtval.length) {
     if (txtval.substring(i,i+1) != '')
        return true;
     i=i+1;
  }
  return false;
}
<%if session("userlevel")<9 then%>
function click() {
if (event.button==2) {
var char_=navigator.appName
if(char_=="Netscape"){
window.parent.resizeTo(3,1)
window.parent.moveTo(100,30)
}
else{
window.parent.resizeTo(4,1)
window.parent.moveTo(100,30)
}
alert("请不要按右键，现在请在任务栏最大化聊天室的窗口，谢谢");
}
}
document.onmousedown=click
<%end if%>
if (self==top)
{top.location="index.htm";}
//-->
</script>
</head>
<body bgcolor="#cccccc" text="<%=m_text2%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
  <form method="POST" name="formtalk" target="frm_content" action="content.asp<%if r_refresh="down" then%>#dian<%end if%>" onsubmit="return(checksays(this));">
    <tr> 
      <td height="1" colspan="3" bgcolor="#8C001A"></td>
    </tr>
    <tr> 
      <td height="7" colspan="3"></td>
    </tr>
    <tr>
	  <td align="right" valign="top" width="10%"> 
        <div align="left">
          <table width="50%" border="1" align="right" bordercolor="#ffffff" bgcolor="#ffffff">
            <tr>
              <td bordercolor="#000000" bgcolor="#ffff00" align="center"> 
                <div align="center">
                  <SCRIPT language=JavaScript>
<!--  
function defaul_home()
{
this.home.style.behavior='url(#default#homepage)';this.home.setHomePage('http://www.nbjoe.com');
}

function openmusicbox()
{ window.open("http://www.nbjoe.com/musicbox/musicbox.htm","musicbox","width=500,height=150,status=no"); }

//-->
</SCRIPT>
                  
          </table>
          <div align="right"><br>
            </a><br>
            </div>
        </div>
      </td> 
      <td align="right" valign="top" width="100%"> 
        <input type=hidden name="job" value="says">
        <input type=hidden name="says" value="">
        <input type=hidden name="towho1" value="">
        <input type=hidden name="lastno" value="">
        <input type=hidden name="saysnoc" value="0">
        <input type=hidden name="filteruser" value="">
        <%
        set rs=my_conn.execute("select * from "&dbtable_gbook&" where "&dbfield_gbook_name&"='"& session("user") &"'")
        session("havebook")=0
        if not rs.eof then
        %>
        <Script language="JavaScript">
          <!--
            window.open("letter.asp","letter","width=420,height=160,resizable=0,status=no,scrollbars=yes");
          -->
        </Script>
        <%
        end if
        rs.close
        set rs=nothing
        my_conn.close
        set my_conn=nothing
        %>
        <select onchange="bs();">
          <option value="" selected>您的资料</option>
          <option value="">积分:<%=session("oldrate")%></option>
          <option value="">等级:<%=session("userlevel")%></option>
          <option value=""> 
          <% if session("userlevel")>7 then response.write "管理员" %>
          </option>
        </select>
        <input type="button" value="<<" name="cmdPrev" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="goPrev(document.forms.formtalk);">
        <input type="button" value=">>" name="cmdNext" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="goNext(document.forms.formtalk);">
        <input type="text" name="saystemp" id="saystemp" size="40" maxlength="100">
        <input type="submit" name="sub" value="送出" onclick="bs();" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
        <input type="reset" value="清除" onclick="bs();" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
        &nbsp; <br>
        <input type="button" value="刷新上区" class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';" onClick="confirm('刷新功能可以在谈话区长时间不动或者\n谈话内容太多影响机器速度的时候使用，\n刷一下就会正常了。');reld();bs();" name="button">
        <%if session("userlevel")>3 then%>
        <input type="button" value="/指令集" class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';" onClick="alert('右边将会显示动作列表!');top.frm_onlineuser.location.href='listfunction.asp'">
        <%end if%>
        <%if session("userlevel")>6 then%>
        <select name="htmlit" onChange="htmlstyle(this);bs();">
          <option value="" selected>HTML格式 
          <option value="<font size=4 face=楷体_GB2312>?</font>">楷体4号 
          <option value="<font size=5 face=楷体_GB2312>?</font>">楷体5号 
          <option value="<font size=6 face=楷体_GB2312>?</font>">楷体6号 
          <option value="<font size=4 face=宋体>?</font>">宋体4号 
          <option value="<font size=5 face=宋体>?</font>">宋体5号 
          <option value="<font size=6 face=宋体>?</font>">宋体6号 
          <option value="<i>?</i>">字体倾斜 
          <option value="?">字体加粗 
          <option value="<img src=?>">图片 
          <option value="<marquee>?</marquee>">对象移动 
        </select>
        <%end if%>
        <select name="addwordcolor" id="addwordcolor" onChange="bs();">
          <%
        randomize
        displaycolor=int(rnd(1)*5)
        %>
          <option style="color: #000000" value="#000000">绝对黑色</option>
          <option style="color: #000088" value="000088">深蓝忧郁</option>
          <option style="color: #0000ff" value="0000ff">草原之蓝</option>
          <option style="color: #008800" value="008800">橄榄树绿</option>
          <option style="color: #008888" value="008888">灰蓝种族</option>
          <option style="color: #0088ff" value="0088ff">海洋之蓝</option>
          <%if displaycolor=1 then%>
          <option style="color: #00a010" value="00a010" selected>绿色回忆</option>
          <%else%>
          <option style="color: #00a010" value="00a010">绿色回忆</option>
          <%end if%>
          <option style="color: #1100ff" value="1100ff">蓝色月光</option>
          <option style="color: #111111" value="111111">夜幕低垂</option>
          <option style="color: #333333" value="333333">灰色轨迹</option>
          <option style="color: #50b000" value="50b000">春草青青</option>
          <%if displaycolor=2 then%>
          <option style="color: #880000" value="880000" selected>暗夜兴奋</option>
          <%else%>
          <option style="color: #880000" value="880000">暗夜兴奋</option>
          <%end if%>
          <option style="color: #8800ff" value="8800ff">发亮蓝紫</option>
          <option style="color: #888800" value="888800">卡其制服</option>
          <option style="color: #888888" value="888888">伦敦灰雾</option>
          <option style="color: #8888ff" value="8888ff">兴份过蓝</option>
          <%if displaycolor=3 then%>
          <option style="color: #aa00cc" value="aa00cc" selected>紫的拘谨</option>
          <%else%>
          <option style="color: #aa00cc" value="aa00cc">紫的拘谨</option>
          <%end if%>
          <option style="color: #aaaa00" value="aaaa00">流金岁月</option>
          <option style="color: #ccaa00" value="ccaa00">卡布其诺</option>
          <option style="color: #ff0000" value="ff0000">正宗喜红</option>
          <%if displaycolor=4 then%>
          <option style="color: #ff0088" value="ff0088" selected>爱的暗示</option>
          <%else%>
          <option style="color: #ff0088" value="ff0088">爱的暗示</option>
          <%end if%>
          <option style="color: #ff00ff" value="ff00ff">红的发紫</option>
          <option style="color: #ff8800" value="ff8800">黄金岁月</option>
          <option style="color: #ff0005" value="ff0005">红袍飘飘</option>
          <option style="color: #ff88ff" value="ff88ff">紫金绣帖</option>
          <option style="color: #ee0005" value="ee0005">焚心似火</option>
          <option style="color: #ee01ff" value="ee01ff">红粉佳人</option>
          <option style="color: #3388aa" value="3388aa">我不知道</option>
        </select>
        <select name="towho" onChange="bs();">
          <option value="大家" selected>大家 
          <% for i=1 to 100
             if application("user"&i) <> "" then %>
          <option value="<%=application("user"&i)%>"><%=application("user"&i)%> 
          <% end if
             next %>
        </select>
        <%if session("userlevel")>0 then%>
        <input type="checkbox" name="autocls" value="yes" checked onClick="document.forms.formtalk.autocls.blur(); bs();">
        自动清除 
        <%end if%>
        <input type="checkbox" name="towhoway" value="yes" onClick="document.forms.formtalk.towhoway.blur(); bs();">
        私聊 &nbsp; <br>
        <%if session("userlevel")>4 then%>
        <select name="fs" onChange="bs();">
          <option value="" selected>字效 
          <option value="1">加粗 
          <option value="2">斜体 
          <option value="3">小字 
          <option value="4">大字 
          <option value="5">左移 
          <option value="6">右移 
          <option value="7">来回 
          <option value="8">下线 
        </select>
        <%end if%>
        <select name="addsays" onChange="bs();">
          <option value="" selected>选择表情 
          <option value="抖着腿，">抖腿 
          <option value="平静的">平静 
          <option value="惊慌失措的">慌乱 
          <option value="很高兴的">高兴 
          <option value="潇洒的">潇洒 
          <option value="微微笑的">微微笑 
          <option value="惋惜的">惋惜 
          <option value="深情的">深情 
          <option value="哈!哈!哈!的笑着">大笑叁声 
          <option value="笑嘻嘻的">嘻笑 
          <option value="幸灾乐祸的">幸灾乐祸 
          <option value="眉飞色舞的">眉飞色舞 
          <option value="阴阳怪气的">阴阳怪气 
          <option value="哭天抢地的">哭天抢地 
          <option value="牢骚满腹的">牢骚满腹 
          <option value="理屈词穷的">理屈词穷 
          <option value="江郎才尽的">江郎才尽 
          <option value="凶神恶煞的">凶神恶煞 
          <option value="眉歪眼斜的">眉歪眼斜 
          <option value="轻轻牵着对方的手，小声的">牵手 
          <option value="一把甩掉对方的手，恶恨恨的">甩手 
          <option value="满脸俱是轻蔑不屑之色的">轻蔑 
          <option value="目光凛然的">目光凛然 
          <option value="声嘶力竭的">声嘶力竭 
          <option value="意乱情迷的">意乱情迷 
          <option value="脸色铁青的">脸色铁青 
          <option value="失魂落魄的">失魂落魄 
          <option value="怨毒的">怨毒 
          <option value="目光呆滞的">目光呆滞 
          <option value="柔声的">柔声 
          <option value="眼波晶莹明亮的">眼波动人 
          <option value="目光尽赤，怒不可遏的">怒不可遏 
          <option value="娇笑起来的">娇笑 
          <option value="茫然的">茫然 
          <option value="给了对方一耳光，大骂着">打耳光 
          <option value="奸笑着">奸笑 
          <option value="奴才嘴脸十足的">扮奴才像 
          <option value="毛手毛脚的">毛手毛脚 
          <option value="得意洋洋的">得意 
          <option value="一脸无赖泼皮样子的">耍无赖 
          <option value="嘟着嘴的">嘟嘴 
          <option value="抛去一个媚眼，然后">抛媚眼 
          <option value="一把将对方抱住，哭哭涕涕的">抱住对方哭 
          <option value="快要哭的">快要哭 
          <option value="拳打脚踢的">拳打脚踢 
          <option value="不怀好意的">不怀好意 
          <option value="二球劲十足的">二楞子 
          <option value="遗憾的">遗憾 
          <option value="瞪大了眼睛,很讶异的">讶异 
          <option value="幸福的">幸福 
          <option value="翻箱倒柜的">翻箱倒柜 
          <option value="悲伤的">悲伤 
          <option value="淫淫笑的">淫淫笑 
          <option value="流口水的">流口水 
          <option value="正气凛然的">正气凛然 
          <option value="生气的">生气 
          <option value="大声的">大声 
          <option value="傻乎乎的">傻乎乎 
          <option value="一付很满足的">很满足 
          <option value="手足无措的">手足无措 
          <option value="很无辜的">很无辜 
          <option value="喃喃自语的">喃喃自语 
          <option value="恶狠狠的瞪着眼，">瞪眼 
          <option value="快要吐的">想吐 
          <option value="不舒服的">不舒服 
          <option value="无精打采的">无精打采 
          <option value="依依不舍的">依依不舍 
          <option value="吐白沫着">白沫 
          <option value="掩饰不住狂喜的心情，">狂喜 
          <option value="笑呵呵一拱手，">拱手 
          <option value="很有礼貌地作了一揖,">作揖 
          <option value="慨叹万千的">慨叹 
          <option value="很不好意思地">致歉 
          <option value="望着窗外细雨淅淅，不由得双眼朦胧，">伤感 
          <option value="含泪要哭的">含泪 
          <option value="想到伤心处，泪流如注的">大哭 
          <option value="抱头放声大哭，">痛哭 
          <option value="好害怕的">害怕 
          <option value="感到很是奇特，">奇特 
          <option value="眯着小眼睛的">眯眼 
          <option value="咯咯一笑，很大方的">大方 
          <option value="脸上泛起了红晕，">脸红 
          <option value="气愤的">生气 
          <option value="提高嗓门，">大声 
          <option value="运足气一声，">断喝 
          <option value="一脸的迷茫的">迷茫 
          <option value="无奈地耸耸肩，">耸肩 
          <option value="使劲敲敲自己脑门，">拍脑 
          <option value="看着别人谈笑，无聊的很的">无聊 
          <option value="顾作沉思状，">沉思 
          <option value="一付无辜的样子，">无辜 
          <option value="不舒服的，">不适 
        </select>
        <%if session("userlevel")>1 then%>
        <select name="addsign" onChange="bs();">
          <option value="" selected>选择图案 
          <option value="A">胜利 
          <option value="B">OK 
          <option value="C">赞 
          <option value="D">逊 
          <option value="I">布 
          <option value="J">笑脸 
          <option value="L">扁嘴 
          <option value="M">炸弹 
          <option value="N">毒药 
          <option value="O">旗 
          <option value="R">太阳 
          <option value="S">水滴 
          <option value="T">雪花 
          <option value="V">十字架 
          <option value="Y">星星 
          <option value="Z">月亮 
          <option value="b">蚵蚪 
          <option value="d">肚脐 
          <option value="f">箭头 
          <option value="m">圆 
          <option value="r">四角 
          <option value="1">档案夹 
          <option value="4">备忘录 
          <option value="6">沙漏 
          <option value="8">滑鼠 
          <option value="!">笔 
          <option value="#">剪刀 
          <option value="$">眼镜 
          <option value="%">铃当 
          <option value="&">书 
          <option value="*">明信片 
          <option value="(">电话 
          <option value="-">信箱 
          <option value="<">磁片 
          <option value="|">黑花 
          <option value="]">舵 
          <option value="{">白花 
          <option value="'">腊烛 
          <option value=":">电脑 
        </select>
        <%end if%>
        <select name="emote" onchange="receiver(this);bs();">
          <option value="" selected>选择趣语 
          <option value="翘起一对修长的双腿，筛糠似的抖了起来...">抖腿 
          <option value="抱拳团团一拜道:&quot;敝人对各位的景仰之情，有如涛涛江水连绵不绝。&quot;">讨好 
          <option value="抓了抓头皮，露出迷惑的神情.... ">迷惑 
          <option value="一脸的晦气，好象谁牵走了他的牛 ">晦气 
          <option value="举双脚赞成">同意 
          <option value="象断线的风筝一样徒劳地挣扎了两下，然后无可救药的落进了对方的温柔陷阱……">被温柔包围 
          <option value="突然大喊一声，这是我的第三千八百次初恋，求求你，接受我的这一小片痴情吧！">求爱 
          <option value="强忍住烦躁和一刹那的烈火焚心，从房间里冲到花园，果断地纵身跳进冰凉的游泳池中，大家惊讶的发现，不一会儿，诺大一个游泳池的水居然全沸腾了！">男人的冲动 
          <option value="紧紧咬住自己的嘴唇，眼睛象水滴一样生动地注视着对方，用可以摧毁一个国家，杀死一个国王，揉碎一颗心的声音说道，云想衣裳，花想容，我又怎么会不想呢。">女人的冲动 
          <option value="刷的抽出一把闪闪发光的戒刀，一刀将自己的大好头颅切了下来，嘴里还不停的嘀咕，和我比酷，谁能比我酷？">比酷 
          <option value="突然拔出一把手枪，指着自己的头，泪眼婆娑的看着大家。">想自杀 
          <option value="狂笑三声，从身后抽出一柄斧头，大喊着，天下没有公理了，让我来伸张正义！">要砍人 
          <option value="浑身酒气挣扎着踉跄走了两步，突然不留神，一个跟头直直插进了路边的水沟里。">喝醉了 
          <option value="在血盘大口里吐了吐青白的舌头! ">吐舌头 
          <option value="重重地拍了一下脑袋, 终于想到了! ">想到了 
          <option value="作出生气的表情!火冒三丈～想要打人～ ">生气 
          <option value="使劲地拍巴掌,啪啪啪啪啪.....拍得手都红起来了! ">拍巴掌 
          <option value="脸都红了,恨没有地洞, 好钻进去躲起来 :~ ">脸红 
          <option value="象朵含羞草一样躲了起来，又突然跑回来，轻轻地、快快地吻了对方一下，转身就跑了。">羞答答地吻 
          <option value="象暴风雨一样扑了过来，嘴里说着，不能没有你，……可惜，脚下一滑，一嘴吻在对方的鞋帮子上。">狂吻对方 
          <option value="向在场的人道别: 走咯! 我会想你们的">道别1 
          <option value="如果你们记得我，别忘记给我留言哦！">道别2 
          <option value="飘啊飘的一段情，有雨也有风……蓦然回首你仍在，浪漫红尘中……我走了……">道别3 
          <option value="凄婉地说道：“世上没有不散的宴席，我先走一步了，大家保重。” ">道别2 
          <option value="向大家说 BYE BYE! 千山我独行... 不必相送.">道别3 
          <option value="把胸脯拍得噼啪响：“武林中拳头大的说话，有种的上来比划比划！” ">拍胸脯 
          <option value="两眼翻白，腿儿蹬了几下，脑袋一歪，死了。 ">扮死 
          <option value="想啊想，结果把脑袋给想破了，白白的脑浆流了一地! ">想破头 
          <option value="被吓得双腿直哆嗦. ">害怕 
          <option value="口吐白沫，昏倒在地 ">昏倒 
          <option value="盈盈一拜,笑笑地说:各位大哥,小女子有礼了.. ">问好(girl) 
          <option value="大咧咧地说:大姐姐,小姐姐!本公子这厢有礼啦!&quot;">问好(boy) 
          <option value="环顾四周，全聊天室山河一片大好，于是清了清喉咙：“同志们好。” ">问好1 
          <option value="感动极了，两行热泪夺眶而出。 ">感动 
          <option value="一把眼泪一把鼻涕地喊：冤枉啊!! ">喊冤 
          <option value="自言自语道:“我,先天下人之忧而忧, 后天下人之乐而乐...这个这个好象不太妥” ">自言自语 
          <option value="大叫一声：“风紧，扯呼”,撒开腿就跑，片刻就消失在一片茫茫白雾中 ">逃跑 
          <option value="兴奋地唱：“我是一只小鸭子，伊呀伊呀哟！” ">兴奋地唱 
          <option value="深情地唱道：“我早已为你种下，九百九十九朵玫瑰。”">深情地唱 
          <option value="咚一声! 坐在地上哇啦大哭~ ">大哭 
          <option value="笑得象花儿一样 ">笑 
          <option value="露出像白痴般的笑容....:O~ ">傻笑 
          <option value="向天狂笑：“普天之下，竟然没有我的对手...”">狂笑 
          <option value="深情的凝视前方 ">凝视 
          <option value="躲在一边一声不吭，独个儿大口大口地喝着闷酒。">喝闷酒 
          <option value="低着头，玩弄着自己的手指，羞得满脸通红">害羞 
          <option value="正疯狂的亲吻着镜子中的自己...">自吻 
          <option value="丢几罐冰凉的啤酒给大家，然后说:「干啦 !」">敬酒 
          <option value="像发了疯似的一个人在聊天室中跳起舞来">跳舞 
          <option value="躲在聊天室的角落偷偷哭泣">哭泣 
          <option value="嘴里念念有辞的,好像正在背三字经.">念经 
          <option value="对自己的长像极度不满意, 决定向双亲提出抗议.">自卑 
          <option value=" 大喊加油~ 加油~">喊加油 
          <option value="像疯狗似的见人就吠. 嗷~呜~">发狗疯 
          <option value="大声叫了一声: 宾果!">宾果 
          <option value="牙齿又在痒...想找跟狗骨头咬咬.">牙痒 
          <option value="像只小猫似的正在凌虐刚从聊天室抓到的蟑螂.">玩蟑螂 
          <option value="拿起一杯香香醇醇的咖啡咖啡喝了一口. 嗯, 好香!">喝咖啡 
          <option value="认为自己一定是头壳坏掉了,才会有这样的言行.">头坏 
          <option value="觉得自己的狗腿功夫已经到了炉火纯青的地步,只差还没出国比赛拿金牌而已.">狗腿 
          <option value="最近一直觉得自己头晕目眩,恶心想吐,不知是不是.....">头晕 
          <option value="害怕地发起抖来，感觉背脊发凉, 头皮发麻~">发抖 
          <option value="一直觉得自已不笨,只是凡人不能理解自己想法而已.">清高 
          <option value="四处张望了一下,然后对着凳子深情款款地吻了一下.">吻凳 
          <option value="不知想到了什么,一个人在那吃吃的傻笑.">傻笑 
          <option value="嘿嘿嘿....地奸笑了几声，八成想到什么坏事头上。">奸笑 
          <option value="荡漾着无限温情，淡描着依稀的梦景。泡好了茶，你可知道，我在这儿等你......">等你 
          <option value="等待着你带着笑容走来，等待着你用心的关怀……">梦里飞花 
          <option value="躲在墙脚瑟瑟发抖，嘴里喃喃自语“我不好，我检讨，我不对，我有罪">检讨 
          <option value="默默地走了，风中飘散的长发和孤寂的背影在斜阳下显得分外潇洒">默默地走了 
          <option value="默默地走了，风中飘散的长发和孤寂的背影在斜阳下显得分外潇洒">默默地走了 
          <option value="弯腰拱手，淡淡一笑“久闻踏雪园人杰地灵，小弟初来乍到，还望诸位提点提点">多多关照 
          <option value="大叫一声:“扯呼”，油门一加，意大利版的NSR车头提起，屁股冒烟，消失在一片茫茫白雾中">兜风 
          <option value="从腰间抽出一支玉箫，吹出优雅的箫声。箫声入耳，众人激荡的心情便即平复，随即和道：“苍生笑，不再寂聊，豪情仍在痴痴笑笑">优雅的箫声 
          <option value="坐在墙脚边，孤独地唱:“寂寞难奈，寂寞难奈.....">寂寞难奈 
          <option value="遥望着在场的众人，微微一笑。一双妙目间情意流动，顾盼生姿">顾盼生姿 
          <option value="左思又想，自己无才无德，无钱无貌，怎么偏偏天下女子都痴缠我一人">爱上我 
          <option value="往脸上涂脂抹粉，描眉毛，抹口红，夹睫毛，打扮得花枝招展，越看自己越像林青霞">花枝招展 
          <option value="从烟缸中拣出个烟屁股，用两根手指夹着点着了，眯起眼睛狂嘬几口">抽烟 
          <option value="累了，直感到头重脚轻，眼看就要倒在地上">累了 
          <option value="郎声说道：拜山拜水拜码头，在下初到宝地，还请各位老大们多多关照">初来驾到 
          <option value="微蹙娥眉，幽幽一叹“但见泪痕湿，不知心恨谁">微蹙娥眉 
          <option value="手舞木棒，对准的后脑一记闷棍，哼也没哼一声，软软的摔倒在地">手舞木棒 
          <option value="看着默默远走的背影，看着她在风中飘散的长发和那孤寂的背影，泪水止不住滑落到胸口，心中呼喊着“其实我真的爱你”">真的爱你 
          <option value="看着离去的背影，不禁黯然神伤。留声机里飘出邓丽君的歌声：“今宵离别后，何日君再来">何日君再来 
          <option value="拉着你的手，温柔地默默对视。目光中孕含着千种柔情，万般蜜意">默默对视 
          <option value="轻轻地搂着心上人指着天上的月亮说:“今晚的月亮是我们的证人”">搂着心上人 
          <option value="勇敢的跪了下来:“你愿意嫁给我吗”">求婚 
          <option value="左手拼命挥动着小旗，右手拿着扩音器，语音哽咽地喊到“欢迎，欢迎，热烈欢迎">热烈欢迎 
          <option value="小鸟依人般地靠在宽阔的胸膛上，娇羞地说道：“我真幸运得到了你”">小鸟依人 
        </select>
        <%if session("userlevel")>2 then%>
        <select name="addpic" onChange="bs();">
          <option value="" selected>选择贴图 
          <option value="smile.gif">笑 
          <option value="frown.gif">皱眉 
          <option value="redface.gif">尴尬 
          <option value="biggrin.gif">大笑 
          <option value="wink.gif">挤眼 
          <option value="tongue.gif">吐舌 
          <option value="shy.gif">腼腆 
          <option value="sleepy.gif">瞌睡 
          <option value="sunglasses.gif">带墨镜 
          <option value="supergrin.gif">露齿笑 
          <option value="mad.gif">生气 
          <option value="cry.gif">哭 
          <option value="scared.gif">害怕 
          <option value="love.gif">爱 
        </select>
        <%end if%>
        &nbsp; </td>
      <td align="right" valign="top">&nbsp;</td>
      
    </tr>
    <tr> 
      <td height="7" colspan="3"></td>
    </tr>
    <tr> 
      <td height="1" colspan="3" bgcolor="#8C001A"></td>
    </tr>
  </form>
</table>
  </body>
</html>
<%end if%>