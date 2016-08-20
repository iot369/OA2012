<%Response.Expires=0%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<meta http-equiv=refresh content='10;url=f1.asp#bottom'>

<script language="JavaScript">
<!--
function selectwho(list){
parent.f2.document.forms[0].towho.options[0].value=list;
parent.f2.document.forms[0].towho.options[0].text=list;
parent.f2.document.forms[0].saystemp.focus();
}
//-->
</script>

<style type="TEXT/CSS"> 
<!--
body,table {color:#000000;font-family: 宋体_GB2312; font-size: 9pt; line-height: 12pt}
A:link {text-decoration: none; color:#E00000; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:visited {text-decoration: none; color: #E00000; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:active {text-decoration: underline; color: #E00000; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:hover {text-decoration: none; color: F00000; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
-->
</style>

</head>
<%If IsEmpty(Session("UserName")) then Response.Write("与系统失去联系，请重新登录。"):Response.End%>
<%
emote="//? 抓了抓头皮，露出迷惑的神情.... //?? 抓呀抓，把头皮都抓破了，也没有想出个所以然来。 //:) 脸上露出了一丝笑容 //:( 一脸的晦气，好象谁牵走了他的牛 //:p 在血盘大口里吐了吐青白的舌头！ //:D 耶...又一个反肚乌龟耶...:DDD //@@ 眨眨眼！一对会说话的眼睛闪闪动人 //8d8d 腮上飞红，秋波闪动，动情地说：所有人让我们交个朋友吧! //aah 重重地拍了一下脑袋，终於想到了！ //addoil 使劲给自已加油！ //admire 抱拳团团一拜道：“敝人对各位的景仰之情，有如涛涛江水连绵不绝。” //agree 完全同意。 //agree2 觉得就自己最对，一贯正确，别人都是扯蛋... //ah 惊讶地“啊！”了一声。 //ahh 重重地拍了一下自己的脑袋，“我怎麽没想到？！” //angry 作出生气的表情！火冒三丈～想要打人～ //ahh 撅了撅嘴说“我生气了！” //ann 大吼一声：天帝来了！把在场的所有人都吓得发抖 //applaud 啪啪啪啪啪啪啪…… //away 无从参与，唯有告辞。 //bad “完了，我要做个堕落天使，到地狱去找哈哈，我的最爱！” //bad2 坏坏，你欺负人！ //bb 越来越年轻 //bbb 很想邀请人对聊啦~ //b_hi 大咧咧地说：大姐姐，小姐姐！本公子这厢有礼啦！ //bear 往洞里一躺，嘟咙到：“熊熊我要冬眠了，不要打搅我！” //beaut 觉得自己挺臭美 //bicycle 推出了一辆‘吱吱’作响的自行车，梦想骑上它在本站一游。 //bigfool “我是成奎安我怕谁！” //birthday 逢人就打招呼：大家好，今天是我的生日！请吃糖！ 请吃糖！ //bite 磨了磨牙 ... //bless 默默地祝福着... //blush 脸都红了，恨没有地洞，好钻进去躲起来~~ //boring 觉得话题沉闷，只有坐在一旁发呆... //bow 一鞠躬，二鞠躬，三鞠躬.... //boy 双眼一翻：“大丈夫可杀不可辱……” //brag 把胸脯拍得噼啪响：“武林中拳头大的说话，有种的上来比划比划！” //brag2 双眼一翻：“大丈夫可杀不可辱……” //brag3 大吼一声：“老子今天就是死，也要拉几个垫背的！！！” //breath 感觉呼吸越来越... //brother 弯弯腰，道了个万福“各位大哥，有礼了！” //buddha 在这一瞬间，觉得自己简直跟神一样。 //bug 嘴角一撇，狞笑道：“我是害虫我怕谁？！” //bye 回眸一笑，一切尽在不言中 //byeall 向在场的人道别：走咯！ 我会想你们的。 //cat “喵～！ 喵～～！” //care 孤独寂寞地躲在角落痛哭 ~~>_<~~ //caress 觉得自已好可怜... //caxie 以一个优雅的动作拨了下头发... //clap “啪啪啪啪啪啪啪啪！！” //club 晃动着手中的木棒，到处寻找目标。 //coffin 长叹了一口气，那神情仿佛是在说：“我都是要进棺材的人了，还在乎些什么？？！” //cold “阿——秋！”打了个喷嚏。 //cold2 听了，冷得直起鸡皮疙瘩.... //cometo 突然张开眼睛，站起来了。“唉。阎君大人不肯收留我，只好回来了。” //comfort 安慰自己说，“那有什么关系！” //corpse 两眼翻白，腿儿蹬了几下，脑袋一歪，死了。 //cool 大叫起来：“哇塞～～～好cool哦～～” //cough 咳了几声 //crazy 仰天狂笑，看上去很疯狂 //cry 咚一声！ 坐在地上哇啦大哭～～ //cry2 想到伤心处，忍不住放声大哭 //cu 一一抱拳，道：“一路顺风～～ ” //cu2 看着远处的背影，感到无比的失落 ... //cut 一刀把自己的脑袋切了下来，提在手里... //dan 坏了，看到老板正在向这逼近。快退出吧！ //dance 快乐得手舞足蹈！挺自得其乐的嘛！ //ding 摇头晃脑地吟：“千江有水...这个...千江月...啊...” //die 要死了，好难受呀，呜呜呜呜呜呜呜呜，好想活下去呀... //drivel 垂涎不已，忍不住淌了一地口水 :O~~！觊觎那DOVE很久了... //drinkup 躲在一边一声不吭，独个儿大口大口地喝着闷酒。 //dog “汪～！汪汪～！” //doubt 想啊想，结果把脑袋给想破了，白白的脑浆流了一地！ //dove 咦???我的DOVE呢? 嘿！！你们谁拿了我的DOVE... //ds 使出吃奶的力气叫喊“杀人啦~ 救命啊~” //ds2 比划了一个恶狠狠的姿势：“你还敢来？” //ds3 哼了一下：“笑什么笑，有什么好笑的？！” //duh 用力地拍了一下自己的脑袋，大声地说：“对呀！我怎麽没想到？！” //eat 肚子饿死了，呜，虽然不好吃，还是快去吃吧 //ecstasy 蹲在地上，捧腹大笑！！！ //eye 楚楚动人的明眸早已说明一切 //en “嗯”的一声... //face 顽皮地做了个鬼脸 ^_* //faint 口吐白沫，昏倒在地。 //fat 照照镜子：唉！我怎么越来越猪了？ //fear 被吓得双腿直哆嗦。 //fear2 缩著身说：“好怕好怕哦！:p” //flook 深情的凝视前方 //fly 口吐白沫，两眼翻白，向后倒飞了出去。 //fool 露出像白痴般的笑容....:O~ //foolish 觉得自己真笨，这么简单的道理都不懂 //forgive 内心经过一场痛苦的思考，总算原谅了自已。 //frankie 掏出番茄，大咬了几口。“味道好极了！” //friend 很高兴认识大家唷！！！ ^_^ //giggle 发出一串母鸡下蛋般的笑声。 //go 紧紧抱着自己的钱袋，走路，吃饭，连做梦都乐得笑出声儿来。 //go_eat 肚子饿死了，忽想起了食堂，呜，虽然不好吃，还是快去吃吧。 //g_hi 盈盈一拜，笑笑地说：各位大哥，小女子有礼了... //goodbye 凄婉地说道：“世上没有不散的宴席，我先走一步了，大家保重。” //goodbye2 看着离去的背影，不禁黯然神伤：“今宵离别后，何日君再来！” //grin 邪恶地笑了一笑，又有什么坏点子了！ //girl 盈盈一拜，笑笑地说：“各位大哥，小女子有礼了...” //ha 很得意摆出胜利的姿势！「 哈哈哈...」 //hand 很热情地握手 //haha 蹲在地上，捧腹大笑！！ //haha2 向天狂笑：“普天之下，竟然没有我的对手...” //hammer 晃动着手中好大的木棒，到处寻找目标... //happy r-o-O-m....听了真爽！ //happy2 作揖回礼，笑呵呵地说道：大家好，恭喜发财！ //hehe 呵呵地傻笑了几声…… //headshake 摇摇头，一副很无辜的样子。 //heihei 满脸一堆奸笑 //hello 愉快的和大家打招呼 //help 大叫：救命啊！ 救命啊！ //heng 哼了一声，一脸不懈的样子。 //heng2 “哼哼！” “哼！！”地两声…… //hero 声嘶力竭地大声叫道：哈！砍头不过头点地，老子十八年後又是一条好汉！ //hi 说道：Hi～ //hit 弹了弹衣上的灰尘道：如此稀松平常的功夫，也敢闯江湖！ //hmm 「嗯」的一声，一副欲言又止的样子。 //howl 对著天空狼嚎，ㄠ～ㄨ～～，ㄠ～ㄨ～～。 //hot 好热，好热....我要 icecream…… //hp 一副很无助的样子…… //hug 轻轻的拥著自己，显露出孤寂的表情。 //hug2 觉得世界真是美好，每样东西都值得拥抱。 //humble 拱了拱手道：“过奖，过奖！” //idiot 突然觉得自己从没像现在这麽愚蠢过。 //idle 无聊地坐在地上，开始发起呆来了。 //ill 说道：我生病了 //inn 感到十分委曲，流下泪来。 //itnsr 大叫一声：“扯呼”，油门一加，意大利版的NSR车头提起，屁股冒烟，消失在一片茫茫白雾中…… //jianbo 说：“我要过沙漠，快备交通工具。” //joe 喃喃而语“你说什么我不懂耶……” //jump 说：“别象个兔子，跳来跳去的。” //kgb 对大家嚎道：快拿把刀杀了我吧！呜~~ //kick 一阵的狂踢 //kill “我心情不好，别理我，让我去死吧！” //kiss 抛了下飞吻 ~ ~ //laugh 开怀大笑…… //lag 慢慢地挥动双手... 哇... 怎麽这麽慢！ //lazy 觉得整个人懒懒的，提不起精神来。 //lean 感到那么的温柔 //lick 舔了舔嘴边的口水。 //lie 道：俺除了实话，什么都说了。 //life 哇！搞出人命啦！！ //lonely 坐在墙脚边，孤独地唱：“寂寞难奈，寂寞难奈……” //look 四下张望了一阵，发现一个人都没有…… //lookme 两眼一翻：看什么看?再看把你眼珠子给挖出来 ！ //love 心想：喔，多么迷人耶，我爱上了！ //lovesee 微微一笑！一双妙目间情意流动，顾盼生姿！ //loveu 楚楚动人的明眸早已说明一切 //luck 哇！福气啦！ //lure 满脸的媚态，作出一副迷人的姿态诱惑所有在场的人。 //marry 泪眼朦胧地，不知道怎么把求婚的话说出来。 //microant 面包会有的，蚂蚁会偷吃的，结果一切都是没有的。 //miss 好挂念心中的人儿，什么事也干不下去了…… //mistake 说道：“又搞错了！” //mm 不知道想到什么诱人的东西，口水都流出来了…… //mmm 欲言又止... //new 新人新猪肉，不要宰割啊！ //newyear 衷心地祝愿：祝各位新年快乐，心想事成，财源广进…… //nod 点了点头 //no 摇了摇头，“那怎么可以呢，不可以的，不可能的啦。” //now 叹了口气，说：“算了吧……” //oh 恍然大悟的样子：“喔...原来如此，我知道啦。” //out 道：风紧，扯呼！！ //paste 拿出一张狗皮膏药，在小炉上细细地煨热后，“啪”地捂住了自己的嘴巴！ //pat 摸了摸自已的脑袋 //pinch 很用力的拧了自己一下，看看是不是做梦.... //pig 哼哼叽叽地说，我比猪还慢，要么大伙拱我算了。 //pk “这次真黑！”话音刚落，一下扑倒在街上。 //poe 你拔剑长吟道：“十年磨一剑，霜寒未曾试。今日把君问，谁有不平事？” //poke 无聊地伸出手指，前后左右瞎捅了一气，发现还是没有人理他。 //pout 不高兴的翘起嘴来 //ppp 轻轻擦擦自己的脸…… //praise 你伸出指头把大家挨个点了一遍：“高手，都是高手！” //puke 觉得一阵恶心，哇啦哇啦地吐了满地…… //puke2 真恶心，我听了都想吐！快去找垃圾桶吐…… //punch 抱住自已的头往豆腐里撞…… //qiang 嘿嘿，小心，我有枪…… //raise 拼命地伸长自己的手臂，高声叫道：“我，我，我！” //rabbit 说：“别象个兔子，跳来跳去的。” //regard 叹了口气道：“你当我是谁呀，我也是苦出身！” //revenge 突然觉得这世界上已经没有什么比报仇更加重要的事了。 //right 对自己的所作所为很满意…… //rob 一声大喊：“此山是我开，此树是我栽，若要从此过，留下买路财！” //rob2 掏出大刀。“要你的钱，还要你的命！” //rose 深情地唱道：“我早已为你种下，九百九十九朵玫瑰。” //sad 越来越沮丧…… //sad2 十分伤心的流下泪来。 //seeyou 走了，前路风风雨雨，各位江湖路上多珍重啊！ //shake 摇了摇头 //shake2 把头摇得象拨浪鼓一样 //shi 小声“嘘”的一声，“大佬来了！！！” //shiver 因寒冷而发抖。 //shy 脸红得像熟透的苹果 *^o^* //shout 咬牙切齿地对着天空大叫：“贼老天！” //shrug 无奈地耸耸肩。 //sigh “唉”叹了口气，不知道哪里不对了…… //sing 快乐唱起歌来“我悄悄地蒙上你的眼睛，让你猜猜我是谁？” //sit 四处张望，找张椅子坐下。 //sit2 坐怀不乱 //slap 一巴掌打在空气里 //slow 拼命敲打键盘，口里嚷到：“天哪，我变慢了！” //slow2 慢慢地挥动双手... 哇... 怎麽这麽慢！ //slogan 高举右拳，咬牙切齿地高呼：“打倒一切恶势力！” //sleep ZZZZzzZZzzzzzZ，真无聊，都快睡著了 //sleep2 伸个长长的懒腰，嘟着小嘴：“我好困，要觉觉猪了。” //smile 愉快地微笑著。 //smoke 从烟缸中拣出个烟屁股，用两根手指夹着点着了，眯起眼睛狂嘬几口！ //smooch 拥吻著…… //smug 得意洋洋地翘起了二郎腿，哼着小调，心里打着如意算盘。 //snicker 在旁边偷偷地笑。 //so 就酱子！！ //so2 摇摇食指，「小朋友，这样不可以喔！」 //sob 难过地抽泣起来。 //sorry 对所有在场的人表示歉意…… //spider 接着就把一窝毛绒绒，黑呼呼的小蜘蛛扔了出来…… //ssmile 露出甜蜜的微笑 //stare 眼中露出很奇怪的光…… //strut 大摇大摆地走起路来都有风 //stretch 舒舒服服的伸伸懒腰～ //suck 伤心失望之余，真想买块臭豆腐撞死，摸摸口袋却发现身边没有零钱。 //sun 非常高兴地说：“边度都有阳光！” //sweat 急得快哭了 //smileyue 笑了笑，象月儿一样迷人…… //tape 你所说的话，将会成为呈堂证供。哼哼...小心啦！ //tea 坐下来，喝口茶，吃个包。 //tear 开始掉眼泪了…… //tear 一把眼泪一把鼻涕地喊：冤枉啊！！ //thank 向在场的所有人表示衷心的感谢！ //these 道：这个，我觉得，也太那个了吧？ //think 歪著头想了一下 //think2 “这个问题，嗯……我想想看……” //timeout =隐退一会！ //tired 感到自已很累了 //touch 感动极了，两行热泪夺眶而出。 //touch2 舒舒服服的伸伸懒腰～ //tu 耳朵一搭、兔牙一撇，“哼”的一声决定不理人了！ //violet “嘘！老板娘来了，大家小心。” //violet2 “老板娘，这酒没下蒙汗药吧......啊......又死了！” //visit 郎声说道：拜山拜水拜码头，在下初到宝地，还请各位老大们多多关照！ //visit 道：哪里哪里，岂敢岂敢，有事尽管分付！ //wait 觉得不能再呆了…… //wake 揉揉眼睛，清醒了过来…… //wave 拼命的摇手 //who 嚷道：“你当你是谁呀，你以为你是张学友呀？” //wine 躲在一边一声不吭，独个儿大口大口地喝着闷酒。 //wing 在BBS里自由自在地飞翔 //wk 我会悟空的瞬间移动！你打不着，打不着... //worry 自言自语道：“我，先天下人之忧而忧，后天下人之乐而乐...这个这个好象不太妥” //work 一天三顿饱两个倒，也不知到自己忙个啥…… //wrong 心想：难道我错了吗？ //wrong2 心想：好在红心挂到墙上了。 //wrong3 心想：好在我会丐帮的打狗棒法，来多一点更好，可以做狗肉煲。 //xinku 环顾四周，全站山河一片大好，于是清了清喉咙：“同志们好。” //xinku2 雄赳赳气昂昂地说道：“首长好！” //xinku3 的手在空中挥了挥：“同志们辛苦了！” //xinku4 挺起胸膛，扯着嗓子喊道：“为人民服务！” //xixi 嘻皮笑脸 //ya 好得意的样子 //ykiss 再亲一个，“啵！” //yy 说：“臭丫头，还不过来帮我捶捶！打死你！” //yy2 兴奋地唱：“我是一只小鸭子，伊呀伊呀哟！” //zap 剑眉一轩，冷冷的瞥了一眼，背转身淡淡说道：你，不是我的，对手！ //zzz 揉了揉双眼，打了个呵欠, 好困呀…… //ask 请问： //chant 歌颂： //cheer 喝采： //chuckle 轻笑： //curse 咒骂： //demand 要求： //frown 蹙眉： //groan 呻吟： //grumble 发牢骚： //hum 喃喃自语： //moan 悲叹： //notice 注意： //order 命令： //ponder 沈思： //pout 撅起像樱桃般的小嘴说： //pray 祈祷： //request 恳求： //shout 大叫： //sing 唱歌： //smile 微笑地说： //swear 发誓： //smirk 假笑： //sob 不停地哭哭啼啼道： //tease 嘲笑： //whimper 呜咽的说： //yawn 哈欠连天： //yell 大喊： //"

emote2="//? 很疑惑的看着对象... //:( 一肚子苦水向对象倒了出来 :((((((((((((((( //:) 对对象露出愉快的表情 //:)2 对着对象笑了笑 //:p 对对象惊讶得吐出了舌头 ... //@@ 大大的眼睛， 天真的望著对象... (@_@) //addoil 憋足了劲大喊：对象！加油！加油！下面就是宝贝了！ //agree 完全同意对象的看法。 //ah 对着对象惊讶地「啊！」了一声。 //angry 脸臭臭的， 一副懒得理对象的模样。 //allen 向对象用力将鸡蛋扔去。“啪！”好爽！ //bad 噪着对象“坏坏，欺负人！！” //bb 抱着对象轻轻摇晃，“小宝宝，食蛋糕。” //bbb BB...BBB...唉...又要去复机啦！等我哦！对象 //beaut 指着对象的鼻子说，看把你美的，整个一大傻冒儿.... //bite 张开血盆大口，狠狠地咬了下去， 把对象咬的哇哇大叫。 //blush 对着对象说：羞羞脸！ //bearhug 热情的拥抱对象 //birthday 祝对象生日快乐， 献花！ //bless 祝福对象心想事成 //bow 毕恭毕敬的向对象弯腰鞠躬 //boy 大咧咧地对象说“大姐姐，小姐姐！本公子这厢有礼啦！” //breath 赶快给对象做人工呼吸！ //bt 哇！对象真...真...真... ，你这变态！ //bug 大叫“对象，你这条臭虫！” ... //bye 对对象说道：再见！ //care 轻轻抚摸对象 //caress 抚摸对象 //cat 靠在对象的耳朵旁边“喵～！ 喵～～” //caxie 拿出一块破抹布，一脸妩媚地给对象擦鞋... //clap 向对象热烈鼓掌 //cold 听了对象的话，冷得直起鸡皮疙瘩.... //comfort 温言安慰对象 //comfort1 安慰对象说“面包会有的，牛奶也会有的，老婆会有很多的。” //cool 哈哈大笑，对着对象拱了拱手道：壮士过奖了！ //crazy 恶毒的眼神看着对象... //cringe 向对象卑躬屈膝，摇尾乞怜 //cry 越想越伤心，不禁趴在对象的肩膀上嚎啕大哭起来。 //cu 对着对象抱拳道“青山不改，绿水常流，咱们后会有期！” //cu2 轻轻吻了对象一下，低声说：“我走了，有缘的，我们会再见。” //cu3 望着对象离去的背影渐渐消失，两滴晶莹的泪花悄悄从腮边滑落。 //cut 扬起牛角解腕尖刀， 三两下就把对象剁成了许多小块， 放在太阳底下晒干。 //dance 拉起对象的小手跳起舞来， 两人好陶醉的样子！ //date 对对象妩媚一笑：“你？？想约我去街？” //die 对对象说：“嚷什么嚷？这样的小贱人，死了活该！” //dog “落闸，放狗！”把对象咬得七零八落。 //dogleg 对对象狗腿 //dove 给了对象一块DOVE，说：“呐，给你一块DOVE，要好好听话哦！” //drivel 对著对象流口水 //ds 恶狠狠的冲着对象喝道：“你再笑，想我用嘴巴堵你是不是？” //ds2 掐着对象的脖子，恶狠狠的叫道：“看你说不说！” //ds3 极富表情， 一个大喷嚏以每秒三米的速度向对象打了过去！ //eat 说道“开饭啦！对象饭饭去哦。” //en 望了望对象，似乎想说什么... //face 对着对象做了个鬼鬼脸，*&* //faceless 对着对象大叫道：“嘿嘿，面子卖多少钱一斤？” //faint 突然昏倒在对象的怀抱中 ... //fat 施展展催肥大法，不一会儿，对象就胖得连路都走不动了。 //fear 对对象说：「怕了吧！ 哈哈哈！」 //finger 对对象伸出一个指头说：“No，No，不是这样的耶” //flook 痴痴的望着对象，那深情的眼神说明了一切。 //fly 向对象飞了过去... //fool 对着对象说：“这可是一个笨问题哦！” //forgive 内心经过一场痛苦的思考， 总算原谅了对象 //giggle 对著对象傻傻的呆笑 //girl 皱了皱眉， 看着对象唱道：“这个女人哪~ 啊~ 啊~ 不寻~常！” //go 不怀好意地紧盯着对象的口袋，一边笑嘻嘻地踱过去答讪。 //go2 说：对象你趴下，我上！ //goodbye 对着对象凄婉地说道：“世上没有不散的宴席， 请保重！” //goodbye2 向对象一拱手， 朗声长笑：千山我独行，何劳相送，有缘后会自有期！ //grin 对对象露出邪恶的笑容， 大家瞎子吃饺子>>>>>>肚里有数 ... //greet 愉快地向对象打招呼 //grow 对对象咆哮不已 //gzxj 夹起一只小蘑菇，放到对象的嘴里，喊着：“来，别客气，吃！” //ha 皮笑肉不笑地对对象打了个哈哈 //haha 冲着对象哈哈大笑... //hammer 举起惠香的50000000T铁锤往对象上用力一敲！，*** 『 锵 ！』 ***. //hand 跟对象握手 //handpat 一巴掌打在对象的屁屁上，把他震得头皮发麻 //happy 打了个大揖， 对着对象笑道：大家发财， 大家发财！ //he 对对象傻笑几声 //hehe 对对象呵呵地傻笑了几声 //heihei 对对象「嘿嘿嘿....」地奸笑了几声。 //hello 对对象很有礼貌地说了一声：“Hello！ 你好！” //heng 对着对象“哼！”地冷笑一声 ... //hi 对对象很有礼貌地说了一声：“Hi！ 你好！” //hit 一拳打在对象的肚皮上， 正中红心， 爽啊！ //hp 请求得到对象的帮助... //hug 轻轻地拥抱对象 //hug2 紧紧地拥抱著对象 //ice 从北极唤来一阵凛冽的寒风，把对象吹成一个冰雕。 //ill 皱着眉头对对象说“有点不太舒服了哦” //inn 的眼中充满泪水， 无辜的望著对象... //jump 高兴地跳入对象的怀里 //joe 对着对象说：“你真是---不~~知~~所~~谓！” //kick 一脚踢在对象的屁屁上，印出一个清楚的鞋印。 //kiss 啵！偷偷亲了对象一下！ //kiss2 在对象的额头上吻了一下 //kiss3 轻轻地吻着对象的脸颊…… //kill 开始认真考虑杀死对象的可能性。 //koxia 面色一黑，提起特大号的篮球，向对象扣下去，扣得他口吐白沫直呼：爽啊爽啊！ //laugh 张大眼睛地瞪着对象，慢慢地咧开嘴，捧腹大笑起来…… //lean 小猫猫般地依偎在对象的怀里 //lly 捧着对象月亮般的脸蛋，吧唧吧唧地亲了N口！ //look 贼贼地看着对象，不知道在打什么馊主意。 //look2 瞪着对象，不知想说什么…… //look3 不怀好意地看了对象一眼。 //love 心想：喔，对象是多么迷人，我想我是爱上你了。 //love2 对对象深情地说：“在天愿作'妈'公仔，在地愿为'油炸鬼'！” //love3 含情脉脉对着对象，说出了世界上最感人的三个字！！！…… //loveu 轻轻地搂着对象指着天上的月亮说：“今晚的月亮是我们的证人。” //lovesee 迷人的眼眸对对象眨了眨眼~ //lonely 觉得话题沉闷，只有和对象坐在一旁发呆... //lovelook 一双水汪汪的大眼睛含情脉脉地看着对象 //lure 摆出撩人的姿态诱惑对象... //marry 单腿跪下，一脸深情地向对象求婚。 //milk 给对象倒一大杯热热的牛奶：“休息一会吧，趁热喝。” //miss 甜甜一笑，眼中却流下眼泪：“对象，真的是你吗？” //mm 色迷迷的对对象说：“美眉好呀！嗯，我们好象哪里见过？” //who 道：提起对象，那真是谁人不知，哪个不晓？！ //no 对对象摇摇头说道：我不知道... //no2 对着对象：“你说得很清楚， 我听得很模糊...” //nod 向对象点头称是... //now 对对象道：“讲正经事，言归正传吧！” //nudge 用手肘顶着对象的肥肚子 //paste 拿出一张狗皮膏药，在小炉上细细地煨热后，啪地捂住了对象的嘴巴！ //pat 轻轻地拍了拍对象的头。 //pig 不耐烦地对对象哼叽， 你怎么比猪还慢啊？ 要么大伙拱你算了... //pinch 用力的将对象拧得黑一块，青一块，红一块... //poke 用手指无聊地捅了捅对象 //poke2 冲着对象道：捅什么捅，再捅剁了你的手指！ //puke 突然脸上一阵青一阵白， 唏哩哗啦地吐了对象满身... //ppp 用脸颊轻轻地磨擦著对象的粉脸， 悄声说道：我好喜欢你哦... //punch 狠狠揍了对象一顿！ //qian 很牛逼的掏出一把五四手枪，“砰”的一声就把对象放倒了！ //qsister 对对象说道“做我的姐姐好吗？” //qsister1 对对象说道：“做我的妹妹好吗？” //qbrother 对对象说道：“做我的弟弟好吗？” //qbrother1 对对象说道：“做我的哥哥好吗？” //right 对对象说：说的对！ //rose 突然从身后拿出一朵玟瑰， 深情地献给对象。 //sad 撩起胳膊上的衣服，对对象说，老子的伤疤比你多，悲伤什么！ //shake 对著对象摇了摇头... //shy 面对着对象，脸好热，好热... //sigh 觉得万念俱灰，对着对象不由得“唉”的叹了口气。 //sing 对对象含情默默的唱起情歌来！“对你爱爱爱不完！” //sit 四处张望， 找张椅子，和对象坐在一起 //slap 左手抓着对象的衣领，右手运掌如风，噼噼啪啪的打了十几个耳光。 //sleep 对对象唱起了摇篮曲…… //slow 对着的对象破口大骂：“你看我这死电脑， 这破线路！” //smile 愉快地对对象微笑着 //smooch 看看四下无人， 与对象热情拥吻著 //smoke 和对象一道腾云驾雾，飘飘欲仙之中... //sorry 很不好意思的向对象赔礼道歉！ //sorry2 请求对象的原谅... //sorry3 内心经过一场痛苦的思考， 总算原谅了对象 //so 对对象说：这件事就这么定了！！ //spider 像蜘蛛精八只有力的脚一样紧紧地缠住对象 //stare 用很奇怪的眼神瞄著对象 //stw 拉着对象的小手“别走嘛，别走哦，再陪陪俺哦....” //sweat 替对象擦擦汗 //tea 给对象敬上热茶，还有包子... //tear 看着对象，难过得要哭了。 //thank 满脸诚意的说：“谢谢你了对象，你真好！” //think 想啊想，想对象，结果把脑袋给想破了！ //touch 轻轻地抚摸对象的脸， 眼中充满爱怜…… //tu 拎着对象的长耳朵一把扔出了会议室！ //tired 对对象道：我真的好累了！ //visit 对对象道：拜山拜水拜码头， 在下初到宝地， 还请各位老大们多多关照！ //wait 手头有事儿， 隐退一会， sorry！... 对象 //wake 试著把对象摇醒， 大叫：“猪！起来啦！” //wave 向着对象挥了挥手。 //welcome 欢迎欢迎，热烈欢迎对象的到来！ //welcome2 哪里哪里，对象你太客气了！ //wine 眯着眼睛对对象说：“来来来， 喝了这一杯再说吧！” //work 对着对象叹了口气道：“当我是谁呀，我也是苦出身！” //wrong 对对象说道：你错了！ //xixi 疯吻对象 //yeah 你得意的对对象作出胜利的手势！ 「 V 」说：「 哈哈哈...」 //ykiss 在对象的嘴角上轻轻的吻了一下。 //zap 把对象拖到广场中...从天上召来一道闪电，把他化为灰烬。 //zzz 白了对象一眼，说：“无聊不无聊啊？” //znw 哈哈哈哈，哈哈哈哈，对着对象天使般地哈哈狂笑…… //ask 请问： //chant 歌颂： //cheer 喝采： //chuckle 轻笑： //curse 咒骂： //demand 要求： //frown 蹙眉： //groan 呻吟： //grumble 发牢骚： //hum 喃喃自语： //moan 悲叹： //notice 注意： //order 命令： //ponder 沈思： //pout 撅起像樱桃般的小嘴说： //pray 祈祷： //request 恳求： //shout 大叫： //sing 唱歌： //smile 微笑地说： //swear 发誓： //smirk 假笑： //sob 不停地哭哭啼啼道： //tease 嘲笑： //whimper 呜咽的说： //yawn 哈欠连天： //yell 大喊： //"

sayscolor=Request.Form("sayscolor")

addwordcolor=Request.Form("addwordcolor")
addsays=Request.Form("addsays")
Select Case addwordcolor
Case "0"
addwordcolor="#008888"
Case "1"
addwordcolor="#000000"
Case "2"
addwordcolor="#0088FF"
Case "3"
addwordcolor="#0000FF"
Case "4"
addwordcolor="#000088"
Case "5"
addwordcolor="#888800"
Case "6"
addwordcolor="#008888"
Case "7"
addwordcolor="#008800"
Case "8"
addwordcolor="#8888FF"
Case "9"
addwordcolor="#AA00CC"
Case "10"
addwordcolor="#8800FF"
Case "11"
addwordcolor="#888888"
Case "12"
addwordcolor="#CCAA00"
Case "13"
addwordcolor="#FF8800"
Case "14"
addwordcolor="#FF0088"
Case "15"
addwordcolor="#FF00FF"
Case "16"
addwordcolor="#FF0000"
Case else
addwordcolor="#008888"
End Select
Select Case sayscolor
Case "0"
sayscolor="#660099"
Case "1"
sayscolor="#000000"
Case "2"
sayscolor="#0088FF"
Case "3"
sayscolor="#0000FF"
Case "4"
sayscolor="#000088"
Case "5"
sayscolor="#888800"
Case "6"
sayscolor="#008888"
Case "7"
sayscolor="#008800"
Case "8"
sayscolor="#8888FF"
Case "9"
sayscolor="#AA00CC"
Case "10"
sayscolor="#8800FF"
Case "11"
sayscolor="#888888"
Case "12"
sayscolor="#CCAA00"
Case "13"
sayscolor="#FF8800"
Case "14"
sayscolor="#FF0088"
Case "15"
sayscolor="#FF00FF"
Case "16"
sayscolor="#FF0000"
Case else
sayscolor="#660099"
End Select
Select Case addsays
Case "0"
addsays="对"
Case "1"
addsays="微微笑对"
Case "2"
addsays="温柔地对"
Case "3"
addsays="红着脸对"
Case "4"
addsays="摇头晃脑得意地对"
Case "5"
addsays="哈！哈！哈！笑着对"
Case "6"
addsays="神秘兮兮地对"
Case "7"
addsays="战战兢兢地对"
Case "8"
addsays="毛手毛脚地对"
Case "9"
addsays="嘟着嘴地对"
Case "10"
addsays="慢条斯理地对"
Case "11"
addsays="同情地对"
Case "12"
addsays="幸灾乐祸地"
Case "13"
addsays="快要哭地对"
Case "14"
addsays="哭着对"
Case "15"
addsays="拳打脚踢地对"
Case "16"
addsays="不怀好意地对"
Case "17"
addsays="遗憾地对"
Case "18"
addsays="瞪大了眼睛，很诧异地对"
Case "19"
addsays="幸福地对"
Case "20"
addsays="翻箱倒柜地对"
Case "21"
addsays="悲痛地"
Case "22"
addsays="正义凛然地对"
Case "23"
addsays="严肃地对"
Case "24"
addsays="生气地对"
Case "25"
addsays="大声地对"
Case "26"
addsays="傻乎乎地对"
Case "27"
addsays="很满足地对"
Case "28"
addsays="手足无措地对"
Case "29"
addsays="很无辜地对"
Case "30"
addsays="喃喃自语地对"
Case "31"
addsays="恶狠狠地瞪着眼对"
Case "32"
addsays="快要吐地对"
Case "33"
addsays="无精打采地对"
Case "34"
addsays="依依不舍地对"
Case "35"
addsays="口吐白沫对"
Case else
addsays="对"
End Select
%>

<body bgcolor="#F0F0FF" text="#660099">

<%
'人员登录

OUN=Application(SESSION("CRNAME")&"OUN")
OULT=Application(SESSION("CRNAME")&"OULT")
usernum=Application(SESSION("CRNAME")&"usernum")
cur=Application(SESSION("CRNAME")&"cur")
whotowho=Application(SESSION("CRNAME")&"whotowho")
sentences=Application(SESSION("CRNAME")&"sentences")
UPDA=0
for i=1 to 60
'更新自已最后访问时间(OULT)
If Session("username")=OUN(i) then
UPDA=1
OULT(i)=Now
End If
If len(OUN(i))=0 then usernum=i-1:Exit For
Next

'加入新会议用户
If UPDA=0 then
OUN(usernum+1)=Session("username")
OULT(usernum+1)=Now
usernum=usernum+1
cur=cur+1
if cur>60 then cur=1
sentences(cur)="<font color=#FF0000>[公告]</font>"&Session("username")&"刚刚进入<u>"&Session("CRNAME")&"</u>……<font color=#B0B0B0>("&Now&")</font>"
whotowho(cur,1)="System"
whotowho(cur,2)="大家"
End If
Application.Lock
Application(SESSION("CRNAME")&"OUN")=OUN
Application(SESSION("CRNAME")&"OULT")=OULT
Application(SESSION("CRNAME")&"sentences")=sentences
Application(SESSION("CRNAME")&"whotowho")=whotowho
Application(SESSION("CRNAME")&"cur")=cur
Application(SESSION("CRNAME")&"usernum")=usernum
Application.UnLock

%>


<%
sentences=Application(SESSION("CRNAME")&"sentences")
whotowho=Application(SESSION("CRNAME")&"whotowho")
cur=Application(SESSION("CRNAME")&"cur")
says=ltrim((Request("says")))
If len(says)>0 then

cur=cur+1
If cur>60 then cur=1

whotowho(cur,1)=Session("UserName")

If Request("toone")="ON" then
whotowho(cur,2)=Request("towho")
senhead="<font color=#FF0000>[悄悄话]</font>"
else
whotowho(cur,2)="大家"
senhead=""
End If

sentences(cur)=senhead&"<font color="&addwordcolor&">"&Session("UserName")&"</font>"&addsays&Request("towho")&"说：<font color="&sayscolor&">"&Server.HtmlEncode(says)&"</font><font color=#B0B0B0>("&Now&")</font>"


'---emote Beg---
If left(says,2)="//" then

myemote=Lcase(rtrim(left(says,Instr(says+" "," "))))
othersays=rtrim(right(says,len(says+" ")-Instr(says+" "," ")))

emoloc=instr(emote,myemote+" ")
If emoloc>0 then
emosay=mid(emote,emoloc+len(myemote),(instr(emoloc+len(myemote),emote,"//"))-(emoloc+len(myemote)))
emosay=Replace(emosay,"对象",Request("towho"))
sentences(cur)=senhead&"<font color="&addwordcolor&">"&Session("UserName")&"</font>"&emosay&" "&othersays&"<font color=#B0B0B0>("&Now&")</font>"
End If

If emoloc=0 or Request("towho")<>"大家" then
emoloc=instr(emote2,myemote+" ")
If emoloc>0 then
emosay=mid(emote2,emoloc+len(myemote),(instr(emoloc+len(myemote),emote2,"//"))-(emoloc+len(myemote)))
emosay=Replace(emosay,"对象",Request("towho"))
sentences(cur)=senhead&"<font color="&addwordcolor&">"&Session("UserName")&"</font>"&emosay&" "&othersays&"<font color=#B0B0B0>("&Now&")</font>"
End If
End If

End If
'---End emote---

End If
Application.Lock
Application(SESSION("CRNAME")&"sentences")=sentences
Application(SESSION("CRNAME")&"whotowho")=whotowho
Application(SESSION("CRNAME")&"cur")=cur
Application.UnLock


sentences=Application(SESSION("CRNAME")&"sentences")
whotowho=Application(SESSION("CRNAME")&"whotowho")
cur=Application(SESSION("CRNAME")&"cur")

outputstr=Empty

for i=cur+1 to 60
If len(sentences(i))>0 then
If whotowho(i,2)="大家" then
OutputStr=OutputStr&(sentences(i)&"<br>")
else
if whotowho(i,1)=Session("username") or whotowho(i,2)=Session("username") then
OutputStr=OutputStr&(sentences(i)&"<br>")
End If
End If

End If
next


for i=1 to cur
If len(sentences(i))>0 then

If whotowho(i,2)="大家" then
OutputStr=OutputStr& (sentences(i)&"<br>")
else
if whotowho(i,1)=Session("username") or whotowho(i,2)=Session("username") then
OutputStr=OutputStr& (sentences(i)&"<br>")
End If
End If

End If
next

Response.Write(outputStr)

%>
<a name="bottom"> </a>
</body>
</html>
