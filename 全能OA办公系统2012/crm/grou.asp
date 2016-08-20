
<html>
<head>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<STYLE type=text/css>A:link {
	COLOR: #003366; TEXT-DECORATION: none
}
A:visited {
	COLOR: #003366; TEXT-DECORATION: none
}
A:hover {
	TEXT-DECORATION: underline
}
BODY {
	FONT-SIZE: 12px; SCROLLBAR-ARROW-COLOR: #dde3ec; SCROLLBAR-BASE-COLOR: #f8f9fc; BACKGROUND-COLOR: #e9edf7
}
TABLE {
	FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana
}
TEXTAREA {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
INPUT {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
OBJECT {
	FONT-WEIGHT: normal; FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana; BACKGROUND-COLOR: #f8f9fc
}
SELECT {
	FONT-WEIGHT: normal; FONT-SIZE: 11px; COLOR: #000000; FONT-FAMILY: Tahoma; BACKGROUND-COLOR: #f8f9fc
}
.nav {
	FONT-WEIGHT: bold; FONT-SIZE: 12px; FONT-FAMILY: Tahoma, Verdana
}
.header {
	FONT-WEIGHT: bold; FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/default/headerbg.gif); COLOR: #ffffff; FONT-FAMILY: Tahoma, Verdana
}
.category {
	FONT-SIZE: 12px; BACKGROUND-IMAGE: url(images/default/catbg.gif); COLOR: #000000; FONT-FAMILY: Tahoma
}
.multi {
	FONT-SIZE: 11px; COLOR: #003366; FONT-FAMILY: Tahoma
}
.smalltxt {
	FONT-SIZE: 11px; FONT-FAMILY: Tahoma
}
.mediumtxt {
	FONT-SIZE: 12px; COLOR: #000000; FONT-FAMILY: Tahoma, Verdana
}
.bold {
	FONT-WEIGHT: bold
}
BLOCKQUOTE {
	BORDER-RIGHT: #dde3ec 1px dashed; PADDING-RIGHT: 5px; BORDER-TOP: #dde3ec 1px dashed; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 20px; BORDER-LEFT: #dde3ec 1px dashed; MARGIN-RIGHT: 20px; PADDING-TOP: 5px; BORDER-BOTTOM: #dde3ec 1px dashed; BACKGROUND-COLOR: #ffffff
}
.code {
	PADDING-RIGHT: 5px; PADDING-LEFT: 5px; PADDING-BOTTOM: 5px; MARGIN-LEFT: 20px; MARGIN-RIGHT: 20px; PADDING-TOP: 5px; BACKGROUND-COLOR: #ffffff
}
</STYLE>

</head>
<SCRIPT LANGUAGE="JavaScript">
<!-- Begin

// Compare two options within a list by VALUES

function compareOptionValues(a, b) 

{ 

  // Radix 10: for numeric values

  // Radix 36: for alphanumeric values

  var sA = parseInt( a.value, 36 );  

  var sB = parseInt( b.value, 36 );  

  return sA - sB;

}



// Compare two options within a list by TEXT

function compareOptionText(a, b) 

{ 

  // Radix 10: for numeric values

  // Radix 36: for alphanumeric values

  var sA = parseInt( a.text, 36 );  

  var sB = parseInt( b.text, 36 );  

  return sA - sB;

}



// Dual list move function

function moveDualList( srcList, destList, moveAll ) 

{

  // Do nothing if nothing is selected

  if (  ( srcList.selectedIndex == -1 ) && ( moveAll == false )   )

  {

    return;

  }



  newDestList = new Array( destList.options.length );



  var len = 0;



  for( len = 0; len < destList.options.length; len++ ) 

  {

    if ( destList.options[ len ] != null )

    {

      newDestList[ len ] = new Option( destList.options[ len ].text, destList.options[ len ].value, destList.options[ len ].defaultSelected, destList.options[ len ].selected );

    }

  }



  for( var i = 0; i < srcList.options.length; i++ ) 

  { 

    if ( srcList.options[i] != null && ( srcList.options[i].selected == true || moveAll ) )

    {

       // Statements to perform if option is selected



       // Incorporate into new list

       newDestList[ len ] = new Option( srcList.options[i].text, srcList.options[i].value, srcList.options[i].defaultSelected, srcList.options[i].selected );

       len++;

    }

  }



  // Sort out the new destination list

  newDestList.sort( compareOptionValues );   // BY VALUES

  //newDestList.sort( compareOptionText );   // BY TEXT



  // Populate the destination with the items from the new array

  for ( var j = 0; j < newDestList.length; j++ ) 

  {

    if ( newDestList[ j ] != null )

    {

      destList.options[ j ] = newDestList[ j ];

    }

  }



  // Erase source list selected elements

  for( var i = srcList.options.length - 1; i >= 0; i-- ) 

  { 

    if ( srcList.options[i] != null && ( srcList.options[i].selected == true || moveAll ) )

    {

       // Erase Source

       //srcList.options[i].value = "";

       //srcList.options[i].text  = "";

       srcList.options[i]       = null;

    }

  }



} // End of moveDualList()



//  End -->

</script>
<script>
	function getPValue(srcList)
	{
		var reValue="";
		for( var i = srcList.options.length - 1; i >= 0; i-- ) 
		{ 
		    	if ( srcList.options[i] != null)
		    	{
				reValue=reValue + "," + srcList.options[i].value;
		    	}
	  	}
	  	if(reValue.indexOf(",")==0){
	  		reValue=reValue.substr(1,reValue.length-1)
	  	}
	  	document.GG.client_id.value=reValue;
	  	document.GG.submit();
	}
</script>
  <table border="1" bordercolordark=#FFFFFF bordercolorlight=#000000 cellpadding="0" cellspacing="0" align="center">
  <form name="GG" action="groupMember.asp" method="post">
    <tr height="22" class="tablehead" align="center"> 
      <td width="25%">选择客户</td>
      <td width="25%">操作</td>
      <td width="25%">已选择客户</td>
    </tr>
    <tr align="center"> 
      <td width="25%"> 
      <select multiple size="25" style="width:100" name="listLeft" class="input">
      
		<option value="93">合肥腾飞</option>
	
		<option value="94">wqwqwqw</option>
	
		<option value="95">dadadad</option>
	
		<option value="96">sadafafda</option>
	
		<option value="97">重庆桐君阁股份有限公司医药批发分公司</option>
	
		<option value="98">广州番禺</option>
	
		<option value="99">s </option>
	
		<option value="101">京城条码公司</option>
	
		<option value="102">yuy</option>
	
		<option value="103">李卫强</option>
	
		<option value="104">成都公司</option>
	
		<option value="105">werfgwsf</option>
	
		<option value="106">dfdfdf</option>
	
	</select>
        
      </td>
      <td width="25%">
      	     <input  name="Add     &gt;&gt;" type="button" style="width:90" onClick="moveDualList( this.form.listLeft,  this.form.listRight, false )"  value="右移   &gt;&gt;" class="button"> 
        <br> <input  name="Add     &lt;&lt;" type="button" style="width:90" onClick="moveDualList( this.form.listRight, this.form.listLeft,  false )"  value="左移     &lt;&lt;" class="button"> 
        <br> <input  name="Add All &gt;&gt;" type="button" style="width:90" onClick="moveDualList( this.form.listLeft,  this.form.listRight, true  )"  value="全部 &gt;&gt;" class="button"> 
        <br> <input  name="Add All &lt;&lt;" type="button" style="width:90" onClick="moveDualList( this.form.listRight, this.form.listLeft,  true  )"  value="全部 &lt;&lt;" class="button">

      </td>
      <td width="25%">
	      <select multiple size="25" style="width:100" name="listRight" class="input">
	      	
			<option value="105">werfgwsf</option>
		
			<option value="104">成都公司</option>
		
			<option value="103">李卫强</option>
		
			<option value="102">yuy</option>
		
			<option value="101">京城条码公司</option>
		
			<option value="99">s </option>
		
			<option value="98">广州番禺</option>
		
			<option value="97">重庆桐君阁股份有限公司医药批发分公司</option>
		
			<option value="96">sadafafda</option>
		
			<option value="95">dadadad</option>
		
			<option value="94">wqwqwqw</option>
		
			<option value="94">wqwqwqw</option>
		
			<option value="93">合肥腾飞</option>
		
			<option value="93">合肥腾飞</option>
		
	      </select>
       </td>
    </tr>
    <input type="hidden" name="client_id">
     <input type="hidden" name="Group_ID" value="4">
    </form>
  </table>
</body>
</html>
