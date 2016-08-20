function openwinfun(hrefvalue,winname,heightvalue,widthvalue)
{
	var statusstr;
	statusstr='location=no,height='+heightvalue+', width='+widthvalue+', toolbar=no, menubar=no, scrollbars=yes, resizable=no, location=no, status=no'
	window.open(hrefvalue,winname,statusstr);
}

function openwinfun1(hrefvalue,winname,heightvalue,widthvalue)
{
	var statusstr;
	statusstr='location=no,height='+heightvalue+', width='+widthvalue+', toolbar=no, menubar=no, scrollbars=yes, resizable=yes, location=no, status=no'
	window.open(hrefvalue,winname,statusstr);
}