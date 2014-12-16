

/* ******************************************************************************* * 
*
*	목 적	: 	입력폼이 포커를 얻으면 현재의 배경색을 변경함
*
* ******************************************************************************* */
	function myF(me) {
		me.style.backgroundColor = "#F1F4F5";
		me.style.color = "#000000";
		myMessOver(me);
	}


	function myB(me) {
		me.style.backgroundColor = "white";
		me.style.color = "#000000";
		myMessOut(this);
	}


/* *************************************************************************************** *
*
*	목 적	: 정보 메세지 창을 출력 합니다.
*	설 명	: 
*
* *************************************************************************************** */
	var myTrans =0;
	var myid = "msgBox";
	var myid2 = "msgBox2";

	function myMessOver(me) {
		var mess = me.title;
		document.all("msgBox2").innerHTML = mess;

		if(myTrans == '23') { myTrans = 0; };
		document.all(myid).filters.revealTrans.stop();
		document.all(myid).filters.revealTrans.transition = myTrans++;
		document.all(myid).filters.revealTrans.apply();
		
		document.all(myid).style.visibility =  "visible";
		document.all(myid).filters.revealTrans.play();
	}


	function myMessOut(me) {
		document.all(myid).style.visibility =  "hidden";
	}	


	window.onscroll = fixbox;
	
	function fixbox() {
		document.all(myid).style.left = document.body.scrollLeft + (document.body.clientWidth / 2) - 200;
		document.all(myid).style.top = document.body.scrollTop + document.body.clientHeight - 120;
	}

