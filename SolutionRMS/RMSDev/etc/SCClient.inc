<SCRIPT LANGUAGE="vbscript">
<!--
'-----------------
'  전역 상수 정의
'-----------------
'*****TransMode
const meNO_TRANS = 0   'No Transaction
const meINS_TRANS = 1  'Insert Transaction
const meUPD_TRANS = 2  'Update Transaction
const meDEL_TRANS = 3  'Delete Transaction

'*****각종 처리후 화면의 대기상태를 제어한다.
const meWAIT_ON = True    
const meWAIT_OFF = False

'*****스프레드쉬트의 INS/DEL/UP/DOWN/ENTER Function key
const meINS_ROW = 120
const meDEL_ROW = 121
const meCR= 13
const meUP_ROW = 38
const meDOWN_ROW = 40
const meENTER = 13

'*****스프레드쉬트의 데이터 변경 표시 플래그
const meCLS_FLAG = 0      '플래그 클리어
const meINS_FLAG = 1      'INSERT Flag
const meUPD_FLAG = 2      'UPDATE Flag
const meDEL_FLAG = 4      'DELETE FLAG
const meIU_FLAG = 3       'Insert & Update Flag
const meID_FLAG = 5       'Insert & Delete Flag
const meUD_FLAG = 6       'Update & Delete Flag
const meIUD_FLAG = 7      'Insert & Update & Delete Flag

Const meSPREAD_HEADER = -1000 'Spread Sheet Header Constant
Const SPREAD_HEADER = -1000

'*****메시지 상수
const meNO_DATA = "자료가 없습니다."
const meIS_DATA = "자료가 이미 있습니다."
const meWRONG_CHOICE = "할 자료의 선택이 잘못되었습니다"
const mePROC_DONE = " 되었습니다."
const meMAKE_CHOICE = "(을)를 선택하시기 바랍니다."
const meWRONG_INPUT = "입력이 잘못되었습니다."
const meSURELY_INPUT = "(을)를 반드시 입력하시기 바랍니다."
const meJOB_YESNO = "분 정도의 시간이 소요됩니다. 작업을 하시겠습니까?"
const meHAPPEN_ERROR = "처리 중 오류가 발생했습니다."
const meWAITING = "잠시만 기다리십시오."
const meJOB_DONE = "(이)가 완료되었습니다."
const meCRITICAL_ERROR = "프로그램 오류입니다. 전산실로 연락하십시오."
const meSERVERIP = "localhost"
'*****Tab/TabOn Image URL
Const meURL_TABON    = "url(../../../images/imgTabOn.gif)"
Const meURL_TAB      = "url(../../../images/imgTab.gif)"

'*****시스템 정보
'------------------------------------------------------
const meSYSID		  = "HCM"		'시스템에 맞게 변경
'------------------------------------------------------
const meWIN_NAME      = "WORK"
const mePARAM_SEP	  = "&"			'파라메터간 구별자
const mePARAMS_SEC    = "Params"		'공통/개별 파라메터의 등록 섹션
const meCOMPARAMS_KEY = "CM"	    		'공통 파라메터의 등록 Key
const meCONFIG_KEY    = "CONFIG"		'Config 등록 Key

'-----------------
'  전역 변수 정의
'-----------------
'*****공통 파라메터 정의
Dim gstrUsrID		'로그인ID
Dim gstrEmpNo		'사번
Dim gstrUsrName		'사용자명 
Dim gstrUsrIP		'사용자 IP
Dim gstrUsrBU		'사용자 회사코드
Dim gstrUsrBUName	'사용자 회사명
Dim gstrUsrLOCId	'사용자 지역ID
Dim gstrUsrLOC		'사용자 지역코드
Dim gstrUsrLOCName	'사용자 지역명
Dim gstrUsrCC		'사용자 부서코드
Dim gstrUsrCCName	'사용자 부서명
Dim gstrUsrOC		'사용자 OC코드
Dim gstrUsrOCName	'사용자 OC명
Dim gstrUsrMU		'사용자 MU코드
Dim gstrUsrMUName	'사용자 MU명
Dim gstrUsrPU		'사용자 PU코드(대표PU)
Dim gstrUsrPUName	'사용자 PU명
Dim gstrUsrGrade	'사용자 직급
Dim gstrUsrORGAuth  '사용자 COA에 대한 권한 LOC/MU/OC/CC/PU 11111111, 00000000 

Dim gstrSysID        '시스템ID
Dim gstrComParams	 '공통 파라메터
Dim gstrConfigXml	 '공통기본정보 XML
Dim gstrMDDIR
gstrMDDIR = http://10.110.10.89:8080/RMSDev_/MD/SrcWeb/MD/
'*****형식변환 여부
Dim gblnCnvtChkbox		'Checkbox의 값을 0,1,-1 <--> N,Y로 변환 여부
gblnCnvtChkbox = True	'변환으로 초기화

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 테스트 로그인 공통파라메터 설정 --> 개발자용임, 서비스 모드에서는 로그인 정보로 바뀜
' 개발자 모드 TRUE ::: 아래의 공통코드를 Config 정보로 사용합니다.
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Dim DevMode : DevMode = True
If DevMode=True then
	gstrDSN  = "MCDEV_NEW"
	
	gstrUsrName="관리자" 
	gstrUsrIP="127.0.0.1"
	gstrUsrBU="MC"
	gstrUsrBUName=""
	gstrUsrLOCId="1"
	gstrUsrLOC="01"
	gstrUsrLOCName=""
	gstrUsrCC=""
	gstrUsrCCName=""
	gstrUsrOC="40"
	gstrUsrOCName=""
	gstrUsrMU="10"
	gstrUsrMUName=""
	gstrUsrPU="10"
	gstrUsrPUName=""
	gstrUsrGrade="01"
	gstrUsrORGAuth="11111111"
	    
	gstrUsrInfo = "USRID="& gstrUsrID &"&EMPNO="& gstrEmpNo &"&USRNAME="& gstrUsrName &"&USRIP="& gstrUsrIP &"&USRBU="& gstrUsrBU &"&USRBUNAME="& gstrUsrBUName & "&USRLOCID="& gstrUsrLOCId &"&USRLOC="& gstrUsrLOC &"&USRLOCNAME="& gstrUsrLOCName &"&USRCC="& gstrUsrCC &"&USRCCNAME="& gstrUsrCCName &"&USROC="& gstrUSROC &"&USROCNAME="& gstrUSROCName &"&USRMU="& gstrUsrMU &"&USRMUNAME="& gstrUsrMUName &"&USRPU="& gstrUsrPU &"&USRPUNAME="& gstrUsrPUName &"&USRGRADE="& gstrUsrGrade &"&USRORGAUTH="& gstrUsrORGAuth
	gstrConfigInfo = "<config><authentication>test</authentication><company>"& gstrUsrBU &"</company><dsn>"& gstrDSN &"</dsn><wrkusr>"& gstrEmpNo &"</wrkusr><usrip>"& gstrUsrIP &"</usrip><mode>Dev</mode></config>"
end if
Sub gSetSession(byval strgstrUsrID,byval strgstrEmpNo,byval strgstrUsrName)
	gstrDSN  = "MCDEV_NEW"
	gstrUsrID= strgstrUsrID
	gstrEmpNo= strgstrEmpNo
	gstrUsrName= strgstrUsrName
	gstrUsrIP="127.0.0.1"
	gstrUsrBU=""
	gstrUsrBUName=""
	gstrUsrLOCId="1"
	gstrUsrLOC="01"
	gstrUsrLOCName=""
	gstrUsrCC=""
	gstrUsrCCName=""
	gstrUsrOC="40"
	gstrUsrOCName=""
	gstrUsrMU="10"
	gstrUsrMUName=""
	gstrUsrPU="10"
	gstrUsrPUName=""
	gstrUsrGrade="01"
	gstrUsrORGAuth="11111111"
	    
	gstrUsrInfo = "USRID="& gstrUsrID &"&EMPNO="& gstrEmpNo &"&USRNAME="& gstrUsrName &"&USRIP="& gstrUsrIP &"&USRBU="& gstrUsrBU &"&USRBUNAME="& gstrUsrBUName & "&USRLOCID="& gstrUsrLOCId &"&USRLOC="& gstrUsrLOC &"&USRLOCNAME="& gstrUsrLOCName &"&USRCC="& gstrUsrCC &"&USRCCNAME="& gstrUsrCCName &"&USROC="& gstrUSROC &"&USROCNAME="& gstrUSROCName &"&USRMU="& gstrUsrMU &"&USRMUNAME="& gstrUsrMUName &"&USRPU="& gstrUsrPU &"&USRPUNAME="& gstrUsrPUName &"&USRGRADE="& gstrUsrGrade &"&USRORGAUTH="& gstrUsrORGAuth
	gstrConfigInfo = "<config><authentication>test</authentication><company>"& gstrUsrBU &"</company><dsn>"& gstrDSN &"</dsn><wrkusr>"& gstrEmpNo &"</wrkusr><usrip>"& gstrUsrIP &"</usrip><mode>Dev</mode></config>"

End Sub
'--------------<<GET방식의 권한정보 처리>>------------------------
Dim gstrAuth		  '현재 웹페이지의 기본기능권한
Dim gstrRoleAuth	  '현재 웹페이지의 Role권한
Dim gstrRoleID		  '현재 웹페이지가 속한 Role ID

gstrAuth =     "<%=Request.QueryString("AUTH")%>"
gstrRoleAuth = "<%=Request.QueryString("ROLEAUTH")%>"
gstrRoleID =   "<%=Request.QueryString("ROLEID")%>"

Dim mstrServer
mstrServer = "http://<%=Request.ServerVariables("SERVER_NAME")%>"

Dim mobjSelected
set mobjSelected = nothing

Dim mblnErrDisp
mblnErrDisp = true
'--------------------
'Sheet Color
'--------------------
Dim mlngDefTextColor
Dim mlngShadowColor, mlngShadowDark, mlngShadowText
Dim mlngOddRowBackColor, mlngOddRowForeColor
Dim mlngEvenRowBackColor, mlngEvenRowForeColor
Dim mlngSelBackColor, mlngSelForeColor
Dim mlngLockBackColor, mlngLockForeColor
Dim mlngGrayAreaBackColor, mlngGridColor

sub gSetSheetDefaultColor()
    mlngDefTextColor = &H666666
    mlngShadowColor = RGB(221, 221, 219): mlngShadowDark = 0: mlngShadowText = 0
    mlngOddRowBackColor = RGB(255, 255, 255): mlngOddRowForeColor = 0
    mlngEvenRowBackColor = RGB(227, 237, 244): mlngEvenRowForeColor = 0
    mlngSelBackColor = RGB(0, 128, 192): mlngSelForeColor = RGB(255, 255, 255)
    mlngLockBackColor = 0: mlngLockForeColor = 0
    mlngGrayAreaBackColor = RGB(255, 255, 255): mlngGridColor = 0
end sub

sub gSetSheetColor(objSCGLSpr, objSht)
	objSCGLSpr.SpreadColors objSht,mlngDefTextColor, mlngShadowColor, , , _
						   mlngOddRowBackColor, , mlngEvenRowBackColor, , _
						   mlngSelBackColor, mlngSelForeColor, _
						   mlngLockBackColor, mlngLockForeColor, mlngGrayAreaBackColor
end sub
'*************************
'  전역 프로시져 정의
'*************************
Function gCreateRemoteObject (byval strProgID)
	'strServer = "http://<%=Request.ServerVariables("SERVER_NAME")%>"
	set gCreateRemoteObject =  ADS.CreateObject(strProgID,mstrServer)
End Function
'====================================================
'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
'====================================================
sub gInitPageSetting (objSCGLCtl, strSysID)
	'++++++++++++++++++++++++++++++
	' 개발자 모드에서만 사용합니다
	'++++++++++++++++++++++++++++++
	If DevMode = True Then
		gSetSysID strSysID
		gSetComParams mobjSCGLCtl, gstrUsrInfo
		gSetConfigXml mobjSCGLCtl, gstrConfigInfo
	End if

	gSetPageAuthority	'권한설정	
	gFitFormSize		'화면크기조정	
	gInitComParams   objSCGLCtl, strSysID
	gGetConfigXml    objSCGLCtl
	gInitPageControl document.forms(0)
end sub

Sub gEndPage
	'Main Window인 경우는 Main Page로 분기
	if window.name = meWIN_NAME then 
		'닫기버튼 클릭시 Home Page로 이동
		if not window.event.srcElement is nothing then
			if ucase(window.event.srcElement.id) = "IMGCLOSE"  then 
			   if DevMode=True Then
			      gShowWindow "../../../Home.aspx", meWIN_NAME, 790, 625, ""
		       else
		          '------------ 시스템에 맞게 변경
		          'gShowWindow "../../../SFARGW/GW/Welcome.aspx?USERID=" & gstrUsrID & "&PNOW="& now, meWIN_NAME, 790, 625, ""
			      gShowWindow "../../../Home.aspx", meWIN_NAME, 790, 625, ""
		       end if 
	 	    end if
	    end if
	   'Main Window가 아닌 경우는 close시킴
	else
	   self.opener = self
	   window.close
	end if
End Sub
Sub gEndPage2	
	   self.opener = self
	   window.close
End Sub
Sub gSetAuthority (imgBtn, blnAuth)
	dim strImgSrc, strImgName
	
	on error resume next
	
	'권한에 따른 이미지 버튼 변경 처리
	if blnAuth=true then
		if imgBtn.tagName = "IMG" then imgBtn.style.visibility = "visible" else imgBtn.disabled=false
		imgBtn.style.cursor="HAND"
		imgBtn.accessKey=""	'권한있음 표시
	else
		if imgBtn.tagName = "IMG" then imgBtn.style.visibility = "hidden" else imgBtn.disabled=true
		imgBtn.style.cursor="default"
		imgBtn.accessKey="NO"	'권한없음 표시
	end if
	Err.Clear			'오류정보를 삭제
end sub

sub gSetAuthorityAll (strAuth)
	dim i, blnAuth

	on error resume next
	with document.forms(0)
		for i = 1 to len(strAuth)
			if mid(strAuth,i,1) = "0" then blnAuth=false else blnAuth=true
			select case i
				case 1	'신규/저장
					gSetAuthority .imgNew, blnAuth
					gSetAuthority .imgSave, blnAuth				
					gSetAuthority .imgAddRow, blnAuth
				case 2	'조회
				    gSetAuthority .imgFind, blnAuth
					gSetAuthority .imgQuery, blnAuth
				case 3	'삭제
					gSetAuthority .imgDelete, blnAuth
					gSetAuthority .imgDelRow, blnAuth
				case 4	'인쇄
					gSetAuthority .imgPrint, blnAuth
				case 5	'H처리 -> 서버 배치 프로그램 Submit
					gSetAuthority .imgManage, blnAuth
			end select	
		next
	end with
	Err.Clear
end sub

'Role권한을 이용하여 현재 페이지의 권한을 설정
sub gSetPageAuthority
	gSetAuthorityAll gstrRoleAuth
end sub

function gChkAuthority (imgBtn)
	if imgBtn.accessKey="NO" then
		msgbox  "사용권한이 없습니다!",vbExclamation
		gChkAuthority = false
	else
		gChkAuthority = true
	end if
end function

function gChkAuthoritySpr (form,byval KeyCode)
	dim strInsAuth, strDelAuth
	dim objIns, objDel
	
	on error resume next

	if not form is nothing then 
		set objIns = form.imgSave
		set objDel = form.imgDelete
		'Delete 버튼이 없는 경우
		if not isobject(objDel) then set objDel = form.imgDelRow
	else
		set objIns = imgSave
		set objDel = imgDelete
		'Delete 버튼이 없는 경우
		if not isobject(objDel) then set objDel = imgDelRow
	end if
	
	strInsAuth = objIns.accessKey
	strDelAuth = objDel.accessKey
	
	gChkAuthoritySpr = false
   	select case KeyCode
		case meINS_ROW, vbcr
			if strInsAuth="NO" then
				msgbox  "신규입력권한이 없습니다!",vbExclamation
				exit function
			end if
		case meDEL_ROW
			if strDelAuth="NO" then
				msgbox  "삭제권한이 없습니다!",vbExclamation
				exit function
			end if
	end select
	
	gChkAuthoritySpr = true
	Err.Clear			'오류정보를 삭제
end function

'==========================================================================================================
' ORGAuthority로 사용자 화면 제어
'==========================================================================================================
Sub gSetORGAuthority(strUsrORGAuth, txtLOC, imgLOC, txtMU, imgMU, txtOC, imgOC, txtCC, imgCC, txtPU, imgPU) 
 Dim i, blnAuth
 if strUsrORGAuth = "" then strUsrORGAuth = "00000000"
 on error resume next                  
		for i = 1 to len(strUsrORGAuth)
			if mid(strUsrORGAuth,i,1) = "0" then blnAuth=false else blnAuth=true
			select case i
				case 1	'LOC
					gSetAuthority  imgLOC, blnAuth
				    gSetInputStyle txtLOC, blnAuth
				case 2	'MU
					gSetAuthority  imgMU, blnAuth
				    gSetInputStyle txtMU, blnAuth
				case 3	'OC
					gSetAuthority  imgOC, blnAuth
				    gSetInputStyle txtOC, blnAuth
				case 4	'CC
					gSetAuthority  imgCC, blnAuth
				    gSetInputStyle txtCC, blnAuth
				case 5	'PU
					gSetAuthority  imgPU, blnAuth
				    gSetInputStyle txtPU, blnAuth
			end select	
		next
End Sub
'=========================================================================
'제목:승인관련 Script 모음
'입력:Approval 여부
'기능:승인화면일 경우 버튼들을 없앤다.
'=========================================================================
'---------------------------------------------------
'승인 페이지인지 아닌지를 true, false 를 Return 한다
'---------------------------------------------------
Function isApprovalPage()
	Dim gStrApproval   '현재 웹페이지가 APPROVAL PAGE 인가?
	gstrApproval = "<%=Request.QueryString("Approval")%>"
	IF gStrApproval="OK" then
		isApprovalPage = true
	Else
		isApprovalPage = false	
	End if   
End Function

'------------------------------------------------------------------------------------------------------------
'승인 화면을 요청하는 함수 완료시 true 반환	실패시 false 반환
'vntRet = gShowApprovalPopup("PO_HDR_ID=24", "SC", "PO요청입니다.","PO-2003-02011","PO_HDR_TABLE", "PO_HDR_STATUS", "PO_HDR_ID=24", "24")
'------------------------------------------------------------------------------------------------------------
Function gShowApprovalPopup(byval sViewURL, byval sSysID, byval sTitle, byval sDocNum, byval sTable, byval sField, byval sCondition, byval sLineKey)
	'++++++++++++++++++++++++++++++++++++++++
	'승인의 STATUS SC_APPROVE_STATUS 테이블에 정의가 되어 있다
	'PO,SC: 진행중-->승인요청--> 승인, 승인거절	
	'++++++++++++++++++++++++++++++++++++++++
    Dim vntRet, vntInParams, strSEP : strSEP=""
     if right(document.location,1) <> "&" then strSEP="&"
     sViewURL = document.location & strSEP & "Approval=OK&" & sViewURL
	 
	 If sSysID = "" then sSysID="SC"
	 
	 If DevMode=True Then
		vntInParams = Array(sViewURL, sSysID, sTitle, sDocNum, sTable, sField, sCondition, "S003023", "김대희", sLineKey )
		vntRet = window.showModalDialog ("../../../SC/SrcWeb/SCAP/PopupDraft.aspx?UserID="& gStrUsrID,vntInParams,"dialogWidth=810px;dialogHeight=670px;status=no;center=yes;border=thin;resize=yes;statusbar=yes;")
		gShowApprovalPopup = vntRet
		Exit Function
	 End If	
	
	 vntInParams = Array(sViewURL, sSysID, sTitle, sDocNum, sTable, sField, sCondition, gstrUsrId, gstrUsrName, sLineKey)
	 vntRet = window.showModalDialog ("../../../SFARGW/GW/Approval/PopupDraftform.aspx?UserID="& gStrUsrID,vntInParams,"dialogWidth=810px;dialogHeight=670px;status=no;center=yes;border=thin;resize=yes;statusbar=yes;")
	 gShowApprovalPopup = vntRet
End Function

'--------------------------------------------------------------------------------------------------
' (입력불가) 승인된 페이지내의 UI 부분중 Textbox 및 form 에 해당하는 UI Control을 입력불가로 만들때
'--------------------------------------------------------------------------------------------------
Sub gSetApprovalPage(form, strAppStatus)
 	  dim intCnt,i
	  dim strClass, strPosition
	  dim objElement
	  dim vntParams
	  
	  On Error Resume Next
	  '각 모듈별 Status 가 들어가도록 한다. SC, PO, MF, AP :시스템 기본
	  if strAppStatus = "진행중" Or strAppStatus = "승인거절" then exit sub
	  '영업
	  if strAppStatus = "주문" then exit Sub
	  
	  if form is nothing then exit sub
	  
	  intCnt = form.elements.length
	  for i=0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.type = "text" or objElement.type = "select-one" or _
	 		objElement.type = "select-multiple" or objElement.type="radio" or objElement.type="checkbox" or _
	 		objElement.tagName = "TEXTAREA" then
		
			'비입력모드설정
			strClass = replace (objElement.className, "NO", "")
	  		strPosition = Mid (strClass, Len(strClass)-2, Len(strClass))
				Select Case strPosition
					Case "T_L"
						strClass	= "INPUTB_L"
					Case "T_R"
						strClass	= "INPUTB_R"
					Case "PUT"
						strClass	= "INPUTB"
				End Select
			objElement.className	= "NO" & strClass
	  		gSetInputStyle objElement, false
	 	end if
	next
	gSetAuthorityAll ("00000000")
	err.clear
end sub
'-------------------------------------------------------
' (입력불가) 승인된 페이지내의 쉬트를 입력불가로 만들때
'-------------------------------------------------------
Sub gSetApprovalSht(sprSht, strAppStatus)
     On Error Resume Next
     '각 모듈별 Status 가 들어가도록 한다. SC, PO :시스템 기본
     if strAppStatus = "진행중" Or strAppStatus = "승인거절" then exit sub
	 
	 mobjSCGLSpr.SetOperationMode sprSht, 1	
End Sub  
'----------------------------------------------------------------------------------------------------
' (입력가능) 승인된 페이지내의 UI 부분중 Textbox 및 form 에 해당하는 UI Control을 입력가능으로 만들때
'----------------------------------------------------------------------------------------------------
Sub gSetNotApprovalPage(form)
 	  dim intCnt,i
	  dim strClass, strPosition
	  dim objElement
	  dim vntParams
	  
	  On Error Resume Next
	 
	  if form is nothing then exit sub
	  
	  intCnt = form.elements.length
	  
	  for i=0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.type = "text" or objElement.type = "select-one" or _
	 		objElement.type = "select-multiple" or objElement.type="radio" or objElement.type="checkbox" or _
	 		objElement.tagName = "TEXTAREA" then
		
			gSetInputStyle objElement, true
					
			'비입력모드설정
			strClass = replace (objElement.className, "NO", "")
	  		strPosition = Mid (strClass, Len(strClass)-2, Len(strClass))
				Select Case strPosition
					Case "B_L"
						strClass	= "INPUT_L"
					Case "B_R"
						strClass	= "INPUT_R"
					Case "UTB"
						strClass	= "INPUT"
				End Select
				
				objElement.className	=  strClass
				
				If objElement.disabled = True Then
					objElement.ClassName	= "NO" & Replace(strClass,"INPUT", "INPUTB")
				End If
	  	end if
	  	
	  	'gPOSetInputStyle objElement, true
	  	
	next
	gSetAuthorityAll ("11111111")
	err.clear
end sub
'--------------------------------------------------------
' (입력가능) 승인된 페이지내의 쉬트를 입력가능으로 만들때
'--------------------------------------------------------
Sub gSetNotApprovalSht(sprSht)
  On Error Resume Next
  if sprSht is nothing then exit sub
  mobjSCGLSpr.SetOperationMode sprSht, 0	
End Sub  
'========================================================

sub gFitFormSize
    '새창이 아니면 resizeto 를 할 필요 없음
	if window.document.all("tblForm") is nothing Or window.name = meWIN_NAME then exit sub
	tblForm.style.left=0
	tblForm.style.top=0

	'기본 윈도우
	'window.resizeto tblForm.offsetWidth+38,tblForm.offsetheight+145
	'메뉴바, 툴바가 없는 경우
	window.resizeto tblForm.offsetWidth+38,tblForm.offsetheight+52
end sub

Sub gInitComParams (objSCGLCtl, strSysID)
	gSetSysID strSysID
	gGetComParams objSCGLCtl
end sub

Sub gSetSysID (byval strSysID)
	gstrSysID = strSysID
end sub

Sub gGetComParams (objSCGLCtl)
	'등록섹션으로부터 공통파라메터를 읽는다.
	gstrComParams = objSCGLCtl.RegGetSetting(meSYSID,mePARAMS_SEC,meCOMPARAMS_KEY,"")
	
	'각 공통파라메터 변수로 분리
	gstrUsrID = objSCGLCtl.SplitData2(gstrComParams,"USRID","")
	gstrEmpNo = objSCGLCtl.SplitData2(gstrComParams,"EMPNO","")
	gstrUsrName = objSCGLCtl.SplitData2(gstrComParams,"USRNAME","")
	gstrUsrIP = objSCGLCtl.SplitData2(gstrComParams,"USRIP","")
	gstrUsrBU = objSCGLCtl.SplitData2(gstrComParams,"USRBU","")
	gstrUsrBUName = objSCGLCtl.SplitData2(gstrComParams,"USRBUNAME","")
	gstrUsrLocId = objSCGLCtl.SplitData2(gstrComParams,"USRLOCID","")
	gstrUsrLoc = objSCGLCtl.SplitData2(gstrComParams,"USRLOC","")
	gstrUsrLocName = objSCGLCtl.SplitData2(gstrComParams,"USRLOCNAME","")
	gstrUsrOC = objSCGLCtl.SplitData2(gstrComParams,"USROC","")
	gstrUsrOCName = objSCGLCtl.SplitData2(gstrComParams,"USROCNAME","")
	gstrUsrCC = objSCGLCtl.SplitData2(gstrComParams,"USRCC","")
	gstrUsrCCName = objSCGLCtl.SplitData2(gstrComParams,"USRCCNAME","")
	gstrUsrOC = objSCGLCtl.SplitData2(gstrComParams,"USROC","")
	gstrUsrOCName = objSCGLCtl.SplitData2(gstrComParams,"USROCNAME","")
	gstrUsrMU = objSCGLCtl.SplitData2(gstrComParams,"USRMU","")
	gstrUsrMUName = objSCGLCtl.SplitData2(gstrComParams,"USRMUNAME","")
	gstrUsrPU = objSCGLCtl.SplitData2(gstrComParams,"USRPU","")
	gstrUsrPUName = objSCGLCtl.SplitData2(gstrComParams,"USRPUNAME","")
	gstrUsrGrade = objSCGLCtl.SplitData2(gstrComParams,"USRGRADE","")
	gstrUsrORGAuth = objSCGLCtl.SplitData2(gstrComParams,"USRORGAUTH","")
	
	'업무개체로 전달될 공통파라메터 조정
	gstrComParams = gstrComParams & mePARAM_SEP & "SYSID=" & gstrSysID
        gstrConfigInfo = "<config><authentication>test</authentication><company>"& gstrUsrBU &"</company><dsn>"& gstrDSN &"</dsn><wrkusr>"& gstrEmpNo &"</wrkusr><usrip>"& gstrUsrIP &"</usrip><mode>Dev</mode></config>"
	gSetConfigXml mobjSCGLCtl, gstrConfigInfo
end sub

Sub gSetComParams (objSCGLCtl, byval strComParams)
	'업무개체로 전달될 공통파라메터 조정
	gstrComParams = strComParams & mePARAM_SEP & "SYSID=" & gstrSysID
	
	'공통파라메터를 등록섹션에 저장
	objSCGLCtl.RegSaveSetting meSYSID,mePARAMS_SEC,meCOMPARAMS_KEY,strComParams
end sub

Function gGetPriParams (objSCGLCtl)
	gGetPriParams = objSCGLCtl.RegGetSetting(meSYSID,mePARAMS_SEC,gstrSysID, "")
end function

Sub gSetPriParams (objSCGLCtl, strPriParams)
msgbox meSYSID & "~" & mePARAMS_SEC & "~" & gstrSysID & "~" & strPriParams

	objSCGLCtl.RegSaveSetting meSYSID,mePARAMS_SEC,gstrSysID,strPriParams
end sub

Function gGetConfigXml (objSCGLCtl)
	gstrConfigXml = objSCGLCtl.RegGetSetting(meSYSID,mePARAMS_SEC,meCONFIG_KEY, "")
	gGetConfigXml = gstrConfigXml
end function

Sub gSetConfigXml (objSCGLCtl, strConfig)
	objSCGLCtl.RegSaveSetting meSYSID,mePARAMS_SEC,meCONFIG_KEY,strConfig
	gstrConfigXml = strConfig
end sub

Function gFileOpenDialog(CMDialog,strFile)
	on error resume next
    gFileOpenDialog = ""
    CMDialog.CancelError = True
    CMDialog.FileName = strFile
    CMDialog.ShowOpen
    '[취소] 버튼을 감지하기 위해 CancelError = True로 해야 함.
    If Err <> 32755 Then    '[취소]를 선택했을 때
        gFileOpenDialog = CMDialog.FileName
    else
		gFileOpenDialog = ""
    End If
    err.clear
End Function
'=================================================
' COMBO Box 처리
'=================================================
sub gLoadComboBox (cmbBox, vntData, blnDeftData)
	'현재 콤보 내용 삭제
	gRemoveComboBox cmbBox
	'콤보 내용 추가
	gAddComboBox cmbBox, vntData, blnDeftData
	'첫번째 데이터로 설정
	cmbBox.SelectedIndex = 0
end sub

sub gRemoveComboBox (cmbBox)
	dim i
	
	'현재 콤보 내용 삭제
	for i = 0 to cmbBox.Length-1
		cmbBox.remove(i)
	next
	cmbBox.length=0
end sub

sub gAddComboBox (cmbBox, vntData, blnDeftData)
	dim objOption, vntElement
	dim intColCnt, intRowCnt, i
	
	on error resume next
	
	'콤보 데이터를 만든다.
	intColCnt=-1: intColCnt=ubound(vntData,1)
	intRowCnt=-1: intRowCnt=ubound(vntData,2)
	
	'Default 데이터 ""를 삽입
	if blnDeftData then
		set objOption = document.createElement ("OPTION")
		objOption.value = ""
		objOption.Text = ""
		cmbBox.add objOption
		set objOption = nothing
	end if
	
	if intRowCnt = -1 then '1차원 데이터 경우
		for each vntElement in vntData
			set objOption = document.createElement ("OPTION")
			objOption.value = vntElement
			objOption.Text = vntElement
			cmbBox.add objOption
			set objOption = nothing
		next
	else					'2차원 데이터 경우
		for i = 0 to intRowCnt
			set objOption = document.createElement ("OPTION")
			objOption.value = vntData (0, i)
			if intColCnt = 0 then objOption.Text = vntData (0, i) else objOption.Text = vntData (1, i)
			cmbBox.add objOption
			set objOption = nothing
		next
	end if	
	Err.Clear	'오류정보를 삭제
end sub

sub gSetComboBoxListIndex (cmbBox, byval vntData, byval blnText)
	dim i
	
	'현재 콤보 내용 검색
	for i = 0 to cmbBox.Length-1
		if blnText then	'Text값으로 설정할 경우
			if cmbBox.options(i).text = vntData then
				cmbBox.SelectedIndex = i: exit sub
			end if
		else			'Value값으로 설정할 경우
			if cmbBox.options(i).value = vntData then
				cmbBox.SelectedIndex = i: exit sub
			end if
		end if
	next
	cmbBox.SelectedIndex = -1
end sub

function gSelectedComboText (cmbBox)
	gSelectedComboText = cmbBox.options(cmbBox.selectedIndex).text
End function
'-------------------------------------
'2차원배열을 1차원배열로 변환하는 함수
'-------------------------------------
function gArray2Single(ByVal vntData, ByVal intColidx)
    Dim intColCnt, intRowCnt, i
    Dim vntData2()
   
    On error resume next
	intColCnt=-1: intColCnt=ubound(vntData,1)
	intRowCnt=-1: intRowCnt=ubound(vntData,2)

	if intColidx="" or intColidx < 0 then intColidx=0
	    
	if intRowCnt = -1 then '1차원 데이터 경우
	else				   '2차원 데이터 경우
	Redim vntData2(intRowCnt)
		for i = 0 to intRowCnt
	 	  vntData2(i) = vntData(intColidx, i)
		next
	end if	
	gArray2Single = vntData2	
end function
'------------------------------------------------------------------------
'입  력: 배열형 데이타 (1차원 또는 2차원)
'	   blnDeftData = true이면 디폴트로 ""를 삽입
'기  능:  Sheet 콤보에 맞는 데이터 형태로 반환 (chr(9)를 분리자로 하는 문자열로 반환)
'------------------------------------------------------------------------
function gCnvtSheetComboData (vntData, blnDeftData)
	Dim lngRowCnt, lngColCnt, i
	Dim strRtn

	on error resume next
	
	if not isarray(vntData) then
		strRtn = vntData
	else
		'콤보 데이터를 문자열로 만든다.
		lngColCnt=-1: lngColCnt=ubound(vntData,1)
		lngRowCnt=-1: lngRowCnt=ubound(vntData,2)

		if lngRowCnt = -1 then '1차원 데이터 경우
			for i= 0 to lngColCnt
				strRtn = strRtn & vntData(i) & chr(9)
			next
		else					'2차원 데이터 경우
			for i = 0 to lngRowCnt
				if vntData(0,i) = vntData(1,i) then
					strRtn = strRtn & vntData(0,i) & chr(9)
				else
					strRtn = strRtn & vntData(0,i) & "-" & vntData(1,i) & chr(9)
				end if
			next
		end if
   	end if
	'디폴트 데이터 삽입
	if blnDeftData then strRtn = "" & chr(9) & strRtn
	gCnvtSheetComboData = strRtn
   	err.clear
end function
'=================================================
' Window Open 및 Popup Window Open
'=================================================
' 일반적인 팝업 윈도우
Sub gShowWindow (byval strPageURL, byval strWindowName, byval lngWidth, byval lngHeight, byval strOptions)
	dim lngTop, lngLeft
	
	'호출파라메터와 권한정보를 추가
	if instr(1,strPageURL,"?") = 0 then 
		strPageURL = strPageURL & "?"
	else
		strPageURL = strPageURL & mePARAM_SEP
	end if
		
	strPageURL = strPageURL & _
				"AUTH=" & gstrAuth & mePARAM_SEP & _
				"ROLEAUTH=" & gstrRoleAuth & mePARAM_SEP & _
				"ROLEID=" & gstrRoleID

	if strOptions="" then
		strOptions = "toolbar=no, location=no, menubar=no, scrollbars=no, status=yes, resizable=yes"
	end if
	'화면의 중앙에 위치시킨다.
	lngTop = (window.screen.height - lngHeight) / 2
	lngLeft = (window.screen.width - lngWidth) / 2
	strOptions = strOptions & ", width=" & lngWidth & ", height=" & lngHeight & _
				 ", top=" & lngTop & ", left=" & lngLeft

	window.open strPageURL,strWindowName,strOptions
end sub

'Modal 윈도우 Return 값이 있음
function gShowModalWindow (byval strPageURL, byval vntInParams, byval lngDialogWidth, byval lngDialogHeight)
	dim vntRet

	'호출파라메터와 권한정보를 추가
	if instr(1,strPageURL,"?") = 0 then 
		strPageURL = strPageURL & "?"
	else
		strPageURL = strPageURL & mePARAM_SEP
	end if
		
	strPageURL = strPageURL & _
				"AUTH=" & gstrAuth & mePARAM_SEP & _
				"ROLEAUTH=" & gstrRoleAuth & mePARAM_SEP & _
				"ROLEID=" & gstrRoleID

	vntRet = window.showModalDialog (strPageURL,vntInParams,"dialogWidth:" & lngDialogWidth & "px;dialogHeight:" & _
				 					lngDialogHeight & "px; status:no; center:yes; border:thin; resizable:yes; help:no;")
	gShowModalWindow = vntRet
end function

'ModelessWindow Return 값이 없음??
function gShowModelessWindow (byval strPageURL, byval vntInParams, byval lngDialogWidth, byval lngDialogHeight)
	dim vntRet

	'호출파라메터와 권한정보를 추가
	if instr(1,strPageURL,"?") = 0 then 
		strPageURL = strPageURL & "?"
	else
		strPageURL = strPageURL & mePARAM_SEP
	end if
		
	strPageURL = strPageURL & _
				"AUTH=" & gstrAuth & mePARAM_SEP & _
				"ROLEAUTH=" & gstrRoleAuth & mePARAM_SEP & _
				"ROLEID=" & gstrRoleID
	
	vntRet= window.showModelessDialog (strPageURL,vntInParams,"dialogWidth:" & lngDialogWidth & "px;dialogHeight:" & _
									   lngDialogHeight & "px; status:no; center:yes; border:thin; resizable:yes; help:no;")
	gShowModelessWindow = vntRet
end function

'PopUp Window 인가??
function gIsPopupWindow ()
	if ucase(window.clientInformation.toString) = "[OBJECT]" then 
	   gIsPopupWindow = true 
	else 
	   gIsPopupWindow = false
	end if
end function

'---------------------------------- 
' 레포트 보고서를 여는 Popup Window
'----------------------------------
Sub gShowReportWindow9  (Module, RptName, Params, Opt)  	
	If Opt<>"B" Then Opt="A"									 ''' Option은 A(Default-열기), B(Printer 인쇄)
	Dim path      : path="../../../SC/SrcWeb/SCRT/SCRTMAIN.aspx" ''' SCRTMAIN.aspx(Config)
	Dim addParams : addParams ="?Module="& Module &"&RptName="& URLEncode(RptName) &"&Params="& URLEncode(Params) &"&Opt="& Opt
	Dim myRndNum  : Randomize : myRndNum=Int((10000 * Rnd) + 1)  ''' 난수발생 초기화 1에서 10000까지 무작위 값 추출
	
	gShowWindow path & addParams, "Report"&myRndNum, "1024", "768", ""      ''' Open Window(일반창)	
end Sub

'----> (Crystal Report 10 버젼을 사용한다.)
Sub gShowReportWindow(Module, RptName, Params, Opt)  	
	If Opt<>"B" Then Opt="A"									 ''' Option은 A(Default-열기), B(Printer 인쇄)
	Dim path      : path="../../../SC/SrcWeb/SCRT/SCRTMAIN10.aspx" ''' SCRTMAIN.aspx(Config)
	Dim addParams : addParams ="?Module="& Module &"&RptName="& RptName &"&Params="& URLEncode(Params) &"&Opt="& Opt
	Dim myRndNum  : Randomize : myRndNum=Int((10000 * Rnd) + 1)  ''' 난수발생 초기화 1에서 10000까지 무작위 값 추출
	
	gShowWindow path & addParams, "Report"&myRndNum, "1024", "768", ""      ''' Open Window(일반창)	
end Sub

Sub gShowiFrameReport(iFrameName, Module, RptName, Params, Opt)
	If Opt <> "B" then Opt="A"
	path = "../../../SC/SrcWeb/SCRT/SCRTMAIN.aspx?Module="& Module &"&RptName="& RptName &"&Params="& URLEncode(Params) & "&Opt="& Opt
	iFrameName.location.href = path
End Sub                         

Function URLEncode(URLStr)	
	Dim sURL         '** 입력받은 URL 문자열
	Dim sBuffer      '** URL 인코딩 처리 중 URL 을 담을 버퍼 문자열
	Dim sTemp        '** 임시 문자열
	Dim cChar        '** URL 문자열 중 현재 인텍스의 문자
	Dim Index        
		
	sURL = Trim(URLStr)     '** URL 문자열을 얻는다.
	sBuffer = ""            '** 임시 버퍼용 문자열 변수 초기화.
	'----------------------------------------------------
	'* URL 인코딩 작업
	'----------------------------------------------------
	For Index = 1 To Len(sURL)
		'** 현재 인덱스의 문자를 얻는다.
		cChar = Mid(sURL, Index, 1)
		If cChar = "0" Or _
		(cChar >= "1" And cChar <= "9") Or _
		(cChar >= "a" And cChar <= "z") Or _
		(cChar >= "A" And cChar <= "Z") Or _
		cChar = "-" Or _
		cChar = "_" Or _
		cChar = "." Or _
		cChar = "*" Then        
			'** URL 에 허용되는 문자들 :: 버퍼 문자열에 추가한다.
			sBuffer = sBuffer & cChar 			            
		ElseIf cChar = " " Then  			        
			'** 공백 문자 :: + 로 대체하여 버퍼 문자열에 추가한다.
			sBuffer = sBuffer & "+"			            
		Else                       			        
			'** URL 에 허용되지 않는 문자들 :: % 로 인코딩해서 버퍼 문자열에 추가한다.
			sTemp = CStr(Hex(Asc(cChar)))   			            
			If Len(sTemp) = 4 Then
				sBuffer = sBuffer & "%" & Left(sTemp, 2) & "%" & Mid(sTemp, 3, 2)
			ElseIf Len(sTemp) = 2 Then  
				sBuffer = sBuffer & "%" & sTemp 
			End If 			                
		End If	
	Next         
		'** 결과를 리턴한다.
		URLEncode = sBuffer
End Function  

'-------------------------------
' 파일 첨부를 여는 Popup Window
'-------------------------------
Function gShowFileWindow(ByVal DOC_MODULE, ByVal FILE_ID)
       Dim fpath, vntInParams, vntRet
	   fpath = "../../../SC/SrcWeb/SCFU/SCFUFILEMODAL.aspx?FILE_ID="& FILE_ID &"&DOC_MODULE="& DOC_MODULE &"&now="& now
	   vntInParams = Array(DOC_MODULE, FILE_ID)
	   vntRet = gShowModalWindow(fpath, vntInParams, 410, 485)
	   gShowFileWindow = vntRet
End Function

'-------------------------------- 
'Excel Upload를 여는 PopUp Window (성공하면 Return True, 실패하면 Return false)
'--------------------------------
Function gShowExcelWindow(ByVal ID)
	Dim vntRet, vntInParams
	vntRet = gShowModalWindow("../../../SC/SrcWeb/SCEX/SCEXMAIN01.aspx?ID="& ID, vntInParams , 810,650)		
	gShowExcelWindow = vntRet
End Function

'-------------------------------
' CC에 기본 Mapping된 정보 조회 
'-------------------------------
Sub gGetCCDefault (objSCCDGet, strCC_CODE, objCC_CODE, objCC_NAME, objLOC_CODE, objLOC_NAME, _
				   objOC_CODE, objOC_NAME, objMU_CODE, objMU_NAME, objPU_CODE, objPU_NAME)
	dim vntData, lngRowCnt, lngColCnt, i
	
	On error resume next
	'Long Type의 ByRef 변수의 초기화
	lngRowCnt=clng(0)
	lngColCnt=clng(0)

	vntData = objSCCDGet.GetCCDefault(gstrConfigXml,lngRowCnt,lngColCnt,strCC_CODE)
	if not gDoErrorRtn ("gGetCCDefault") and lngRowCnt > 0 then
		'CC_CODE,CC_NAME,LOC_CODE,LOC_NAME,OC_CODE,OC_NAME,SC_MU_CODE,CODE_NAME,PU_CODE,PU_NAME
		for i = 0 to ubound(vntData)
			select case i
				case 0: if isobject(objCC_CODE) then objCC_CODE.value = vntData(i,0) 
				case 1: if isobject(objCC_NAME) then objCC_NAME.value = vntData(i,0)
				case 2: if isobject(objLOC_CODE) then objLOC_CODE.value = vntData(i,0)
				case 3: if isobject(objLOC_NAME) then objLOC_NAME.value = vntData(i,0)
				case 4: if isobject(objOC_CODE) then objOC_CODE.value = vntData(i,0)
				case 5: if isobject(objOC_NAME) then objOC_NAME.value = vntData(i,0)
				case 6: if isobject(objMU_CODE) then objMU_CODE.value = vntData(i,0)
				case 7: if isobject(objMU_NAME) then objMU_NAME.value = vntData(i,0)
				case 8: if isobject(objPU_CODE) then objPU_CODE.value = vntData(i,0)
				case 9: if isobject(objPU_NAME) then objPU_NAME.value = vntData(i,0)
			end select
		next
   	end if
End Sub
'------------------------------------------------------------------------
'입  력: strName = 분류명,vntInParams = 입력 파라메터 배열
'반  환: 팝업화면으로부터 전달된 배열
'기  능: 전체공통으로 사용하는 팝업처리 함수 
'------------------------------------------------------------------------
function gPopupCodeWindow (byval strName, byval vntInParams)
	dim vntRet
		Select Case Ucase(strName)
			case "SCCODE"			'클래스코드
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDClass.aspx",vntInParams , 414,425)
			case "LOC"				'지역
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDLOC.aspx",vntInParams , 414,425)
			case "OC"				'OC
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDOC.aspx",vntInParams , 414,425)
			case "MU"				'MU
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDMU.aspx",vntInParams , 414,425)
			case "PU"				'PU
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDPU.aspx",vntInParams , 416,430)
			case "CC"				'CC
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCC.aspx",vntInParams , 416,430)
			case "EMP"				'EMP
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDEMP.aspx",vntInParams , 414,425)
			case "USER"				'EMP
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDUSER.aspx",vntInParams , 414,425)
			case "POST"				'우편번호
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDPOST.aspx",vntInParams , 494,425)
			case "CURRENCYTYPE"		'CURRENCYTYPE
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCURRENCYTYPE.aspx",vntInParams , 410,425)
			case "CURRENCY"			'CURRENCY
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCURRENCY.aspx",vntInParams , 410,425)
			case "CURRENCYDAILY"	'CURRENCYDAILY
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCURRENCYDAILY.aspx",vntInParams , 416,425)
			case "MODCATEGORY"		'MODCATEGORY
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDMODCATEGORY.aspx",vntInParams , 410,425)
			case "VOUCHER"			'VOUCHER 전표번호설정
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDVOUCHER.aspx",vntInParams , 410,425)
			case "VOUCHER2"			'VOUCHER 전표번호설정
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDVOUCHER2.aspx",vntInParams , 410,425)
			case "BANKBRANCHREG"	'은행지점
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDBANKBRANCHREG.aspx",vntInParams , 410,425)
			case "BANKACCOUNT"		'은행계좌
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDBANKACCOUNT.aspx",vntInParams , 410,425)
			case "CUSTOMER"			'거래처
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCUSTOMER.aspx",vntInParams , 655, 425)
			case "CUSTOMERHDR"	    '거래처 Header
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCUSTOMERHDR.aspx",vntInParams , 450,425)
			'case "CUSTOMERDTL"		'거래처 Detail
			'	vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCUSTOMERDTL.aspx",vntInParams , 410,425)
			case "CUSTCONTACTS"	    '거래처 담당자
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCUSTCONTACTS.aspx",vntInParams , 410,425)
			case "INVENTORY"		'창고
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDINVENTORY.aspx",vntInParams , 410,425)
			case "CATEGORY"			'ITEM CATEGORY
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCATEGORY.aspx",vntInParams , 460,425)
			case "CATELEMENT"		'ITEM CATEGORY 구성요소
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCATELEMENT.aspx",vntInParams , 410,425)
			case "ITEMMST"			'ITEM 마스터
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDITEMMST.aspx",vntInParams , 530,425)
			case "UOM"				'UOM
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDUOM.aspx",vntInParams , 410,425)
			case "SUFFIXCODE"		'SUFFIX_CODE
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDSUFFIXCODE.aspx",vntInParams , 410,425)
			case "PAYCONDITION"		'지불조건
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDPAYCONDITION.aspx",vntInParams , 410,425)
			case "PAYGROUP"			'지불그룹
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDPAYGROUP.aspx",vntInParams , 410,425)
			case "NATYPE"			'계정TYPE
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDNATYPE.aspx",vntInParams , 410,425)
			case "NA"				'계정C
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDNA.aspx",vntInParams , 610,425)
			case "NA1"				'계정P
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDNA1.aspx",vntInParams , 410,425)
			case "NAALL"		    '계정ALL
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDNAALL.aspx",vntInParams , 410,425)
			case "NAMGMT"			'계정관리항목
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDNAMGMT.aspx",vntInParams , 410,425)
			case "COUNTRY"			'국가코드
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDCOUNTRY.aspx",vntInParams , 410,425)
			case "ITEMTYPE"			'계정타입
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDITEMTYPE.aspx",vntInParams , 410,425)
			case "TAXCODE"			'TAX코드
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDTAXCODE.aspx",vntInParams , 470,425)
			case "APPROVELIST"
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDAPPROVELIST.aspx",vntInParams , 480,365)
   			case "VESSEL"			'VESSEL 모선코드 조회
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCD/SCCDVESSEL.aspx",vntInParams , 410,410)
		End Select
	gPopupCodeWindow = vntRet
end function
'------------------------------------------------------------------------
'입  력: dteBox=날자표시 컨트롤, btnImg=달력팝업 이미지 버튼, runFuncs = 날자변경시 호출 함수
'반  환: 날자가 변경시 dteBox에 표시하고, runFuncs를 호출
'기  능: 전체공통으로 사용하는 달력팝업처리 함수 
'------------------------------------------------------------------------
sub gShowPopupCalendar (dteBox,btnImg,runFuncs)
	const DIALOG_WIDTH = 200
	const DIALOG_HEIGHT =205
	dim intLeftPos, intTopPos

	'Calendar Popup window 위치 계산 : btnImg 근처에 생기도록 한다.
	intLeftPos = getOffsetLeft(btnImg) + window.screenLeft
	intTopPos = getOffsetTop(btnImg) + btnImg.height + window.screenTop + 2
    
	'Calendar의 위치가 윈도우 크기를 벗어난 경우는 위쪽으로 배치
	if intTopPos + DIALOG_HEIGHT > window.screen.height then _
		intTopPos = intTopPos - btnImg.height - DIALOG_HEIGHT
		
	'Calendar Popup window 호출
	vntInParams = array(dteBox.value)
	vntRet = window.showModalDialog ("../../../SC/SrcWeb/SCCM/SCCalendar.aspx",vntInParams,"dialogWidth=" & DIALOG_WIDTH & "px;dialogHeight=" & DIALOG_HEIGHT & "px;dialogLeft=" & intLeftPos & ";dialogTop=" & intTopPos & ";help=no;status=no;unadorned=on;")
	
	'값이 변경되었을 때 지정함수 호출
	if dteBox.value <> vntRet then 
		dteBox.value = vntRet
		gSetChangeFlag dteBox	'XML Binding이 된 경우 플래그 변경 
		if runFuncs <> "" then execute (runFuncs)
	end if
end sub
'=================================================
' 기타  유틸리티
'=================================================
function gDoErrorRtn (Byval strSource)
	if Err.number <> 0 then
		'오류의 발생 위치를 명시한다.
		If strSource <> "" Then
			strSource = "->" & strSource
		End If
    
		MsgBox "오류번호 : " & Err.Number & chr(13) & chr(10) & _
		       		Err.Description & chr(13) & chr(10) & _
			    	"(" & Err.Source & strSource & ")", _
				vbCritical, "시스템 오류"		

		gDoErrorRtn = true	'오류 발생
		Err.Clear			'오류정보를 삭제
	else
		gDoErrorRtn = false	'오류 발생치 않음
	end if
end function

Sub gErrorMsgBox(strMsg, strTitle)
    If strTitle = "" Then strTitle = "Error!"
    MsgBox strMsg, vbExclamation, strTitle
End Sub

Sub gOkMsgBox(strMsg, strTitle)
    If strTitle = "" Then strTitle = "확인!"
    MsgBox strMsg, vbOKOnly + vbInformation, strTitle
End Sub

Function gOkCalcelMsgBox(strMsg, strTitle)
    If strTitle = "" Then strTitle = "확인/취소 선택?"
    gOkCalcelMsgBox = MsgBox(strMsg, vbOKCancel + vbInformation, strTitle)
End Function

Function gYesNoMsgBox(strMsg, strTitle)
    If strTitle = "" Then strTitle = "예/아니오 선택?"
    gYesNoMsgBox = MsgBox(strMsg, vbYesNo + vbQuestion, strTitle)
End Function

Function gYesNoCancelMsgBox(strMsg, strTitle)
    If strTitle = "" Then strTitle = "예/아니오/취소 선택?"
    gYesNoCancelMsgBox = MsgBox(strMsg, vbYesNoCancel + vbQuestion, strTitle)
End Function

sub gFlowWait (blnWait)
	dim objElement, strCursor
	on error resume next

	if blnWait=true then
		document.body.style.cursor ="wait"
		tblWaitP.style.visibility = "visible"
		mobjSCGLCtl.DoEventQueue	'이벤트 큐 처리
	else
		document.body.style.cursor ="default"
		tblWaitP.style.visibility = "hidden"
	end if
	
	Err.Clear	'오류정보를 삭제
end sub

Function gGetFlowWait
	if MainBody.style.cursor="wait" then
		gGetFlowWait = true
	else
		gGetFlowWait = false
	end if
end function

sub gWriteText (byval objTagID, byval strText)
	if objTagID = "" then
		'window.status = strText
	else	
		objTagID.innerText = strText
	end if
end sub

Function gGetFileAccessInfo(filespec, ByVal strInfo)
    Dim fs, d, f
    Set fs = CreateObject("Scripting.FileSystemObject")
      
    Select Case UCase(strInfo)
        Case "PATH"             '파일명을 포함한 패스
                    Set f = fs.GetFile(filespec)
                    gGetFileAccessInfo = f.Path
        Case "DRIVE"            '드라이브
                    Set f = fs.GetFile(filespec)
                    gGetFileAccessInfo = f.Drive
        Case "DATECREATED"      '만든 날짜
                    Set f = fs.GetFile(filespec)
                    gGetFileAccessInfo = f.DateCreated
        Case "DATELASTACCESSED" '마지막으로 액세스한 날짜
                    Set f = fs.GetFile(filespec)
                    gGetFileAccessInfo = f.DateLastAccessed
        Case "DATELASTMODIFIED" '마지막으로 수정한 날짜
                    Set f = fs.GetFile(filespec)
                    gGetFileAccessInfo = f.DateLastModified
        Case "SIZE"
                    Set f = fs.GetFile(filespec)
                    gGetFileAccessInfo = f.Size
        Case "DIR"              '파일이 속한 디렉토리
                    gGetFileAccessInfo = fs.GetParentFolderName(filespec)
        Case "FILENAME"         '파일명
                    gGetFileAccessInfo = fs.GetFileName(filespec)
        Case "EXT"              '확장자명
                    gGetFileAccessInfo = fs.GetExtensionName(filespec)
        Case "DRIVELETTER"
                    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(filespec)))
                    gGetFileAccessInfo = d.DriveLetter
        Case "SHARENAME"
                    Set d = fs.GetDrive(fs.GetDriveName(fs.GetAbsolutePathName(filespec)))
                    gGetFileAccessInfo = d.ShareName
        Case "TYPE"
                    Set f = fs.GetFile(filespec)
                    gGetFileAccessInfo = f.Type    
    End Select
End Function

Function gSplitParams (objSCGLCtl, strParams, strVariable)
	gSplitParams = objSCGLCtl.SplitData2(strParams,strVariable,"")
end function

Sub gIgnoreMouseRightBtn
	'exit sub
	'도큐먼트 오브젝트에서의 마우스 오른쪽 버튼을 사용하지 못하게 한다.
	'document_onmousedown에서 콜
	if window.event.button =2 and _
	   (window.event.srcElement.tagName = "BODY" or window.event.srcElement.tagName = "TD") then
		'window.event.cancelBubble = true
		'window.event.returnValue = false
		alert ("오른쪽 마우스 버튼은 사용할 수 없습니다!.")
	end if
End Sub
'------------------------------------------------------------------------
'기  능: ① Enter Key를 Tab Key로 변환
'------------------------------------------------------------------------
Sub gEnterToTab
	If Window.Event.KeyCode = 13 Then Window.Event.KeyCode = 9
End Sub

sub gSetFocus ()
	document.activeElement.select
end sub

'=============================
'  Document 기본 이벤트 정의
'=============================
'---------------------------------------------------------------------------
' 기능 : 도큐먼트 오브젝트에서의 마우스 오른쪽 버튼을 사용하지 못하게 한다.
'---------------------------------------------------------------------------
Sub document_onmousedown
	exit sub	'for test
	if window.event.button =2 and _
	   (window.event.srcElement.tagName = "BODY" or window.event.srcElement.tagName = "TD") then
		'window.event.cancelBubble = true
		'window.event.returnValue = false
		alert ("오른쪽 마우스 버튼은 사용할 수 없습니다!.")
	end if
End Sub
'---------------------------------------------------------------------------
' 기능 : Text Object에서 focus를 받으면 전체 데이터를 Select시킴
'---------------------------------------------------------------------------
Sub document_onclick
	on error resume next
	dim objElement
	set objElement = document.activeElement
	if objElement.type="text" and objElement.readonly =false then objElement.select
	err.Clear
End Sub

'---------------------------------------------------------------------------
' 기능 : CR을 Tab으로 변환, BackSpace로 인한 화면 전 페이지로 이동을 막음
'---------------------------------------------------------------------------
Sub document_onkeydown
	on error resume next
	dim objElement
	
	set objElement = window.event.srcElement
	If Window.Event.KeyCode = 13 and objElement.tagName <> "TEXTAREA"  and objElement.tagName <> "OBJECT" Then 
		Window.Event.KeyCode = 9
	elseif 	window.event.keyCode = 8 and _
			objElement.tagName <> "INPUT" AND _
			objElement.tagName <> "TEXTAREA" then	'BackSpace 키 무시
		window.event.returnValue = false
	end if
	err.Clear
End Sub

'---------------------------------------------------------------------------
' 기능 : 각 오브젝트의 속성에 맞게 입력 제어
'---------------------------------------------------------------------------
Sub document_onkeypress
	on error resume next
	dim objElement
	dim intKeyVal
	dim vntType, vntParams
	
	set objElement = window.event.srcElement 'document.activeElement
	intKeyVal = window.event.keyCode
	
	if objElement.type="text" and objElement.accesskey <> "" then
		vntType = split(objElement.accesskey,",")
		if vntType(0) <> "" then
			vntParams = split(vntType(0),":")
			select case ucase(vntParams(0))
				case "DATE","MON"
						gDateInput
				case "TIME"
						gTimeInput
				case "NUM"
						gNumericInput
				case "INT"
						gIntegerInput
				case "DEC"
						gDecimalInput
				case "TEL"
						gTelephoneInput
			end select
		end if
	end if
	err.Clear
End Sub

'---------------------------------------------------------------------------
' 기능 : 바인딩 오브젝트의 값이 변경된 경우 변경 작업을 처리
'---------------------------------------------------------------------------
Sub document_onafterupdate
	gSetChange
End Sub

'---------------------------------------------------------------------------
' 기능 : 바인딩이 되지 않은  오브젝트의 값이 변경된 경우 변경 작업을 처리
'---------------------------------------------------------------------------
Sub document_onpropertychange
	if not mobjSelected is nothing then gSetDataFormat mobjSelected
	' set mobjSelected = document.activeElement
End Sub

'---------------------------------------------------------------------------
' 기능 : 오브젝트가 변경된 경우 DataFormat 변경 및 플래그 변경
'---------------------------------------------------------------------------
sub gSetChange
	on error resume next
	dim objElement
	
	'Data Format변경
	set objElement = window.event.srcElement
	gSetDataFormat objElement
	
	gSetChangeFlag objElement
	
	err.Clear
end sub

sub gSetChangeFlag (objElement)
	on error resume next

	'Data Flag 변경
	dim strXmlID, objXmlID

	strXmlID = trim(replace(objElement.dataSrc,"#",""))
	if strXmlID <> "" then 
		set objXmlID = document.getElementById(strXmlID)
		if not objXmlID is nothing then gXMLDataChanged objXmlID
	end if

	err.Clear
end sub
'---------------------------------------------------------------------------
' 기능 : HTML 컨트롤 onchange이벤트에서 자기 자신을 변경할 경우 onblur에서 gSetOnblur 호출
'---------------------------------------------------------------------------
sub gSetOnblur
	on error resume next
	dim objElement

	set objElement = window.event.srcElement
	objElement.value = objElement.value	
		
	err.Clear
end sub
'=================================================
' Formatting 관련
'=================================================
'---------------------------------------------------------------------------
' 기능 : NUM 속성으로 0~9만 입력 허용
'---------------------------------------------------------------------------
sub gNumericInput ()
	dim intKeyVal
	
	intKeyVal = window.event.keyCode
	
	if (intKeyVal >= 48 and intKeyVal <= 57) then
		window.event.returnValue = true
	else
		window.event.returnValue = false
	end if
end sub

'---------------------------------------------------------------------------
' 기능 : INT 속성으로 0~9와 "," "-"만 입력 허용
'---------------------------------------------------------------------------
sub gIntegerInput ()
	dim intKeyVal
	
	intKeyVal = window.event.keyCode
	
	if (intKeyVal >= 48 and intKeyVal <= 57) or intKeyVal = 44 or intKeyVal = 45 then
		window.event.returnValue = true
	else
		window.event.returnValue = false
	end if
end sub

'---------------------------------------------------------------------------
' 기능 : DEC 속성으로 '0~9와 "," "." "-"만 입력 허용
'---------------------------------------------------------------------------
sub gDecimalInput ()
	dim intKeyVal
	
	intKeyVal = window.event.keyCode
	
	if (intKeyVal >= 48 and intKeyVal <= 57) or (intKeyVal >= 44 and intKeyVal <= 46) then
		window.event.returnValue = true
	else
		window.event.returnValue = false
	end if
end sub

'---------------------------------------------------------------------------
' 기능 : DATE 속성으로 0~9와 "-"만 입력 허용
'---------------------------------------------------------------------------
sub gDateInput ()
	dim intKeyVal
	
	intKeyVal = window.event.keyCode
	
	if (intKeyVal >= 48 and intKeyVal <= 57) or intKeyVal = 45 then
		window.event.returnValue = true
	else
		window.event.returnValue = false
	end if
end sub

'---------------------------------------------------------------------------
' 기능 : TIME 속성으로 0~9와 ":"만 입력 허용
'---------------------------------------------------------------------------
sub gTimeInput ()
	dim intKeyVal
	
	intKeyVal = window.event.keyCode
	
	if (intKeyVal >= 48 and intKeyVal <= 57) or intKeyVal = 58 then
		window.event.returnValue = true
	else
		window.event.returnValue = false
	end if
end sub

'---------------------------------------------------------------------------
' 기능 : TEL 속성으로 '0~9와 "(" ")" "-" "+" "," 만 입력 허용
'---------------------------------------------------------------------------
sub gTelephoneInput ()
	dim intKeyVal
	
	intKeyVal = window.event.keyCode
	
	if (intKeyVal >= 48 and intKeyVal <= 57) or _
		intKeyVal = 40 or intKeyVal = 41 or intKeyVal = 45 or intKeyVal = 43 or intKeyVal = 44 then
		window.event.returnValue = true
	else
		window.event.returnValue = false
	end if
end sub

Sub window_onerror(message, url, line)
	msgbox "오류=" & message & chr(13) & chr(10) & _
	       "URL=" & url & chr(13) & chr(10) & _
	       "줄=" & line, vbCritical
	window.event.returnValue=true
End Sub

const XML_ROOT = "DataInfo"
const XML_PARENT= "Table"
const XML_FLAG = "DataFlag"

'=============================
'  XML 처리 공통 모듈
'=============================
Function gXMLCreateElement(ByVal xmlDoc, ByVal xmlParent, ByVal strTagName, ByVal strTagText)
    Dim xmlNewNode
    
    if strTagName="" then exit function

    '기존에 있는 경우는 취소
    if gXMLFindElement(xmlParent,strTagName) then exit function
    
    Set xmlNewNode = xmlDoc.createElement(strTagName)
    xmlParent.appendChild xmlNewNode
    xmlParent.lastChild.Text = strTagText
    
    Set gXMLCreateElement = xmlParent.lastChild
End Function

Sub gXMLSetElement(ByVal xmlDoc, ByVal strTagName, ByVal strTagText)
    Dim xmlNode
    
    Set xmlNode = xmlDoc.selectSingleNode(strTagName)
    If Not (xmlNode Is Nothing) And Not IsNull(xmlNode) Then
        xmlNode.Text = strTagText
    End If
End Sub

Function gXMLGetElement(ByVal xmlDoc, ByVal strTagName)
    Dim xmlNode
    
    Set xmlNode = xmlDoc.selectSingleNode(strTagName)
    If Not (xmlNode Is Nothing) And Not IsNull(xmlNode) Then
        gXMLGetElement = xmlNode.Text
    Else
        gXMLGetElement = ""
    End If
End Function

Function gXMLFindElement(ByVal xmlDoc, ByVal strTagName)
    Dim xmlNode
    
    Set xmlNode = xmlDoc.selectSingleNode(strTagName)
    If Not (xmlNode Is Nothing) And Not IsNull(xmlNode) Then
        gXMLFindElement = true
    Else
        gXMLFindElement = false
    End If
End Function

Function gXMLIsDataChanged (objXmlID)
	dim xmlRoot, strTag, strFlag
	
	set xmlRoot = objXmlID.XMLDocument.documentElement.firstChild
	'Data Flag 변경 여부 검사
	strTag = XML_FLAG
	strFlag = gXMLGetElement(xmlRoot,strTag)
	if strFlag = "" or cint(strFlag) = meCLS_FLAG or cint(strFlag) = meINS_FLAG then
		gXMLIsDataChanged = false
	else
		gXMLIsDataChanged = true
	end if
End Function

Sub gXMLDataChanged (objXmlID)
	dim xmlRoot, strTag, strFlag
	
	set xmlRoot = objXmlID.XMLDocument.documentElement.firstChild
	'Data Flag 변경
	strTag = XML_FLAG
	strFlag = gXMLGetElement(xmlRoot,strTag)
	if strFlag <> "" then gXMLSetElement xmlRoot,strTag, strFlag or meUPD_FLAG
End Sub

Sub gXMLSetFlag (objXmlID, byval intFlag)
	dim xmlRoot, strTag
	
	set xmlRoot = objXmlID.XMLDocument.documentElement.firstChild
	'Data Flag 설정
	strTag = XML_FLAG
	gXMLSetElement xmlRoot,strTag,intFlag
End Sub

Function gXMLGetFlag (objXmlID)
	dim xmlRoot, strTag
	
	set xmlRoot = objXmlID.XMLDocument.documentElement.firstChild
	'Data Flag 읽기
	strTag = XML_FLAG
	gXMLGetFlag = gXMLGetElement (xmlRoot,strTag)
End Function

Function gXMLGetFlagMode (objXmlID)
	dim intFlag

	intFlag = cint(gXMLGetFlag (objXmlID))
	
	 Select Case intFlag
		Case meINS_FLAG, meIU_FLAG
			gXMLGetFlagMode = meINS_TRANS
		Case meUPD_FLAG
			gXMLGetFlagMode = meUPD_TRANS
		Case meDEL_FLAG, meUD_FLAG
			gXMLGetFlagMode = meDEL_TRANS
		Case Else   'meID_FLAG, meIUD_FLAG
			gXMLGetFlagMode = meNO_TRANS
        End Select	
End Function

Function gXMLIsDeletable (objXmlID)
	dim intFlag

	intFlag = gXMLGetFlagMode (objXmlID)
	
	if intFlag = meINS_TRANS then
		gXMLIsDeletable = false
	else
		gXMLIsDeletable = true
	end if	
End Function

'=============================
'  XML을 이용한 Data Binding
'=============================
function gXMLNewBinding(form,objXmlID,strDataSrc)
	dim xmlDoc, xmlErr
	
	set xmlDoc = objXmlID.XMLDocument
	
	xmlDoc.async = False
	xmlDoc.loadXML "<" & XML_ROOT & ">" & "</" & XML_ROOT & ">"
	Set xmlErr = xmlDoc.parseError
	If (xmlErr.errorCode <> 0) Then
		MsgBox "You have error " & xmlErr.errorCode & "=" & xmlErr.reason, vbExclamation, "XML 오류"
		gXMLNewBinding = False
		Exit Function
	End If
    
	'-------------------------------------------
	' xml 문서의 Root Element 및 Parent Element 지정
	'-------------------------------------------
	Dim xmlRoot, xmlParent
	Set xmlRoot = xmlDoc.documentElement
	Set xmlParent = gXMLCreateElement(xmlDoc, xmlRoot, XML_PARENT,"")
    
	'------------------------------
	' xml 문서에 Flag정보 생성
	'------------------------------
	gXMLCreateElement xmlDoc,xmlParent,XML_FLAG,meINS_FLAG	'Insert Flag

	'---------------------------------------------
	' Binding된 모든 오브젝트에 대한 Element 생성
	'---------------------------------------------
	dim intCnt,i
	dim objElement
	intCnt = form.elements.length
	
	for i = 0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.datasrc=strDataSrc then
			gXMLCreateElement xmlDoc,xmlParent,objElement.datafld,objElement.value '""
		end if
	next

	gXMLNewBinding = true
	dim strXML
	strXML =objXmlID.XMLDocument.xml
	
	mblnErrDisp = false	'초기 데이터 바인딩시는 데이터 포맷 오류 무시
	call gXMLDataBinding (form,objXmlID,strDataSrc,strXml)
	mblnErrDisp = true
end function

function gXMLDataBinding (form,objXmlID,strDataSrc,strXml)
	dim xmlDoc
	
	gXMLDataBinding = false

	'데이터 바인딩 전에 각 엘리먼트 값 유형에 따라 포맷설정 
	'(중요)바인딩 후에 포맷 설정은 바인딩 종료 시점을 알 수 없어 불가능
	 if not gSetXMLDataFormat(form,strDataSrc,strXml) then exit function
	 
	'데이터를 화면의 엘리먼트와 바인딩
	set xmlDoc =  objXmlID.XMLDocument
	
	xmlDoc.async = False
	if  not xmlDoc.loadXML(strXml) then
		msgbox ("error:" + xmlDoc.parseError.reason)
		exit function
	end if
	
	'Flag Tag 생성 - Tag가 없는 경우는 meCLS_FLAG로 설정
	if gXMLGetFlag(objXmlID) = "" then
		gXMLCreateElement xmlDoc,xmlDoc.documentElement.firstChild,XML_FLAG,meCLS_FLAG
	end if	

	'(중요)XML ID와 컨트롤 과의 Sync를 위한 대기
	 if not gSetXMLDataSync(form,strDataSrc,strXml) then exit function

	set xmlDoc = nothing
	'결과반환
	gXMLDataBinding = true
end function

Function gXMLGetBindingData(objXmlID)
	gXMLGetBindingData = objXmlID.XMLDocument.xml	
End Function

Function gDataValidation (form)
	dim vntParams

	gDataValidation = false
	if not isobject(form) then
		msgbox "Form Object가 아닙니다!",vbExclamation
		 exit function
	end if

	dim intCnt,i
	dim objElement
	'on error resume next
	intCnt = form.elements.length
	
	for i = 0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.type="text" or objElement.type="select-one" or objElement.tagName = "TEXTAREA" then
			if objElement.accesskey <> "" then
				vntParams = split(objElement.accesskey,",")
				if ubound(vntParams) > 0 then
					if ucase(vntParams(1)) = "M" and objElement.value = "" then
						 err.clear
						 msgbox objElement.title & meSURELY_INPUT,vbExclamation
						 objElement.focus
						 exit function
					end if
				end if
			end if
		end if
	next
	err.clear
	gDataValidation = true
'		if form.elements(i).type="text" or form.elements(i).type="hidden" then
'			msgbox form.elements(i).type & "," & form.elements(i).name
'		elseif form.elements(i).type="radio" or form.elements(i).type="checkbox" then
'			msgbox form.elements(i).type & "," & form.elements(i).name
'		elseif form.elements(i).type="select-one" or form.elements(i).type="select-multiple" then
'			msgbox form.elements(i).type & "," & form.elements(i).name
'		end if
End Function

function gSetXMLDataSync (form,strDataSrc,strXml)
	dim intCnt,i
	dim objElement,xmlDoc,xmlParent
	dim strData, strTag
	
	if not isobject(form) then exit function	
	
	set xmlDoc = CreateObject("Microsoft.XMLDOM") '새로운 XMLDOM 생성
	xmlDoc.async = False

	'데이터를 XMLDOM으로 로딩
	if  not xmlDoc.loadXML(strXml) then
		msgbox ("error:" + xmlDoc.parseError.reason)
		gSetXMLDataSync = false
		set xmlDoc = nothing
		exit function
	end if

	set xmlParent = xmlDoc.documentElement.firstChild
	
	'각 오브젝트 포맷에 맞게 XML 데이터를 포맷팅
	intCnt = form.elements.length
	for i = 0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.datasrc=strDataSrc then
			'바인딩 Tag
			strTag = objElement.datafld
			
			if objElement.type="text" then
				strData = gXMLGetElement(xmlParent,strTag)

				'XML ID와 컨트롤 과의 Sync 검사
				do while trim(strData) <> trim(objElement.value)
					mobjSCGLCtl.DoEventQueue
				loop
			end if
		end if
	next
	gSetXMLDataSync = true
	set xmlDoc = nothing
end function

function gSetXMLDataFormat (form,strDataSrc,strXml)
	dim intCnt,i,intTypeOpt
	dim objElement,xmlDoc,xmlParent
	dim strData, strTag, blnFormat
	dim vntType, vntParams
	
	if not isobject(form) then exit function	
	
	set xmlDoc = CreateObject("Microsoft.XMLDOM") '새로운 XMLDOM 생성
	xmlDoc.async = False

	'데이터를 XMLDOM으로 로딩
	if  not xmlDoc.loadXML(strXml) then
		msgbox ("error:" + xmlDoc.parseError.reason)
		gSetXMLDataFormat = false
		set xmlDoc = nothing
		exit function
	end if

	set xmlParent = xmlDoc.documentElement.firstChild
	
	'각 오브젝트 포맷에 맞게 XML 데이터를 포맷팅
	intCnt = form.elements.length
	for i = 0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.datasrc=strDataSrc then
			'바인딩 Tag가 없는 경우는 Tag를 강제적으로 생성
			strTag = objElement.datafld
			if not gXMLFindElement(xmlParent,strTag) then gXMLCreateElement xmlDoc, xmlParent, strTag,""
			
			if objElement.type="text" then
				if objElement.accesskey <> "" then
					strData = gXMLGetElement(xmlParent,strTag)
					blnFormat = false
					vntType = split(objElement.accesskey,",")
					if vntType(0) <> "" then
						vntParams = split(vntType(0),":")
						select case ucase(vntParams(0))
							case "DATE"
								strData = gFormatDate(strData,true): blnFormat=true
				        	        case "TIME"
						                strData = gFormatTime(strData,true): blnFormat=true
							case "MON"
								strData = gFormatMonth(strData,true): blnFormat=true
							case "INT"
								strData = gFormatNumber(strData,0,true): blnFormat=true
							case "DEC"
								if ubound(vntParams) = 1 then intTypeOpt = vntParams(1) else intTypeOpt = 2
								strData = gFormatNumber(strData,intTypeOpt,true): blnFormat=true
						end select
						if blnFormat then gXMLSetElement xmlParent,strTag, strData
					end if
				end if
				'XML ID와 컨트롤 과의 Sync를 위해 클리어
				objElement.value=""
			elseif objElement.type="checkbox" and gblnCnvtChkbox then
				strData = ucase(gXMLGetElement(xmlParent,strTag))
				blnFormat = false
				if strData ="0" Or strData = "1" Or strData = "-1" Then
				elseif strData = "N" Or strData = "" Then
					strData = "0": blnFormat=true
				else
					strData = "-1": blnFormat=true
				end if
				if blnFormat then gXMLSetElement xmlParent,strTag, strData

			end if
		end if
	next
	strXml = xmlDoc.xml
	gSetXMLDataFormat = true
	set xmlDoc = nothing
end function

function gBindXML2Sheet (form,byval strDataSrc,objSht, byval blnToSheet, byval lngRow)
	dim intCnt, i 
	dim objElement,objXmlID
	dim strData, strTag, blnChecked
	
	gBindXML2Sheet = false
	if not isobject(form) then 
		msgbox "Form Object가 아닙니다!",vbExclamation
		exit function
	end if	
	
	'XML 데이터와 쉬트의 데이터를 바인딩
	intCnt = form.elements.length
	if lngRow <= 0 then lngRow = objSht.ActiveRow
	'XML Flag 복사
	'if left(strDataSrc,1) = "#" then
	'	set objXmlID = window.document.all(right(strDataSrc,len(strDataSrc)-1))
	'	if blnToSheet then	'XML --> Sheet
	'		mobjSCGLSpr.SetFlag objSht, gXMLGetFlag (objXmlID), lngRow
	'	else				'Sheet --> XML
	'		gXMLSetFlag objXmlID, mobjSCGLSpr.GetFlag (objSht, lngRow)
	'	end if
	'end if
	
	for i = 0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.datasrc=strDataSrc then
			'바인딩 Tag
			strTag = objElement.datafld
			if blnToSheet then	'XML --> Sheet
				if objElement.type="checkbox" then
					mobjSCGLSpr.SetTextBinding objSht, strTag, lngRow, cint(objElement.checked)
				else
					mobjSCGLSpr.SetTextBinding objSht, strTag, lngRow, objElement.value
				end if
			else				'Sheet --> XML
				if objElement.type="checkbox" then
						if mobjSCGLSpr.GetTextBinding(objSht, strTag, lngRow) = "Y"  OR _
						mobjSCGLSpr.GetTextBinding(objSht, strTag, lngRow) = "1" Then
							blnChecked = true
						else
							blnChecked = false
						end if	
					objElement.checked = blnChecked
				else
					objElement.value = mobjSCGLSpr.GetTextBinding(objSht, strTag, lngRow) 
				end if
				gSetDataFormat objElement
			end if
		end if
	next
	gBindXML2Sheet = true
end function

sub gSetFormDataFormat (form,strDataSrc)	
	dim intCnt,i
	dim objElement
	
	if not isobject(form) then exit sub	
	intCnt = form.elements.length
	
	for i = 0 to intCnt-1
		set objElement = form.elements(i)
		if strDataSrc = "" then		
			gSetDataFormat(objElement)
		elseif objElement.datasrc = strDataSrc then
			gSetDataFormat(objElement)
		end if
	next
end sub

sub gSetDataFormat (byval objElement)
	dim vntType,vntParams, intTypeOpt
	
	if not isobject(objElement) then exit sub
	if objElement.tagName <> "INPUT" then exit sub		
	if objElement.type="text" then
		if objElement.value <> "" and objElement.accesskey <> "" then
			vntType = split(objElement.accesskey,",")
			if vntType(0) <> "" then
				vntParams = split(vntType(0),":")
				select case ucase(vntParams(0))
					case "DATE"
						objElement.value = gFormatDate(objElement,true)
					case "TIME"
						objElement.value = gFormatTime(objElement,true)
					case "MON"
						objElement.value = gFormatMonth(objElement,true)
					case "INT"
						objElement.value = gFormatNumber(objElement.value,0,true)
					case "DEC"
						if ubound(vntParams) = 1 then intTypeOpt = vntParams(1) else intTypeOpt = 2
						objElement.value = gFormatNumber(objElement.value,intTypeOpt,true)
				end select
			end if
		end if
	end if

'		if form.elements(i).type="text" or form.elements(i).type="hidden" then
'			msgbox form.elements(i).type & "," & form.elements(i).name
'		elseif form.elements(i).type="radio" or form.elements(i).type="checkbox" then
'			msgbox form.elements(i).type & "," & form.elements(i).name
'		elseif form.elements(i).type="select-one" or form.elements(i).type="select-multiple" then
'			msgbox form.elements(i).type & "," & form.elements(i).name
'		end if

end sub

'------------------------------------------------------------------------
'입  력: blnInput = 오브젝트의 입력 또는 선택 가능 여부
'기  능: 오브젝터의 입력 스타일 제어 
'------------------------------------------------------------------------
Sub gSetInputStyle (objElement, blnInput)
	dim strClass
	
	on error resume next
	if objElement.tagName = "IMG" then
		gSetAuthority objElement, blnInput
		exit sub
	end if
	
	if objElement.type="text" or objElement.tagName = "TEXTAREA" then
		strClass = replace (objElement.className,"NO","")
		if blnInput then
			objElement.className=strClass
			objElement.readonly = false
			if objElement.tabindex = -1 then objElement.tabindex = 0 else objElement.tabindex=abs(objElement.tabindex)
		else
			objElement.className="NO" & strClass
			objElement.readonly = true
			if objElement.tabindex = 0 then objElement.tabindex = -1 else objElement.tabindex=objElement.tabindex*-1
		end if
	elseif objElement.type = "select-one" or objElement.type = "select-multiple" or _
		objElement.type="radio" or objElement.type="checkbox" then		   				
		if blnInput then
			objElement.disabled=false
		else
			objElement.disabled=true
		end if		 
	end if
	err.clear
End Sub

sub gInitPageControl (form)
	dim intCnt,i
	dim strClass
	dim objElement
	dim vntParams

	if  form is nothing then exit sub

	intCnt = form.elements.length
	
	on error resume next
	'Title설정
	'if gstrTitle <> "" then gWriteText tdTitleName, gstrTitle
	'각 컨트롤 설정
	for i = 0 to intCnt-1
		set objElement =  form.elements(i)
		if objElement.type = "text" or objElement.type = "select-one" or _
		   objElement.type = "select-multiple" or objElement.type="radio" or objElement.type="checkbox" or _
		   objElement.tagName = "TEXTAREA" then
		   	'비입력모드설정
		   	if left(objElement.className,2) = "NO" then gSetInputStyle objElement, false
		   	
		   	'한영모드 변환
			if objElement.accesskey <> "" then
				vntParams = split(objElement.accesskey,",")
				if ubound(vntParams) > 1 then
					select case ucase(vntParams(2))
						case "H"	'한글
							objElement.style.IMEMode = "active"
						case "E"	'영문
							objElement.style.IMEMode = "inactive"
					end select
				end if
			end if		   	
		end if
	next
	err.clear
end sub
	
sub gClearAllObject (form)
	dim intCnt,i
	intCnt = form.elements.length
	
	on error resume next
	for i = 0 to intCnt-1
		if form.elements(i).type <> "button" and form.elements(i).type <> "select-one" and _
		  form.elements(i).type <> "select-multiple" and form.elements(i).tagName <> "OBJECT"  _
		  then form.elements(i).value=""
	next
	err.clear
end sub

sub gClearChildObject (form, strDataSrc)
	dim intCnt, i, objElement 
	intCnt = form.elements.length
	
	on error resume next
	for i = 0 to intCnt-1
		set objElement = form.elements(i)
		if objElement.datasrc = strDataSrc then
			if objElement.type <> "button" and objElement.type <> "select-one" and _
			  objElement.type <> "select-multiple" and objElement.tagName <> "OBJECT" _
			  then objElement.value=""
		end if
	next
	err.clear
end sub

'---------------------------------------------------------------------------
' 현재 날짜를 Return 받는 함수
' Application Server 로 부터 현재 날짜를 Return 받는다.
'---------------------------------------------------------------------------
Function gNowDate()
	 gNowDate = gFormatDate("<%=Now%>",true)
End Function

Function gNowDate2()
	dim NowTime
	NowTime = "<%=Now%>"
	
	if DatePart("d", NowTime) < 8 then
		NowTime = DateAdd("m",-1,NowTime)
		gNowDate2 = gFormatDate(NowTime,true)
	else
		gNowDate2 = gFormatDate(NowTime,true)
	end if
End Function

'---------------------------------------------------------------------------
' 입력 : 오브젝트 ID, 문자열 또는 DateTime형식을 값(예:now)
'	   blnErrMsg=True이면 오류메시지를 표시하고 해당 오브젝트로 이동
' 기능 : 표준 일자형식으로 변환. ""이면 오류
'---------------------------------------------------------------------------
function gFormatDate (byval objDateTime, blnErrMsg)
	dim strDateTime, strDateTimeOrg

	if isobject(objDateTime) then strDateTime = objDateTime.value else strDateTime = objDateTime 
	if strDateTime = "" then gFormatDate="" : exit function
	strDateTimeOrg = strDateTime

	if isdate(strDateTime) then	'정상적인 형태인 경우 YYYY-MM-DD or YYYY/MM/DD or DateTime Format
		
		if Mid(strDateTime,4,1) = "-" Then
			Dim rtnMsg : rtnMsg = gYesNoMsgbox("현재 사용자께서는 "& Mid(strDateTime,1,3) & " 년도를 입력하셨습니다." & vbcrlf & _
										 	 "계속하시겠습니까?", "날짜타입 경고" )
			If 	rtnMsg = vbNo Then
				gFormatDate = ""
			end if
				exit function
		end if
		
		gFormatDate = replace(FormatDateTime(strDateTime,vbShortDate),"/","-")
		
		if Mid(gFormatDate,1,2) < 19 Then
			gErrorMsgBox  "사용자 PC의 국가 및 언어 옵션이 한국이 아닙니다." & vbcrlf & _ 
		     	          "사용자 PC의 제어판 -> 국가 및 언어 옵션을 한국으로 설정해 주시고 " & vbcrlf & _
		     	          "날짜타입을 YYYY-MM-DD 로 변경하여 주시기 바랍니다." & vbcrlf & _ 
						  "현재 사용자께서 선택하신 날짜타입의 날짜는 "& gFormatDate &" 입니다.", "날짜타입경고"
			gFormatDate = ""
		end if
		exit function
		
	end if
	'다른 형식일 경우 (오라클 예: 2003-04-05T11:34:45.0000???)
	strDateTime = Mid(strDateTime,1,10)
	'혹시 잘못입력된 경우를 거름
	strDateTime = replace(strDateTime,"/","") : strDateTime = replace(strDateTime,"-","")
	if isnumeric(strDateTime) then
		select case Len(strDateTime)
			case 8	'YYYYMMDD
				strDateTime = Mid(strDateTime,1,4) & "-" & Mid(strDateTime,5,2) & "-" & Mid(strDateTime,7,2)
			case 4	'MMDD
				strDateTime = Year(now) & "-" & Mid(strDateTime,1,2) & "-" & Mid(strDateTime,3,2)
			case 3	'MDD
				strDateTime = Year(now) & "-0" & Mid(strDateTime,1,1) & "-" & Mid(strDateTime,2,2)
			case 2	'DD
				strDateTime = Year(now) & "-" & Mid(FormatDateTime(now,vbShortDate),6,2) & "-" & strDateTime
			case 1	'D
				strDateTime = Year(now) & "-" & Mid(FormatDateTime(now,vbShortDate),6,2) & "-0" & strDateTime
			case else	'오류
				strDateTime = ""
		end select
	end if
	
	'다시 검사
	if not isdate(strDateTime)  and blnErrMsg and mblnErrDisp then 
		gErrorMsgBox strDateTimeOrg & "은(는) 잘못된 날짜형식입니다!",""
		strDateTime=""
	end if
	if isobject(objDateTime) then objDateTime.value=strDateTime
	gFormatDate = strDateTime
end function

'---------------------------------------------------------------------------
' 입력 : 오브젝트 ID, 문자열 또는 DateTime/Time형식을 값(예:now)
'	   blnErrMsg=True이면 오류메시지를 표시하고 해당 오브젝트로 이동
' 기능 : 표준 시간형식으로 변환. ""이면 오류
'---------------------------------------------------------------------------
function gFormatTime (byval objDateTime, blnErrMsg)
	dim strDateTime, strDateTimeOrg

	if isobject(objDateTime) then strDateTime = objDateTime.value else strDateTime = objDateTime 
	if strDateTime = "" then gFormatTime="" : exit function
	strDateTimeOrg = strDateTime
	
	if isdate(strDateTime) then	'정상적인 형태인 경우 HH:MM
		gFormatTime = FormatDateTime(strDateTime,vbShortTime)
		exit function
	end if
	'혹시 잘못입력된 경우를 거름
	strDateTime = replace(strDateTime,":",""): strDateTime = replace(strDateTime," ","")
	if isnumeric(strDateTime) then
		select case Len(strDateTime)
			case 4	'HHMM
				strDateTime = Mid(strDateTime,1,2) & ":" & Mid(strDateTime,3,2)
			case else	'오류
				strDateTime = ""
		end select
	end if
	'다시 검사
	if not isdate(strDateTime)  and blnErrMsg and mblnErrDisp then 
		gErrorMsgBox strDateTimeOrg & "은(는) 잘못된 시간형식입니다!(00:00~23:59)",""
		strDateTime=""
	end if
	if isobject(objDateTime) then objDateTime.value=strDateTime
	gFormatTime = strDateTime
end function

function gFormatMonth (byval objDateTime, blnErrMsg)
	dim strDateTime, strDateTimeOrg

	if isobject(objDateTime) then strDateTime = objDateTime.value else strDateTime = objDateTime
	if strDateTime = "" then gFormatMonth="" : exit function
	strDateTimeOrg = strDateTime
	
	'혹시 잘못입력된 경우를 거름
	strDateTime = replace(strDateTime,"/","") : strDateTime = replace(strDateTime,"-","")
	if isnumeric(strDateTime) then
		select case Len(strDateTime)
			case 6	'YYYYMM
				strDateTime = Mid(strDateTime,1,4) & "-" & Mid(strDateTime,5,2)
			case 8  'YYYYMMDD
				strDateTime = Mid(strDateTime,1,4) & "-" & Mid(strDateTime,5,2)
			case else	'오류
				strDateTime = ""
		end select
	end if
	
	'검사	
	if not isdate(strDateTime & "-" & "01")  and blnErrMsg and mblnErrDisp then 
		gErrorMsgBox strDateTime & "은(는) 잘못된 년월형식입니다!",""
		strDateTime=""
	end if
	if isobject(objDateTime) then objDateTime.value=strDateTime
	gFormatMonth = strDateTime
end function

'---------------------------------------------------------------------------
' 입력 : 오브젝트 ID, 문자열 또는 숫자형식을 값 
' 기능 : 1000단위가 삽입된 형태의 숫자
'---------------------------------------------------------------------------
function gFormatNumber (byval objNum, byval intDecLen, blnErrMsg)
	dim strNum

	if isobject(objNum) then strNum = objNum.value else strNum = objNum 
	if strNum = "" then gFormatNumber="0" : exit function
	
	if isnumeric(strNum) then	'정상적인 형태인 경우
		strNum = formatnumber(strNum,intDecLen)
	else
		if blnErrMsg and mblnErrDisp then _ 
			gErrorMsgBox strNum & "은(는) 잘못된 숫자형식입니다!",""
		strNum = "0"
	end if
	if isobject(objNum) then objNum.value=strNum
	gFormatNumber = strNum
end function

'---------------------------------------------------------------------------
' 입력 : 반올림대상값, 자리수
' 기능 : gRound(12.25, 1) ==> 12.3
'---------------------------------------------------------------------------
Function gRound(ByVal xNumber, ByVal xPosition)
    Dim intX, intPositionNum 
    If IsNumeric(xNumber) and IsNumeric(xPosition) then 
        intPositionNum = 10^xPosition 
        intX = Int(xNumber * intPositionNum + 0.5) / intPositionNum
        gRound = intX 
    Else 
        gRound = xNumber
    End If 
End function

'---------------------------------------------------------------------------
' 입력 : 오브젝트 ID, 문자열 또는 DateTime형식을 값(예:now) 
' 기능 : 올바른 일자인지 검사. True or False
'---------------------------------------------------------------------------
function gCheckDate (byval objDate, blnErrMsg)
	gCheckDate = isdate(gFormatDate(objDate,blnErrMsg))
end function

'---------------------------------------------------------------------------
' 입력 : 오브젝트 ID, 문자열 또는 DateTime형식을 값(예:now) 
' 기능 : YYYYMMDD 형태로 반환
'---------------------------------------------------------------------------
function gDateToStr (byval objDate)
	dim strDateTime

	if isobject(objDate) then strDateTime = objDate.value else strDateTime = objDate 
	
	strDateTime = replace(strDateTime,"/","") : strDateTime = replace(strDateTime,"-","")
	gDateToStr = strDateTime
end function

Function gCalcDateDiff (objFromDate, objToDate)
	Dim strFromDate, strToDate
	Dim intDiff
		
	gCalcDateDiff =-1
	if isobject(objFromDate) then strFromDate = objFromDate.value else strFromDate = objFromDate
	if isobject(objToDate) then strToDate = objToDate.value else strToDate = objToDate
	'날짜 검사
	if not gCheckDate(strFromDate,true) then exit function
	if not gCheckDate(strToDate,true) then exit function

	intDiff = DateDiff("d", CDate(strFromDate), CDate(strToDate))
	if intDiff < 0 then gErrorMsgBox strFromDate & "가 " & strToDate & "보다 클 수 없습니다!",""
	
	gCalcDateDiff =intDiff 	
End Function

Function gCleanField (ByVal objField, ByVal objField1)
	if isobject(objField) then objField.value = ""
	if isobject(objField1) then objField1.value = ""

end Function

Function gYearmonCheck (ByVal objField)
	Dim strResult
	if isobject(objField) then
		if len(objField.value) = 6 then
			strResult = REPLACE(MID(objField.value,5,2),"0","")
			if strResult > 12 Then
				MsgBox "년월의 입력형식이 잘못되었습니다.", vbExclamation, "입력오류!"
				objField.value = ""
			End If
		End if 
	end if
end Function

Function gLpad (MyValue, MyPadChar, MyPaddedLength) 
	gLpad = string(MyPaddedLength - Len(MyValue),MyPadChar) & MyValue 
End Function

-->
</SCRIPT>

<script LANGUAGE="javascript">

function gShowModalWindowJ (strPageURL, vntInParams, lngDialogWidth, lngDialogHeight) 
{
	var vntRet

	vntRet= window.showModalDialog (strPageURL,vntInParams,"dialogWidth=" + lngDialogWidth + "px;dialogHeight=" + lngDialogHeight + "px;status=no;center=yes;border=thin;resize=yes;statusbar=yes;")
	return vntRet;
}

document.oncontextmenu = function () {
// 오른쪽 마우스 버튼을 눌렀을때에 동작하기를 원하는 코드를 넣으세요.
//return false;
}                          

document.onkeydown = function () {
// 페이지의 새로고침 불가능
    if(event.keyCode == 116)
    {
        event.keyCode= 2;
        return false;
    }
    
    else if ((event.keyCode == 78) && (event.ctrlKey == true))
    {        
        event.keyCode = 2;
        return false;
    }
 }
 
function getOffsetLeft (el) {
    var ol = el.offsetLeft;
    while ((el = el.offsetParent) != null)
        ol += el.offsetLeft;
    return ol;
}

function getOffsetTop (el) {
    var ot = el.offsetTop;
    while((el = el.offsetParent) != null)
        ot += el.offsetTop;
    return ot;
}


</SCRIPT>