'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 컨트롤 클래스 메이커
'시스템구분    : 솔루션명 /시스템명/Server Control Class
'실행   환경    : COM+ Service Server Package
'프로그램명    : ccMDCMDEPTMST.vb
'기         능    : - 기능을 명시 합니다.
'특이  사항     : - 특이사항에 대해 표현
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 
'            2) 
'****************************************************************************************

Imports System.Xml                  ' XML처리
Imports SCGLControl                 ' ControlClass의 Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타유틸리티 클래스
Imports eSCCO                       '엔터티 추가

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트

Public Class ccSCCOEMPMST
    Inherits ccControl
#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccSCCOEMPMST"                  '자신의 클래스명
    Private mobjceSC_EMPLOYEE_MST As eSCCO.ceSC_EMPLOYEE_MST             '사용할 Entity 변수 선언
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.88;Initial Catalog=MCDEV;DSN=MCDEV;UID=advsa;Pwd=advsa1234" '커넥션Setting
#End Region

#Region "GROUP BLOCK : 외부에비공개"
    Public Function GetEMP(ByVal strInfoXML As String, _
                          ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                          ByVal strCODE As String, _
                          ByVal strNAME As String, _
                          ByVal strGUBUN As String, _
                          ByVal strDEPTCD As String, _
                          ByVal strDEPTNAME As String, _
                          ByVal strYN As String) As Object
        Dim strSQL, strFormat, strWhere As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        Dim Con1, Con2, Con3, Con4, Con5, Con6 As String
        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig
            If Len(strCODE) = 5 Then
                strCODE = "000" & strCODE
            End If

            '한글인 경우
            If strCODE <> "" Then Con1 = String.Format(" AND (EMPNO = '{0}')", strCODE)
            If strNAME <> "" Then Con2 = String.Format(" AND EMP_NAME LIKE '%{0}%'", strNAME)
            If strGUBUN <> "A" Then Con3 = String.Format(" AND SC_EMP_STATUS = '{0}'", strGUBUN)
            If strDEPTCD <> "" Then Con4 = String.Format(" AND (CC_CODE = '{0}')", strDEPTCD)
            If strDEPTNAME <> "" Then Con5 = String.Format(" AND DBO.SC_DEPT_NAME_FUN(CC_CODE) LIKE '%{0}%'", strDEPTNAME)
            If strYN <> "A" Then Con6 = String.Format(" AND A.USE_YN = '{0}'", strYN)
            '조회 필드 설정

            strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6)

            strFormat = strFormat & " SELECT "
            strFormat = strFormat & " A.EMPNO, A.EMP_NAME, "
            strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(B.HIGHDEPT_CD) HIGHCC_NAME, "
            strFormat = strFormat & " A.CC_CODE, DBO.SC_DEPT_NAME_FUN(A.CC_CODE) CC_NAME, "
            strFormat = strFormat & " CASE A.SC_EMP_STATUS WHEN '0' THEN '재직' WHEN '1' THEN '휴직' WHEN '3' THEN '퇴직' END SC_EMP_STATUS, "
            strFormat = strFormat & " CASE A.USE_YN WHEN 'Y' THEN '1' ELSE '0' END AS USE_YN, "
            strFormat = strFormat & " A.PASSWORD, A.MANAGER, "
            strFormat = strFormat & " A.TITLENAME, A.BIRTH, A.E_MAIL, A.TEL, A.CELLPHONE "
            strFormat = strFormat & " FROM SC_EMPLOYEE_MST A"
            strFormat = strFormat & " LEFT JOIN SC_CCTR B ON A.CC_CODE = B.DEPT_CD AND B.TDATE <> '' "
            strFormat = strFormat & " WHERE 1=1  {0} "
            strFormat = strFormat & " ORDER BY CC_CODE "

            strSQL = String.Format(strFormat, strWhere)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblID As Double '자동 ID 를사용할 때만 사용
        Dim strSC_EMP_STATUS As String
        Dim strUSE_YN, strMANAGER As String
        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_EMPLOYEE_MST = New ceSC_EMPLOYEE_MST(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''해당하는Row 만큼 Loop
                    strSC_EMP_STATUS = ""


                    For i = 1 To intRows
                        If GetElement(vntData, "SC_EMP_STATUS", intColCnt, i) = "재직" Then
                            strSC_EMP_STATUS = "0"
                        ElseIf GetElement(vntData, "SC_EMP_STATUS", intColCnt, i) = "휴직" Then
                            strSC_EMP_STATUS = "1"
                        ElseIf GetElement(vntData, "SC_EMP_STATUS", intColCnt, i) = "퇴직" Then
                            strSC_EMP_STATUS = "3"
                        End If

                        If GetElement(vntData, "USE_YN", intColCnt, i) = "1" Then
                            strUSE_YN = "Y"
                        ElseIf GetElement(vntData, "USE_YN", intColCnt, i) = "0" Then
                            strUSE_YN = "N"
                        End If

                        If GetElement(vntData, "MANAGER", intColCnt, i) = "1" Then
                            strMANAGER = "Y"
                        ElseIf GetElement(vntData, "MANAGER", intColCnt, i) = "0" Then
                            strMANAGER = "N"
                        End If

                        intRtn = UpdateRtn(vntData, intColCnt, i, strSC_EMP_STATUS, strUSE_YN, strMANAGER)

                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_EMPLOYEE_MST.Dispose()

            End Try
        End With
    End Function

    Public Function ProcessRtn_PWDCHANGE(ByVal strInfoXML As String, _
                                         ByVal strEMPNO As String) As Integer

        Dim i, intColCnt, intRows As Integer
        Dim strSQL
        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                mobjceSC_EMPLOYEE_MST = New ceSC_EMPLOYEE_MST(mobjSCGLConfig)

                If strEMPNO <> "" Then
                    strSQL = String.Format("UPDATE SC_EMPLOYEE_MST SET PASSWORD='{0}' WHERE EMPNO = '{1}'", strEMPNO, strEMPNO)
                End If

                .mobjSCGLSql.SQLDo(strSQL)
                .mobjSCGLSql.SQLCommitTrans()
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_PWDCHANGE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_EMPLOYEE_MST.Dispose()
            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function UpdateRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strSC_EMP_STATUS As String, _
                               ByVal strUSE_YN As String, _
                               ByVal strMANAGER As String) As Integer
        Dim intRtn As Integer
        'EMPNO|EMP_NAME|CC_CODE|CC_NAME|SC_EMP_STATUS|E_MAIL|TEL|CELLPHONE|PASSWORD
        intRtn = mobjceSC_EMPLOYEE_MST.UpdateDo( _
                                       GetElement(vntData, "EMPNO", intColCnt, intRow), _
                                       GetElement(vntData, "EMP_NAME", intColCnt, intRow), _
                                       GetElement(vntData, "CC_CODE", intColCnt, intRow), _
                                       strSC_EMP_STATUS, _
                                       strUSE_YN, _
                                       GetElement(vntData, "PASSWORD", intColCnt, intRow), _
                                       strMANAGER)
        Return intRtn
    End Function
#End Region
End Class
