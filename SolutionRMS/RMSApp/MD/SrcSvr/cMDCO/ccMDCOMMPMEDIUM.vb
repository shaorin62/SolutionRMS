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
Imports eMDCO                       '엔터티 추가

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트

Public Class ccMDCOMMPMEDIUM
    Inherits ccControl
#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccMDCOMMPMEDIUM"                  '자신의 클래스명
    Private mobjceMD_MMP_MEDIUMT As eMDCO.ceMD_MMP_MEDIUM '사용할 Entity 변수 선언
#End Region

#Region "GROUP BLOCK : 외부에비공개"
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCLIENTNAME As String) As Object

        Dim strSQL, strFormat, strWhere, strSelFields, strKeys As String
        Dim vntData As Object
        Dim Con1, Con2, Con3 As String

        Con1 = "" : Con2 = "" : Con3 = ""

        SetConfig(strInfoXML)   '기본정보 설정

        With mobjSCGLConfig

            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

            strWhere = BuildFields(" ", Con1, Con2, Con3)

            strFormat = " SELECT "
            strFormat = strFormat & " 0 CHK, YEARMON, SEQ, "
            strFormat = strFormat & " CLIENTCODE, "
            strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
            strFormat = strFormat & " REAL_MED_CODE,"
            strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,"
            strFormat = strFormat & " DEPT_CD,"
            strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME,"
            strFormat = strFormat & " DEMANDDAY,"
            strFormat = strFormat & " ISNULL(AMT,0) AMT,"
            strFormat = strFormat & " DBO.SC_MMPRATE_FUN(YEARMON) RATE,"
            strFormat = strFormat & " ISNULL(MMP_AMT,0) MMP_AMT,"
            strFormat = strFormat & " ISNULL(CONFIRMFLAG,'N') CONFIRMFLAG,"
            strFormat = strFormat & " VOCHNO"
            strFormat = strFormat & " FROM MD_MMP_MEDIUM"
            strFormat = strFormat & " WHERE 1=1 "
            strFormat = strFormat & " {0} "
            strFormat = strFormat & " ORDER BY SEQ "

            '데이터 조회
            Try
                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    Public Function SelectRtn_rate(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String) As Object

        Dim strSQL, strFormat As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig


            strFormat = " SELECT RATE"
            strFormat = strFormat & " FROM ("
            strFormat = strFormat & " 	SELECT TOP 1 RATE"
            strFormat = strFormat & " 	FROM SC_MMPRATE_MST"
            strFormat = strFormat & " 	WHERE "
            strFormat = strFormat & " 	("
            strFormat = strFormat & " 		(ISNULL(SUBSTRING(FDATE,1,6),'') <= '" & strYEARMON & "' AND ISNULL(SUBSTRING(EDATE,1,6),'') >= '" & strYEARMON & "') OR"
            strFormat = strFormat & " 		(ISNULL(SUBSTRING(FDATE,1,6),'') <= '" & strYEARMON & "' AND ISNULL(SUBSTRING(EDATE,1,6),'') = '')"
            strFormat = strFormat & " 	)"
            strFormat = strFormat & " 	ORDER BY CDATE DESC"
            strFormat = strFormat & " )A"

            '데이터 조회
            Try
                strSQL = String.Format(strFormat)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_rate")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ''============== ProcessRtn
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Object '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblSEQ As Integer
        Dim strDEMANDDAY As String


        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceMD_MMP_MEDIUMT = New ceMD_MMP_MEDIUM(mobjSCGLConfig)
                    'vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        If GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True) = -999999 Then
                            If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = Mid(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), 1, 4) & Mid(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), 6, 2) & Mid(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), 9, 2)
                            intRtn = InsertRtn(vntData, intColCnt, i, strDEMANDDAY)

                        Else
                            dblSEQ = GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True)

                            If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = Mid(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), 1, 4) & Mid(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), 6, 2) & Mid(GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR), 9, 2)

                            intRtn = UpdateRtn(vntData, intColCnt, i, dblSEQ, strDEMANDDAY)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_MMP_MEDIUMT.Dispose()
            End Try
        End With
    End Function

    ' =============== DeleteRtn 
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal strSEQ As Integer, _
                              ByVal strYEARMON As String) As Integer   '데이터 DELETE

        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceMD_MMP_MEDIUMT = New ceMD_MMP_MEDIUM(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceMD_MMP_MEDIUMT.DeleteDo(strYEARMON, strSEQ)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceMD_MMP_MEDIUMT.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function InsertRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strDEMANDDAY As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceMD_MMP_MEDIUMT.InsertDo( _
                                       GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "REAL_MED_CODE", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "AMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "RATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MMP_AMT", intColCnt, intRow, NULL_NUM, True), _
                                       "N", _
                                       GetElement(vntData, "VOCHNO", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR010", intColCnt, intRow, NULL_NUM, True))

                       

        Return intRtn
    End Function


    Private Function UpdateRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal dblSEQ As Double, _
                               ByVal strDEMANDDAY As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceMD_MMP_MEDIUMT.UpdateDo( _
                                       GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                       dblSEQ, _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "REAL_MED_CODE", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "AMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "RATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MMP_AMT", intColCnt, intRow, NULL_NUM, True), _
                                       "N", _
                                       GetElement(vntData, "VOCHNO", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR010", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function
#End Region
End Class

