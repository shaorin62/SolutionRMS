'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 컨트롤 클래스 메이커 - 한화 S&C
'시스템구분    : 솔루션명 /시스템명/Server Control Class
'실행   환경    : COM+ Service Server Package
'프로그램명    : ccPDCMTRANS.vb
'기         능    : - 기능을 명시 합니다.
'특이  사항     : - 특이사항에 대해 표현
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2004-03-30 오전 10:32:13 By MakeSFARV.2.0.0
'                  2) 2004-03-30 오전 10:32:13 By 작성자명을 씁니다.
'****************************************************************************************

Imports System.Xml                        ' XML처리
Imports SCGLControl                       ' ControlClass의 Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr            '오류처리 클래스
Imports SCGLUtil.cbSCGLXml           'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil            '기타유틸리티 클래스
Imports eMDCO
Imports System.Math

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트

Public Class ccMDETELECCOMMILIST
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccMDETELECCOMMILIST"                  '자신의 클래스명
    Private mobjceMD_TRANS_TEMP As eMDCO.ceMD_TRANS_TEMP            '사용할 Entity 변수 선언
    Private mobjceMD_ELECCOMMI_HDR As eMDCO.ceMD_ELECCOMMI_HDR
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password" '커넥션Setting
#End Region

#Region "GROUP BLOCK : Property 선언"
#End Region

#Region "GROUP BLOCK : Event 선언"
    'VAT 계산   
    Public Function gRound(ByVal xNumber As Double, ByVal xPosition As Double) As Double
        Dim intX, intPositionNum
        If IsNumeric(xNumber) And IsNumeric(xPosition) Then
            intPositionNum = 10 ^ xPosition
            intX = Int(xNumber * intPositionNum + 0.5) / intPositionNum
            gRound = intX
        Else
            gRound = xNumber
        End If
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 공개 Method"
    'datacnt 를 계산하기 위한 함수
    Public Function Get_ELECCOMMI_CNT(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strTRANSYEARMON As String, _
                                       ByVal strTRANSNO As String) As Object

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strSelFields As String      '조회필드
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""

            If strTRANSYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strTRANSYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)

            strWhere = BuildFields(" ", Con1, Con2)
            strFormat = "SELECT count(*) FROM MD_ELECCOMMI_DTL WHERE 1=1 {0} "
            strSQL = String.Format(strFormat, strWhere)
            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_ELECCOMMI_CNT")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_ELECCOMMI_ALLLIST(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                      ByVal strYEARMON As String, _
                                      ByVal strTRANSNO As String, _
                                      ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strSelFields As String      '조회필드
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT 0 CHK ,TRANSYEARMON,  TRANSNO, CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, DBO.SC_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE,DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,AMT, VAT, (AMT+VAT) SUMAMTVAT,DEMANDDAY,PRINTDAY,dbo.MD_COMMITAX_YN2_FUN(TRANSYEARMON,TRANSNO,'A') TAXYN FROM MD_ELECCOMMI_HDR WHERE 1=1 and (attr03 <> 'Y' OR ATTR03 IS NULL)  {0} ORDER BY TRANSNO "
            strSQL = String.Format(strFormat, strWhere)
            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_ELECCOMMI_HDR(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As String

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strSelFields As String      '조회필드
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,  DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT, (AMT+VAT) SUMAMTVAT  FROM MD_ELECCOMMI_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_ELECCOMMI_LIST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strSelFields As String      '조회필드
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, MEDCODE, DBO.SC_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, REAL_MED_CODE,DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, AMT, SUSURATE, SUSU, VAT, CASE MED_FLAG WHEN '01' THEN 'TV' WHEN '02' THEN 'RADIO' WHEN '10' THEN 'DMB' END  MED_NAME FROM MD_ELECCOMMI_DTL WHERE 1=1 {0} ORDER BY TRANSNO "
            strSQL = String.Format(strFormat, strWhere)
            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function



    ' ============== ProcessRtn (Master & Detail) Sample Code 
    Public Function ProcessRtn_TEMP(ByVal strInfoXML As String, _
                                    ByVal strTRANSYEARMON As String, _
                                    ByVal strTRANSNO As String, _
                                    ByRef datacnt As String, _
                                    ByRef strUSERID As String) As Integer
        Dim intRtn As Integer  '결과값 변수
        Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수
        Dim intCnt

        SetConfig(strInfoXML) '기본정보 Setting

        With mobjSCGLConfig '기본정보를 가지고 있는 Config 개체
            Try
                'DB접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                'Master 데이터 처리
                mobjceMD_TRANS_TEMP = New ceMD_TRANS_TEMP(mobjSCGLConfig)

                strUSERID = .WRKUSR 'USERID 를 받아야 출력시에 USERID에 따른 구분을 둘수있다.
                intRtn = InsertRtn_TRANS_TEMP(strTRANSYEARMON, strTRANSNO, datacnt)
                mobjceMD_TRANS_TEMP.Dispose()
                '트랜잭션Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_TEMP")
            Finally
                'Resource해제

                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== DeleteRtn Sample Code
    Public Function DeleteRtn_temp(ByVal strInfoXML As String) As Integer   '데이터 DELETE

        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceMD_TRANS_TEMP = New ceMD_TRANS_TEMP(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceMD_TRANS_TEMP.DeleteDo(.WRKUSR)
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
                mobjceMD_TRANS_TEMP.Dispose()
            End Try
        End With
    End Function
    '수수료 세금계산서 일괄 삭제
    Public Function Delete_TRANS(ByVal strInfoXML As String, _
                                           ByVal strTAXYEARMON As String) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strSQL, strSQL2, strSQL3, strSQL4 As String

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                mobjceMD_ELECCOMMI_HDR = New ceMD_ELECCOMMI_HDR(mobjSCGLConfig)

                '''vntData의 컬럼수, 로우수를 변수입력
                strSQL = "DELETE FROM MD_ELECCOMMI_HDR WHERE TRANSYEARMON = '" & strTAXYEARMON & "'"
                strSQL2 = "DELETE FROM MD_ELECCOMMI_DTL WHERE TRANSYEARMON = '" & strTAXYEARMON & "'"
                strSQL3 = "UPDATE  MD_ELECTRIC_MEDIUM SET COMMI_TRANS_NO = '' WHERE YEARMON = '" & strTAXYEARMON & "'"
                strSQL4 = "UPDATE  MD_ELECTRIC_SUSUTEMP SET ATTR01 = '' WHERE YEARMON = '" & strTAXYEARMON & "'"
                intRtn = mobjceMD_ELECCOMMI_HDR.DeleteTRANS(strSQL)
                intRtn = mobjceMD_ELECCOMMI_HDR.DeleteTRANS(strSQL2)
                intRtn = mobjceMD_ELECCOMMI_HDR.DeleteTRANS(strSQL3)
                intRtn = mobjceMD_ELECCOMMI_HDR.DeleteTRANS(strSQL4)

                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".Delete_TRANS")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_ELECCOMMI_HDR.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"

    Private Function InsertRtn_TRANS_TEMP(ByVal strTRANSYEARMON As String, _
                                         ByRef strTRANSNO As String, _
                                         ByRef datacnt As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceMD_TRANS_TEMP.InsertDo( _
                                       strTRANSYEARMON, _
                                       strTRANSNO, _
                                       datacnt)
        Return intRtn

    End Function
#End Region
End Class
