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

Public Class ccMDITINTERNETCOMMIVOCH
    Inherits ccControl
#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccMDITINTERNETCOMMIVOCH"                  '자신의 클래스명
    Private mobjceMD_COMMIVOCH_MST As eMDCO.ceMD_COMMIVOCH_MST             '사용할 Entity 변수 선언
    Private mobjceMD_VOCHFILE_MST As eMDCO.ceMD_VOCHFILE_MST
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password" '커넥션Setting
#End Region
    '년월,광고주,매체사,생성유무
#Region "GROUP BLOCK : 외부에비공개"
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strREAL_MED_CODE As String, _
                              ByVal strVOCHFLAG As String, _
                              ByVal strFILENO As String) As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2, Con3, Con4, Con5 As String

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            '한글인 경우
            If strYEARMON <> "" Then Con1 = String.Format(" AND (A.TAXYEARMON = '{0}')", strYEARMON)
            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND A.CLIENTCODE = '{0}'", strCLIENTCODE)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND A.REAL_MED_CODE = '{0}'", strREAL_MED_CODE)
            If strVOCHFLAG <> "A" Then
                If strVOCHFLAG = "Y" Then
                    Con4 = String.Format(" AND (CASE ISNULL(A.VOCHNO,'N') WHEN 'N' THEN 'N' WHEN '' THEN 'N' ELSE 'Y' END   = '{0}')", strVOCHFLAG)
                ElseIf strVOCHFLAG = "N" Then
                    Con4 = String.Format(" AND ISNULL(B.RMSNO,'') = '' AND (CASE ISNULL(A.VOCHNO,'N') WHEN 'N' THEN 'N' WHEN '' THEN 'N' ELSE 'Y' END   = '{0}')", strVOCHFLAG)
                ElseIf strVOCHFLAG = "M" Then
                    Con4 = String.Format(" AND ISNULL(B.RMSNO,'') <> '' AND (CASE ISNULL(A.VOCHNO,'N') WHEN 'N' THEN 'N' WHEN '' THEN 'N' ELSE 'Y' END   = '{0}')", "N")
                End If
            End If

            If strFILENO <> "" Then Con5 = String.Format(" AND B.RMSNO = '{0}'", strFILENO)
            '조회 필드 설정

            strSelFields = " A.DEMANDDAY POSTINGDATE,"
            strSelFields = strSelFields & " replace(A.real_med_bisno,'-','') CUSTOMERCODE, A.REAL_MED_NAME, "
            'strSelFields = strSelFields & " case isnull(b.summ,'') when '' then convert(char(12),RTRIM(LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.REAL_MED_CODE))))+' 대행수수료' else b.summ end as  SUMM,"
            strSelFields = strSelFields & " case isnull(b.summ,'') when '' then convert(char(12),RTRIM(LTRIM(DBO.md_get_taxmedname_fun(A.taxyearmon, a.taxno))))+' 대행수수료' else b.summ end as  SUMM,"
            strSelFields = strSelFields & " '3000' BA,"
            strSelFields = strSelFields & " '53105' COSTCENTER,"
            strSelFields = strSelFields & " A.SUMAMT,"
            strSelFields = strSelFields & " A.VAT,"
            strSelFields = strSelFields & " 'B5' SEMU,"
            strSelFields = strSelFields & " '8000' BP,"
            'strSelFields = strSelFields & " A.DEMANDDAY,"
            strSelFields = strSelFields & " convert(char(8) , DATEADD(mm, 3,A.DEMANDDAY),112) DEMANDDAY, "
            strSelFields = strSelFields & " A.TAXYEARMON,"
            strSelFields = strSelFields & " A.TAXNO,"
            strSelFields = strSelFields & " 'S' GBN,"
            strSelFields = strSelFields & " A.VOCHNO,B.RMSNO,A.MEDFLAG,B.ERRCODE,B.ERRMSG"
            strFormat = "SELECT {0} FROM MD_COMMITAX_HDR  A LEFT JOIN MD_COMMIVOCH_MST B ON A.TAXYEARMON = B.TAXYEARMON AND A.TAXNO = B.TAXNO" & _
                                     " WHERE A.MEDFLAG IN ('O')  {1} {2} {3} {4} {5} "
            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2, Con3, Con4, Con5)

            '데이터 조회
            Try
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
    Public Function GetFILENO(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strFILENO As String) As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2 As String
        Con1 = ""
        Con2 = ""

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            '한글인 경우
            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
            If strFILENO <> "" Then Con2 = String.Format(" AND RMSNO = '{0}'", strFILENO)

            '조회 필드 설정

            strSelFields = " RMSNO,CASE ENDFLAG WHEN 'N' THEN '처리중' ELSE '처리완료' END ENDFLAG,DBO.SC_EMPNAME_FUN(CUSER) CUSER,CDATE,YEARMON"

            strFormat = "SELECT {0} FROM MD_VOCHFILE_MST WHERE 1=1 {1} {2}  "
            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetFILENO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object, _
                               ByVal strYEARMON As String, _
                               ByVal strSAVEYEARMON As String, _
                               ByVal strSAVESEQ As Double, _
                               ByVal strSAVERMSNO As String) As Integer
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblID As Double '자동 ID 를사용할 때만 사용
        Dim strSC_EMP_STATUS As String
        Dim vntData2 As Object
        Dim intSEQ As Double
        Dim strRMSNO As String
        Dim strPOSTINGDATE As String
        Dim strDEMANDDAY As String

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    'File 정보저장

                    'POSTINGDATE,DEMANDDAY


                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceMD_COMMIVOCH_MST = New ceMD_COMMIVOCH_MST(mobjSCGLConfig)
                    mobjceMD_VOCHFILE_MST = New ceMD_VOCHFILE_MST(mobjSCGLConfig)

                    mobjceMD_VOCHFILE_MST.FileInsertDo(strSAVEYEARMON, strSAVESEQ, strSAVERMSNO, "N")
                    '''vntData의 컬럼수, 로우수를 변수입력

                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''해당하는Row 만큼 Loop
                    strSC_EMP_STATUS = ""
                    For i = 1 To intRows
                        If Trim(GetElement(vntData, "CHK", intColCnt, i)) = "" Then
                        Else
                            If GetElement(vntData, "CHK", intColCnt, i) = 1 Then
                                If GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strPOSTINGDATE = GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "POSTINGDATE", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)
                                If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(0, 4) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(5, 2) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).SUBSTRING(8, 2)
                                intRtn = UpdateRtn(vntData, intColCnt, i, strSAVERMSNO, strPOSTINGDATE, strDEMANDDAY)
                            End If
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
                mobjceMD_COMMIVOCH_MST.Dispose()
                mobjceMD_VOCHFILE_MST.Dispose()

            End Try
        End With
    End Function
    Public Function SelectRtn_SEQNO(ByVal strYEARMON As String) As Object
        '여기부터 단순조회
        Dim strSQL, strFormat, strRtn As String
        Dim intRowCnt As Double
        Dim intColCnt As Double
        Dim vntData As Object
        'SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체


            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strSQL = "select '" & strYEARMON & "' yearmon,isnull(max(seq),0)+1 seq,'" & strYEARMON & "'+dbo.lpad(isnull(max(seq),0)+1,4,'0')+'_S' RMSNO from md_vochfile_mst where yearmon = '" & strYEARMON & "'"
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
        '여기까지 단순조회
    End Function
    Public Function VOCHDELL(ByVal strInfoXML As String, _
                             ByVal strYEAR As String, _
                             ByVal strVOCHNO As String, _
                             ByVal strTAXYEARMON As String, _
                             ByVal strTAXNO As Double) As Integer
        Dim intRtnDell

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()


                'File 정보저장

                'POSTINGDATE,DEMANDDAY


                '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                mobjceMD_COMMIVOCH_MST = New ceMD_COMMIVOCH_MST(mobjSCGLConfig)

                '전표내역 삭제
                mobjceMD_COMMIVOCH_MST.Delete(strYEAR, strVOCHNO)
                '세금계산서의 전표번호 '' 로 업데이트
                mobjceMD_COMMIVOCH_MST.UpdateDelete(strTAXYEARMON, strTAXNO)

                mobjceMD_COMMIVOCH_MST.Update_vochno(strTAXYEARMON, strTAXNO, "INTERNET")



                .mobjSCGLSql.SQLCommitTrans()
                Return intRtnDell
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".VOCHDELL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_COMMIVOCH_MST.Dispose()

            End Try
        End With
    End Function
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                                   ByVal vntData As Object) As Integer
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblID As Double '자동 ID 를사용할 때만 사용
        Dim strSC_EMP_STATUS As String
        Dim vntData2 As Object
        Dim intSEQ As Double

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    'File 정보저장

                    'POSTINGDATE,DEMANDDAY


                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceMD_COMMIVOCH_MST = New ceMD_COMMIVOCH_MST(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력

                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''해당하는Row 만큼 Loop
                    For i = 1 To intRows
                        If Trim(GetElement(vntData, "CHK", intColCnt, i)) = "" Then
                        Else
                            If GetElement(vntData, "CHK", intColCnt, i) = 1 And GetElement(vntData, "ERRCODE", intColCnt, i) = 1 Then
                                intRtn = DeleteRtn(vntData, intColCnt, i)
                            End If
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".DeleteRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_COMMIVOCH_MST.Dispose()


            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function UpdateRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strRMSNO As String, _
                               ByVal strPOSTINGDATE As String, _
                               ByVal strDEMANDDAY As String) As Integer
        'strPOSTINGDATE,strDEMANDDAY
        Dim intRtn As Integer
        'POSTINGDATE,CUSTOMERCODE,SUMM,BA,SUMAMT,VAT,SEMU,BP,DEMANDDAY,VENDOR,TAXYEARMON,TAXNO,GBN,VOCHNO,RMSNO
        intRtn = mobjceMD_COMMIVOCH_MST.InsertDo( _
                                       strPOSTINGDATE, _
                                       GetElement(vntData, "CUSTOMERCODE", intColCnt, intRow), _
                                       GetElement(vntData, "SUMM", intColCnt, intRow), _
                                       GetElement(vntData, "BA", intColCnt, intRow), _
                                       GetElement(vntData, "COSTCENTER", intColCnt, intRow), _
                                       GetElement(vntData, "SUMAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "VAT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "SEMU", intColCnt, intRow), _
                                       GetElement(vntData, "BP", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "TAXYEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "TAXNO", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "GBN", intColCnt, intRow), _
                                       GetElement(vntData, "VOCHNO", intColCnt, intRow), _
                                       strRMSNO, _
                                       GetElement(vntData, "MEDFLAG", intColCnt, intRow), _
                                       "523204", _
                                       strPOSTINGDATE)
        'GetElement(vntData, "COMMI_RATE", intColCnt, intRow, NULL_NUM, True), _
        Return intRtn
    End Function
    Private Function DeleteRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer) As Integer
        'strPOSTINGDATE,strDEMANDDAY
        Dim intRtn As Integer
        'POSTINGDATE,CUSTOMERCODE,SUMM,BA,SUMAMT,VAT,SEMU,BP,DEMANDDAY,VENDOR,TAXYEARMON,TAXNO,GBN,VOCHNO,RMSNO
        intRtn = mobjceMD_COMMIVOCH_MST.DeleteDo( _
                                       GetElement(vntData, "TAXYEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "TAXNO", intColCnt, intRow, NULL_NUM, True))
        'GetElement(vntData, "COMMI_RATE", intColCnt, intRow, NULL_NUM, True), _
        Return intRtn
    End Function
#End Region
End Class
