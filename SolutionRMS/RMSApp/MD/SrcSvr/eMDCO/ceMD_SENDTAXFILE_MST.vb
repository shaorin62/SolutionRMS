'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceSC_REALMEDCODE_MST.vb ( SC_REALMEDCODE_MST Entity 처리 Class)
'기      능 : SC_REALMEDCODE_MST Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-01-14 오전 11:09:29 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_SENDTAXFILE_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_SENDTAXFILE_MST"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Update 처리
    '참고 : Key 조건과 Value Field가선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function FileInsertDo(Optional ByVal strYEARMON As String = OPTIONAL_STR, _
                                 Optional ByVal intSEQ As Double = OPTIONAL_NUM, _
                                 Optional ByVal strRMSNO As String = OPTIONAL_STR, _
                                 Optional ByVal strENDFLAG As String = OPTIONAL_STR, _
                                 Optional ByVal strSEND_GBN As String = OPTIONAL_STR) As Integer


        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "YEARMON", strYEARMON, strFields, strValues)
            BuildNameValues(",", "SEQ", intSEQ, strFields, strValues)
            BuildNameValues(",", "RMSNO", strRMSNO, strFields, strValues)
            BuildNameValues(",", "ENDFLAG", strENDFLAG, strFields, strValues)
            BuildNameValues(",", "SEND_GBN", strSEND_GBN, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function UpdateRtn_RMSNO_MD_TRUTAX_HDR(ByVal strTAXYEARMON, _
                                                  ByVal intTAXNO, _
                                                  ByVal strMEDFLAG, _
                                                  ByVal strRMSNO, _
                                                  ByVal strCANCEL_YN) As Integer
        Dim strSQL As String

        Try
            strSQL = " UPDATE MD_TRUTAX_HDR "
            strSQL = strSQL & "   SET RMSNO = '" & strRMSNO & "', ATTR04 = '" & strCANCEL_YN & "'"
            strSQL = strSQL & "   WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = '" & intTAXNO & "' AND MEDFLAG = '" & strMEDFLAG & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UPDATE_TRANS_DO")
        End Try
    End Function

#End Region

#Region "객체 생성/해제"
    '*****************************************************************
    '입력 : strInfoXML = 공통기본정보에 대한 XML
    'objSCGLSql = DB 처리 객체 인스턴싱 변수    '반환 : 없음
    '기능 : DB 처리를 위한 공통기본정보 설정
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "MD_SENDTAXFILE_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region

End Class






