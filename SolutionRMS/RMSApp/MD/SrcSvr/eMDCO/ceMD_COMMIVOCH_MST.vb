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

Public Class ceMD_COMMIVOCH_MST
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_COMMIVOCH_MST"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"

    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Update 처리
    '참고 : Key 조건과 Value Field가선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    'POSTINGDATE,CUSTOMERCODE,SUMM,BA,SUMAMT,VAT,SEMU,BP,DEMANDDAY,VENDOR,TAXYEARMON,TAXNO,GBN,VOCHNO,RMSNO

    Public Function InsertDo(Optional ByVal strPOSTINGDATE As String = OPTIONAL_STR, _
                             Optional ByVal strCUSTOMERCODE As String = OPTIONAL_STR, _
                             Optional ByVal strSUMM As String = OPTIONAL_STR, _
                             Optional ByVal strBA As String = OPTIONAL_STR, _
                             Optional ByVal strCOSTCENTER As String = OPTIONAL_STR, _
                             Optional ByVal strSUMAMT As Double = OPTIONAL_NUM, _
                             Optional ByVal strVAT As Double = OPTIONAL_NUM, _
                             Optional ByVal strSEMU As String = OPTIONAL_STR, _
                             Optional ByVal strBP As String = OPTIONAL_STR, _
                             Optional ByVal strDEMANDDAY As String = OPTIONAL_STR, _
                             Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                             Optional ByVal strTAXNO As Double = OPTIONAL_NUM, _
                             Optional ByVal strGBN As String = OPTIONAL_STR, _
                             Optional ByVal strVOCHNO As String = OPTIONAL_STR, _
                             Optional ByVal strRMSNO As String = OPTIONAL_STR, _
                             Optional ByVal strMEDFLAG As String = OPTIONAL_STR, _
                             Optional ByVal strATTR01 As String = OPTIONAL_STR, _
                             Optional ByVal strATTR02 As String = OPTIONAL_STR) As Integer


        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            BuildNameValues(",", "POSTINGDATE", strPOSTINGDATE, strFields, strValues)
            BuildNameValues(",", "CUSTOMERCODE", strCUSTOMERCODE, strFields, strValues)
            BuildNameValues(",", "SUMM", strSUMM, strFields, strValues)
            BuildNameValues(",", "BA", strBA, strFields, strValues)
            BuildNameValues(",", "COSTCENTER", strCOSTCENTER, strFields, strValues)
            BuildNameValues(",", "SUMAMT", strSUMAMT, strFields, strValues)
            BuildNameValues(",", "VAT", strVAT, strFields, strValues)
            BuildNameValues(",", "SEMU", strSEMU, strFields, strValues)
            BuildNameValues(",", "BP", strBP, strFields, strValues)
            BuildNameValues(",", "DEMANDDAY", strDEMANDDAY, strFields, strValues)
            BuildNameValues(",", "TAXYEARMON", strTAXYEARMON, strFields, strValues)
            BuildNameValues(",", "TAXNO", strTAXNO, strFields, strValues)
            BuildNameValues(",", "GBN", strGBN, strFields, strValues)
            BuildNameValues(",", "VOCHNO", strVOCHNO, strFields, strValues)
            BuildNameValues(",", "RMSNO", strRMSNO, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)
            BuildNameValues(",", "ATTR01", strATTR01, strFields, strValues)
            BuildNameValues(",", "ATTR02", strATTR02, strFields, strValues)
            BuildNameValues(",", "CUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "CDATE", strNOW, strFields, strValues)
            BuildNameValues(",", "UUSER", mobjSCGLConfig.WRKUSR, strFields, strValues)
            BuildNameValues(",", "UDATE", strNOW, strFields, strValues)
            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function
    Public Function UpdateDo(Optional ByVal strCC_CODE As String = OPTIONAL_STR, _
            Optional ByVal strOC_CODE As String = OPTIONAL_STR, _
            Optional ByVal strUSEYN As String = OPTIONAL_STR, _
            Optional ByVal strSDATE As String = OPTIONAL_STR, _
            Optional ByVal strEDATE As String = OPTIONAL_STR) As Integer
        Dim strSQL As String
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now
        Try
            strSQL = String.Format("UPDATE {0} SET {1} WHERE {2}", EntityName, _
                     BuildFields(",", _
                        GetFieldNameValue("OC_CODE", strOC_CODE), _
                        GetFieldNameValue("USE_YN", strUSEYN), _
                        GetFieldNameValue("SDATE", strSDATE), _
                        GetFieldNameValue("EDATE", strEDATE)), _
                     BuildFields("AND", _
                        GetFieldNameValue("CC_CODE", strCC_CODE)))
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDo")
        End Try
    End Function
    Public Function Delete(ByVal strYEAR As String, _
                                 ByVal strVOCHNO As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "DELETE FROM  MD_COMMIVOCH_MST WHERE SUBSTRING(TAXYEARMON,1,4)='" & strYEAR & "' AND VOCHNO ='" & strVOCHNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDelete")
        End Try
    End Function
    Public Function UpdateDelete(ByVal strTAXYEARMON As String, _
                                 ByVal strTAXNO As Double) As Integer
        Dim strSQL As String
        Try
            strSQL = "UPDATE MD_COMMITAX_HDR SET VOCHNO ='' WHERE TAXYEARMON = '" & strTAXYEARMON & "' AND TAXNO = '" & strTAXNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDelete")
        End Try
    End Function

    Public Function Update_vochno(ByVal strTAXYEARMON As String, _
                                  ByVal strTAXNO As Double, _
                                  ByVal strFLAG As String) As Integer
        Dim strSQL As String
        Try
            If strFLAG = "CATV" Then
                strSQL = "UPDATE MD_CATV_MEDIUM SET COMMI_VOCH_NO ='' WHERE COMMI_TAX_NO = '" & strTAXYEARMON & "-" & strTAXNO & "'"
            ElseIf strFLAG = "PRINT" Then
                strSQL = "UPDATE MD_BOOKING_MEDIUM SET COMMI_VOCH_NO ='' WHERE COMMI_TAX_NO = '" & strTAXYEARMON & "-" & strTAXNO & "'"
            ElseIf strFLAG = "INTERNET" Then
                strSQL = "UPDATE MD_INTERNET_MEDIUM SET COMMI_VOCH_NO ='' WHERE COMMI_TAX_NO = '" & strTAXYEARMON & "-" & strTAXNO & "'"
            ElseIf strFLAG = "ELEC" Then
                strSQL = "UPDATE MD_ELECTRIC_MEDIUM SET COMMI_VOCH_NO ='' WHERE COMMI_TAX_NO = '" & strTAXYEARMON & "-" & strTAXNO & "'"
            ElseIf strFLAG = "OUTDOOR" Then
                strSQL = "UPDATE MD_OUTDOOR_MEDIUM SET COMMI_VOCH_NO ='' WHERE COMMI_TAX_NO = '" & strTAXYEARMON & "-" & strTAXNO & "'"
            End If

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".UpdateDelete")
        End Try
    End Function

    Public Function DeleteDo(Optional ByVal strTAXYEARMON As String = OPTIONAL_STR, _
                             Optional ByVal strTAXNO As Double = OPTIONAL_NUM) As Integer
        Dim strSQL As String
        Try
            strSQL = "DELETE FROM  MD_COMMIVOCH_MST WHERE TAXYEARMON='" & strTAXYEARMON & "' AND TAXNO =" & strTAXNO

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".Delete")
        End Try
    End Function
    'Public Function DeleteDo(ByVal USERID As String) As Integer
    '    Dim strSQL As String

    '    Try
    '        strSQL = "DELETE FROM MD_TAX_TEMP WHERE USERID ='" & USERID & "'"

    '        Return ProcEntity(strSQL)
    '    Catch err As Exception
    '        Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
    '    End Try
    'End Function
#End Region

#Region "객체 생성/해제"
    '*****************************************************************
    '입력 : strInfoXML = 공통기본정보에 대한 XML
    'objSCGLSql = DB 처리 객체 인스턴싱 변수    '반환 : 없음
    '기능 : DB 처리를 위한 공통기본정보 설정
    '*****************************************************************
    Public Sub New(Optional ByVal objSCGLConfig As SCGLUtil.cbSCGLConfig = Nothing, Optional ByVal strInfoXML As String = "")
        MyBase.SetConfig(objSCGLConfig, strInfoXML)
        MyBase.EntityName = "MD_COMMIVOCH_MST"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region

End Class







