'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : ceMD_MODIFYTRANS_TEMP.vb ( MD_MODIFYTRANS_TEMP Entity 처리 Class)
'기      능 : MD_MODIFYTRANS_TEMP Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-06-25 오전 11:17:52 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class ceMD_MODIFYTRANS_TEMP
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ceMD_MODIFYTRANS_TEMP"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strstrSQL = strSQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(Optional ByVal strTRANSYEARMON As String = OPTIONAL_STR, _
            Optional ByVal strTRANSNO As String = OPTIONAL_STR, _
            Optional ByVal strTRANSNOSEQ As String = OPTIONAL_STR, _
            Optional ByVal strMEDFLAG As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder

        Try
            BuildNameValues(",", "TRANSYEARMON", strTRANSYEARMON, strFields, strValues)
            BuildNameValues(",", "TRANSNO", strTRANSNO, strFields, strValues)
            BuildNameValues(",", "TRANSNOSEQ", strTRANSNOSEQ, strFields, strValues)
            BuildNameValues(",", "MEDFLAG", strMEDFLAG, strFields, strValues)

            strSQL = String.Format("INSERT INTO {0} ({1}) VALUES({2})", EntityName, strFields, strValues)

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".InsertDo")
        End Try
    End Function


    '*****************************************************************
    '입력 : strstrSQL = strSQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Delete 처리
    '참고 : Key 조건이 선택적임(OPTIONAL_STR/OPTIONAL_NUM/OPTIONAL_DTM)
    '*****************************************************************
    Public Function DeleteDo(ByVal strTRANSYEARMON As String, _
                             ByVal strTRANSNO As String) As Integer
        Dim strSQL As String

        Try
            strSQL = "DELETE FROM MD_MODIFYTRANS_TEMP WHERE TRANSYEARMON ='" & strTRANSYEARMON & "' AND TRANSNO = '" & strTRANSNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = strSQL 문 By KTH
    '반환 : 처리건수
    '기능 : 거래명세 디테일 의 세금계산서 번호 업데이트
    '*****************************************************************
    'strTRANSYEARMON, lngTRANSNO, lngSEQ, strTAXYEARMON, intTAXNO
    Public Function INSERTINTO_TEMP(ByVal strTAXYEARMON As String, _
                                    ByVal strTAXSNO As String, _
                                    ByVal strMEDFLAG As String, _
                                    ByVal strTAB_NAME As String) As Integer
        'strTRANSNO, strTAXNO
        Dim strSQL As String
        Try
            strSQL = " INSERT INTO MD_MODIFYTRANS_TEMP (TRANSYEARMON, TRANSNO, TRANSNOSEQ,MEDFLAG)"
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & " TRANSYEARMON, "
            strSQL = strSQL & " TRANSNO, "
            strSQL = strSQL & " SEQ, "
            strSQL = strSQL & " '" & strMEDFLAG & "'"
            strSQL = strSQL & " FROM " & strTAB_NAME
            strSQL = strSQL & " WHERE 1=1"
            strSQL = strSQL & " AND TAXYEARMON = '" & strTAXYEARMON & "'"
            strSQL = strSQL & " AND TAXNO  = '" & strTAXSNO & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".INSERTINTO_TEMP")
        End Try
    End Function


    '*****************************************************************
    '입력 : strSQL = strSQL 문 By KTH
    '반환 : 처리건수
    '기능 : 거래명세 디테일 의 세금계산서 번호 업데이트
    '*****************************************************************
    'strTRANSYEARMON, lngTRANSNO, lngSEQ, strTAXYEARMON, intTAXNO 
    Public Function UPDATE_TRANS_HDR(ByVal strTRANSYEARMON As String, _
                                     ByVal strTRANSNO As String, _
                                     ByVal strAMT As Double, _
                                     ByVal strVAT As Double, _
                                     ByVal strMEDFLAGHDR As String) As Integer

        Dim strSQL As String
        Try
            If strMEDFLAGHDR = "O" Then
                strSQL = " UPDATE MD_INTERNETTRANS_HDR"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMT & ","
                strSQL = strSQL & " VAT = " & strVAT
                strSQL = strSQL & " WHERE 1=1"
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMON & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNO & "'"
            ElseIf strMEDFLAGHDR = "B" Then
                strSQL = " UPDATE MD_PRINTTRANS_HDR"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMT & ","
                strSQL = strSQL & " VAT = " & strVAT
                strSQL = strSQL & " WHERE 1=1"
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMON & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNO & "'" '"
            ElseIf strMEDFLAGHDR = "A" Then
                strSQL = " UPDATE MD_ELEC_TRANS_HDR"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMT & ","
                strSQL = strSQL & " VAT = " & strVAT
                strSQL = strSQL & " WHERE 1=1"
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMON & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNO & "'"
            ElseIf strMEDFLAGHDR = "A2" Then
                strSQL = " UPDATE MD_CATVTRANS_HDR"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMT & ","
                strSQL = strSQL & " VAT = " & strVAT
                strSQL = strSQL & " WHERE 1=1"
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMON & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNO & "'"
            End If

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".INSERTINTO_TEMP")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = strSQL 문 By KTH
    '반환 : 처리건수
    '기능 : 거래명세 디테일 의 세금계산서 번호 업데이트
    '*****************************************************************
    'strTRANSYEARMON, lngTRANSNO, lngSEQ, strTAXYEARMON, intTAXNO
    Public Function UPDATE_TRANS_DTL(ByVal strTRANSYEARMONDTL As String, _
                                     ByVal strTRANSNODTL As String, _
                                     ByVal strSEQDTL As String, _
                                     ByVal strMED_FLAGDTL As String, _
                                     ByVal strAMTDTL As Double, _
                                     ByVal strVATDTL As Double, _
                                     ByVal strMEDCODEDTL As String, _
                                     ByVal strREAL_MED_CODEDTL As String) As Integer
        Dim strSQL As String
        Try
            If strMED_FLAGDTL = "O" Then
                strSQL = " UPDATE MD_INTERNETTRANS_DTL"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMTDTL & ","
                strSQL = strSQL & " VAT = " & strVATDTL & ","
                strSQL = strSQL & " SUMAMTVAT = " & strAMTDTL + strVATDTL & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODEDTL & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODEDTL & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMONDTL & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNODTL & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQDTL & "'"

            ElseIf strMED_FLAGDTL = "B" Then
                strSQL = " UPDATE MD_PRINTTRANS_DTL"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMTDTL & ","
                strSQL = strSQL & " VAT = " & strVATDTL & ","
                strSQL = strSQL & " SUMAMTVAT = " & strAMTDTL + strVATDTL & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODEDTL & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODEDTL & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMONDTL & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNODTL & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQDTL & "'"

            ElseIf strMED_FLAGDTL = "A" Then
                strSQL = " UPDATE MD_ELEC_TRANS_DTL"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMTDTL & ","
                strSQL = strSQL & " VAT = " & strVATDTL & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODEDTL & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODEDTL & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMONDTL & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNODTL & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQDTL & "'"
            ElseIf strMED_FLAGDTL = "A2" Then
                strSQL = " UPDATE MD_CATVTRANS_DTL"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMTDTL & ","
                strSQL = strSQL & " VAT = " & strVATDTL & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODEDTL & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODEDTL & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND TRANSYEARMON = '" & strTRANSYEARMONDTL & "'"
                strSQL = strSQL & " AND TRANSNO = '" & strTRANSNODTL & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQDTL & "'"
            End If
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".INSERTINTO_TEMP")
        End Try
    End Function

    '*****************************************************************
    '입력 : strSQL = strSQL 문 By KTH
    '반환 : 처리건수
    '기능 : 거래명세 디테일 의 세금계산서 번호 업데이트
    '*****************************************************************
    'strTRANSYEARMON, lngTRANSNO, lngSEQ, strTAXYEARMON, intTAXNO
    Public Function UPDATE_MEDIUM(ByVal strYEARMON As String, _
                                  ByVal strSEQ As String, _
                                  ByVal strMED_FLAG As String, _
                                  ByVal strAMTMEDIUM As Double, _
                                  ByVal strCOMMI_RATE As String, _
                                  ByVal strCOMMISSION As Double, _
                                  ByVal strMEDCODE As String, _
                                  ByVal strREAL_MED_CODE As String) As Integer

        Dim strSQL As String
        Try
            If strMED_FLAG = "O" Then
                strSQL = " UPDATE MD_INTERNET_MEDIUM"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMTMEDIUM & ","
                strSQL = strSQL & " COMMI_RATE = '" & strCOMMI_RATE & "',"
                strSQL = strSQL & " COMMISSION = " & strCOMMISSION & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODE & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODE & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQ & "'"

            ElseIf strMED_FLAG = "B" Then
                strSQL = " UPDATE MD_BOOKING_MEDIUM"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMOUNT = " & strAMTMEDIUM & ","
                strSQL = strSQL & " COMMI_RATE = '" & strCOMMI_RATE & "',"
                strSQL = strSQL & " COMMISSION = " & strCOMMISSION & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODE & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODE & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQ & "'"

            ElseIf strMED_FLAG = "A" Then
                strSQL = " UPDATE MD_ELECTRIC_MEDIUM"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMTMEDIUM & ","
                strSQL = strSQL & " COMMI_RATE = '" & strCOMMI_RATE & "',"
                strSQL = strSQL & " COMMISSION = " & strCOMMISSION & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODE & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODE & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQ & "'"

            ElseIf strMED_FLAG = "A2" Then
                strSQL = " UPDATE MD_CATV_MEDIUM"
                strSQL = strSQL & " SET"
                strSQL = strSQL & " AMT = " & strAMTMEDIUM & ","
                strSQL = strSQL & " COMMI_RATE = '" & strCOMMI_RATE & "',"
                strSQL = strSQL & " COMMISSION = " & strCOMMISSION & ","
                strSQL = strSQL & " MEDCODE = '" & strMEDCODE & "',"
                strSQL = strSQL & " REAL_MED_CODE = '" & strREAL_MED_CODE & "'"
                strSQL = strSQL & " WHERE 1=1 "
                strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                strSQL = strSQL & " AND SEQ = '" & strSEQ & "'"
            End If
            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".INSERTINTO_TEMP")
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
        MyBase.EntityName = "MD_MODIFYTRANS_TEMP"     'Entity Name 설정
    End Sub

    '해제 기능은 Base Class에서 구현되어 있음
#End Region
#End Region
End Class

'------->>엔티티 INSERT/UPDATE 샘플입니다. 반드시 자신의 환경에 맞추어서 변경하시기 바랍니다.
'=========================================================
'       'vntData Array를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceMD_MODIFYTRANS_TEMP.InsertDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNOSEQ", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDFLAG", intColCnt, intRow) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceMD_MODIFYTRANS_TEMP.UpdateDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNOSEQ", intColCnt, intRow), _
'                                       GetElement(vntData,"MEDFLAG", intColCnt, intRow) _
'                                       )
'        Return intRtn


'=========================================================
'       'XmlData 를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjceMD_MODIFYTRANS_TEMP.InsertDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO"), _
'                                       XMLGetElement(xmlRoot,"TRANSNOSEQ"), _
'                                       XMLGetElement(xmlRoot,"MEDFLAG") _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjceMD_MODIFYTRANS_TEMP.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO"), _
'                                       XMLGetElement(xmlRoot,"TRANSNOSEQ"), _
'                                       XMLGetElement(xmlRoot,"MEDFLAG") _
'                                       )
'        Return intRtn


