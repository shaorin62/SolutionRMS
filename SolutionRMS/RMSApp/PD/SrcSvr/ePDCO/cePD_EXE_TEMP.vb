'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 엔티티 클래스 메이커 - 한화 S&C
'시스템구분 : 솔루션명/시스템명/Server Entity Class
'실행  환경 : GAC(Global Assembly Cache)
'프로그램명 : cePD_TRANS_TEMP.vb ( PD_TRANS_TEMP Entity 처리 Class)
'기      능 : PD_TRANS_TEMP Entity에 대해Insert/Update/Delete/Select를 처리
'             - 부모엔티티 객체인 SCGLUtil.ceEntity를 상속
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008-03-11 오전 10:44:27 By Making Entity Bean
'****************************************************************************************
Imports SCGLUtil.cbSCGLUtil     '공통 유틸리티 객체
Imports SCGLUtil.cbSCGLErr      '공통 오류처리 객체
Imports SCGLEntity              '엔티티 객체의 부모 객체

Public Class cePD_EXE_TEMP
    Inherits ceEntity

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "cePD_EXE_TEMP"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"
#Region "SQL Insert/Update/Delete/Select"
    '*****************************************************************
    '입력 : strSQL = SQL 문
    '반환 : 처리건수
    '기능 : 해당 Entity에 Insert 처리
    '*****************************************************************
    Public Function InsertDo(Optional ByVal strJOBNO As String = OPTIONAL_STR, _
                             Optional ByVal dblNUM As Double = OPTIONAL_NUM, _
                             Optional ByVal strUSERID As String = OPTIONAL_STR)

        Dim strSQL As String
        Dim strFields As New System.Text.StringBuilder
        Dim strValues As New System.Text.StringBuilder
        Dim strNOW As String '데이트형의 처리는 변수를 받아 텍스트로 처리 한다.. 
        strNOW = Now

        Try
            BuildNameValues(",", "JOBNO", strJOBNO, strFields, strValues)
            BuildNameValues(",", "NUM", dblNUM, strFields, strValues)
            BuildNameValues(",", "USERID", strUSERID, strFields, strValues)

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
    Public Function DeleteDo(ByVal USERID As String) As Integer
        Dim strSQL As String
        Try
            strSQL = "DELETE FROM PD_EXE_TEMP WHERE USERID ='" & USERID & "'"

            Return ProcEntity(strSQL)
        Catch err As Exception
            Throw RaiseSysErr(err, CLASS_NAME & ".DeleteDo")
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
        MyBase.EntityName = "PD_EXE_TEMP"     'Entity Name 설정 
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
'        intRtn = mobjcePD_TRANS_TEMP.InsertDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow), _
'                                       GetElement(vntData,"DATACNT", intColCnt, intRow), _
'                                       GetElement(vntData,"USERID", intColCnt, intRow) _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjcePD_TRANS_TEMP.UpdateDo( _
'                                       GetElement(vntData,"TRANSYEARMON", intColCnt, intRow), _
'                                       GetElement(vntData,"TRANSNO", intColCnt, intRow), _
'                                       GetElement(vntData,"DATACNT", intColCnt, intRow), _
'                                       GetElement(vntData,"USERID", intColCnt, intRow) _
'                                       )
'        Return intRtn


'=========================================================
'       'XmlData 를 사용할 때 Insert/Update 입니다.
'=========================================================
'        Dim intRtn As Integer
'        intRtn = mobjcePD_TRANS_TEMP.InsertDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO"), _
'                                       XMLGetElement(xmlRoot,"DATACNT"), _
'                                       XMLGetElement(xmlRoot,"USERID") _
'                                       )
'        Return intRtn

'        Dim intRtn As Integer
'        intRtn = mobjcePD_TRANS_TEMP.UpdateDo( _
'                                       XMLGetElement(xmlRoot,"TRANSYEARMON"), _
'                                       XMLGetElement(xmlRoot,"TRANSNO"), _
'                                       XMLGetElement(xmlRoot,"DATACNT"), _
'                                       XMLGetElement(xmlRoot,"USERID") _
'                                       )
'        Return intRtn


