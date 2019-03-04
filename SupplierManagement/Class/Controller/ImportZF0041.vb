Imports System.Text

Public Class ImportZF0041
    Inherits BaseAdapter   
    Private myForm As FormImportZF0041
    Private FileName As String
    Private sb As New StringBuilder
    Private AgvalueSB As New StringBuilder
    Private AgvalueUpdateSB As New StringBuilder
    Private AGDTSB As New StringBuilder
    Private AGDTUpdateSB As New StringBuilder
    Private _ErrorMessage As String = String.Empty
    Private myselectedDate As Date
    Public ReadOnly Property ErrorMessage As String
        Get
            Return _ErrorMessage
        End Get
    End Property

    Public Sub New(ByVal myForm As FormImportZF0041, ByVal FileName As String)
        Me.myForm = myForm
        Me.FileName = FileName
        myselectedDate = myForm.DateTimePicker1.Value
    End Sub

    Public Function doImport() As Boolean
        Dim myret As Boolean = False

        'Fill dataset agvalue
        'agreementdt
        'agreementtx
        sb.Append("select * from agvalue;")
        sb.Append("select * from agreementdt;")
        DS = New DataSet
        Dim mymessage As String = String.Empty
        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try
                Dim pk(0) As DataColumn
                pk(0) = DS.Tables(0).Columns("agreement")
                DS.Tables(0).PrimaryKey = pk
                DS.Tables(0).TableName = "AGValue"

                Dim pk1(1) As DataColumn
                pk1(0) = DS.Tables(1).Columns("agreement")             
                pk1(1) = DS.Tables(1).Columns("material")
                DS.Tables(1).PrimaryKey = pk1
                DS.Tables(1).TableName = "AgreementDT"


            Catch ex As Exception
                myForm.ProgressReport(1, "Loading Data. Error::" & ex.Message)
                myForm.ProgressReport(2, "Continuous")

            End Try

      
        End If


        'Find agvalue, if not avail insert else update
        'find agreementdt if not avail insert else update
        'find agreementtx if not avail insert else update
        'Use Copy and Update from List
        AGDTSB.Clear()
        AGDTUpdateSB.Clear()
        AgvalueSB.Clear()
        AgvalueUpdateSB.Clear()

        Dim myrecord() As String
        Using objTFParser = New FileIO.TextFieldParser(FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                myForm.ProgressReport(1, "Read Data")
                myForm.ProgressReport(2, "Continuous")
                Dim PrevAgreement As String = String.Empty
                Dim result As DataRow
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 1 Then

                        If PrevAgreement <> myrecord(1) Then
                            If myrecord(1) = 9330035766 Then
                                Debug.Print("debug")
                            End If
                            PrevAgreement = myrecord(1)
                            Dim pkey(0) As Object
                            pkey(0) = myrecord(1)
                            result = DS.Tables(0).Rows.Find(pkey)
                            'agreement bigint NOT NULL,  value numeric, closingdate date,  item integer,
                            'cmmf integer,  shorttext character varying,vendorcode bigint,  reference character varying,  trackingno character varying,
                            'amount numeric, crcy character(3), agqty integer, desiredqty numeric, oun character(2), totalqty integer, deliveredqty integer,
                            'status character varying, startdate date, enddate date, headertext character varying, agreementdate date, created character varying,
                            ' tgtqty integer, periodyear character varying,
                            If Not IsNothing(result) Then
                                'update
                                If AgvalueUpdateSB.Length > 0 Then
                                    AgvalueUpdateSB.Append(",")
                                End If
                                AgvalueUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying,'{6}'::character varying,'{7}'::character varying,'{8}'::character varying,'{9}'::character varying,'{10}'::character varying,'{11}'::character varying,'{12}'::character varying,'{13}'::character varying,'{14}'::character varying,'{15}'::character varying,'{16}'::character varying,'{17}'::character varying,'{18}'::character varying,'{19}'::character varying,'{20}'::character varying,'{21}'::character varying,'{22}'::character varying,'{23}'::character varying]",
                                                                  myrecord(1), myrecord(16), validSAPDate(myrecord(25)), myrecord(2), myrecord(4), myrecord(5), myrecord(6), myrecord(8), myrecord(9), validNum(myrecord(10)), myrecord(11), validNum(myrecord(12)), validNum(myrecord(13)), myrecord(14), validNum(myrecord(21)), validNum(myrecord(23)), myrecord(24), validSAPDate(myrecord(17)), validSAPDate(myrecord(26)), myrecord(27), validSAPDate(myrecord(28)), myrecord(29), validNum(myrecord(22)), myrecord(15)))


                            Else
                                'Create
                                AgvalueSB.Append(myrecord(1) & vbTab &
                                                 myrecord(16) & vbTab &
                                                 validSAPDate(myrecord(25)) & vbTab &
                                                 myrecord(2) & vbTab &
                                                 myrecord(4) & vbTab &
                                                 validStr(myrecord(5)) & vbTab &
                                                 myrecord(6) & vbTab &
                                                 validStr(myrecord(8)) & vbTab &
                                                 validStr(myrecord(9)) & vbTab &
                                                 validNum(myrecord(10)) & vbTab &
                                                 myrecord(11) & vbTab &
                                                 validNum(myrecord(12)) & vbTab &
                                                 validNum(myrecord(13)) & vbTab &
                                                 validStr(myrecord(14)) & vbTab &
                                                 validNum(myrecord(21)) & vbTab &
                                                 validNum(myrecord(23)) & vbTab &
                                                 validStr(myrecord(24)) & vbTab &
                                                 validSAPDate(myrecord(17)) & vbTab &
                                                 validSAPDate(myrecord(26)) & vbTab &
                                                 myrecord(27) & vbTab &
                                                 validSAPDate(myrecord(28)) & vbTab &
                                                 validStr(myrecord(29)) & vbTab &
                                                 validNum(myrecord(22)) & vbTab &
                                                 myrecord(15) & vbCrLf)
                            End If
                        End If


                        Dim pkey1(1) As Object
                        pkey1(0) = myrecord(1)
                        pkey1(1) = myrecord(3)

                        result = DS.Tables(1).Rows.Find(pkey1)
                        If Not IsNothing(result) Then
                            'update
                            If AGDTUpdateSB.Length > 0 Then
                                AGDTUpdateSB.Append(",")
                            End If
                            AGDTUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying]",
                                                              result.Item("agdtid"), myrecord(3), myrecord(6), validSAPDate(myrecord(30)), validSAPDate(myrecord(31))))

                        Else
                            AGDTSB.Append(myrecord(1) & vbTab &
                                                myrecord(3) & vbTab &
                                                myrecord(6) & vbTab &
                                                validSAPDate(myrecord(30)) & vbTab &
                                                validSAPDate(myrecord(31)) & vbCrLf)
                        End If
                    End If



                    count += 1
                Loop
            End With
        End Using
        'update record
        Dim sqlstr As String = String.Empty
        If AgvalueSB.Length > 0 Then
            sqlstr = "copy agvalue(agreement,value,closingdate,item,cmmf,shorttext,vendorcode,reference,trackingno," &
                     "amount,crcy,agqty,desiredqty,oun,totalqty,deliveredqty,status,startdate,enddate,headertext," &
                     "agreementdate,created,tgtqty,periodyear) from stdin with null as 'Null';"
            _ErrorMessage = DbAdapter1.copy(sqlstr, AgvalueSB.ToString, myret)
            If Not myret Then
                Return myret
            End If
        End If
        If AgvalueUpdateSB.Length > 0 Then
            'agreement bigint NOT NULL,  value numeric, closingdate date,  item integer,
            'cmmf integer,  shorttext character varying,vendorcode bigint,  reference character varying,  trackingno character varying,
            'amount numeric, crcy character(3), agqty integer, desiredqty numeric, oun character(2), totalqty integer, deliveredqty integer,
            'status character varying, startdate date, enddate date, headertext character varying, agreementdate date, created character varying,
            ' tgtqty integer, periodyear character varying,
            myForm.ProgressReport(1, "Update AGValue")
            sqlstr = "update agvalue set value= foo.value::numeric,closingdate = foo.closingdate::date,item=foo.item::integer,cmmf=foo.cmmf::integer,shorttext=foo.shorttext::character varying,vendorcode = foo.vendorcode::bigint," &
                     " reference=foo.reference,trackingno = foo.trackingno,amount = foo.amount::numeric,crcy = foo.crcy::character varying,agqty=foo.agqty::integer,desiredqty=foo.desiredqty::numeric," &
                     " oun=foo.oun::character varying,totalqty=foo.totalqty::integer,deliveredqty=foo.deliveredqty::integer,status=foo.status,startdate=foo.startdate::date,enddate=foo.enddate::date,headertext = foo.headertext," &
                     " agreementdate=foo.agreementdate::date,created=foo.created,tgtqty=foo.tgtqty::integer,periodyear=foo.periodyear" &
            " from (select * from array_to_set24(Array[" & AgvalueUpdateSB.ToString &
            "]) as tb (id character varying,value character varying,closingdate character varying,  item character varying,cmmf character varying,  shorttext character varying,vendorcode character varying,  reference character varying,  trackingno character varying,amount character varying, crcy character varying, agqty character varying, desiredqty character varying, oun character varying, totalqty character varying, deliveredqty character varying," &
            " status character varying, startdate character varying, enddate character varying, headertext character varying, agreementdate character varying, created character varying, tgtqty character varying, periodyear character varying))foo " &
            " where agreement = foo.id::bigint;"

            myret = DbAdapter1.ExecuteNonQuery(sqlstr.Replace("'Null'::character varying", "Null").Replace("''::character varying", "Null"), message:=_ErrorMessage)
        End If

        If AGDTSB.Length > 0 Then
            sqlstr = "copy agreementdt(agreement,material,vendorcode,validfrom,validto) from stdin with null as 'Null';"
            _ErrorMessage = DbAdapter1.copy(sqlstr, AGDTSB.ToString, myret)
            If Not myret Then
                Return myret
            End If
        End If
        If AGDTUpdateSB.Length > 0 Then
            myForm.ProgressReport(1, "Update Agreement Detail")
            sqlstr = "update agreementdt set material = foo.material::bigint,vendorcode=foo.vendorcode::bigint,validfrom = foo.validfrom::date,validto=foo.validto::date" &
            " from (select * from array_to_set5(Array[" & AGDTUpdateSB.ToString &
            "]) as tb (id character varying,material character varying,vendorcode character varying,  validfrom character varying,validto character varying))foo " &
            " where agdtid = foo.id::bigint;"

            myret = DbAdapter1.ExecuteNonQuery(sqlstr.Replace("'Null'::character varying", "Null").Replace("''::character varying", "Null"), message:=_ErrorMessage)
        End If

        'Update agv1
        sqlstr = String.Format("update cmmfvendorprice cvp set agv1  = foo.value from (select at.material,at.postingdate,ag.value,ag.vendorcode from agreementtx at " &
                 " left join agvalue ag on ag.agreement = at.agreement )foo " &
                 " where myyear = {0:yyyy} and cvp.cmmf = foo.material and cvp.initialtx = foo.postingdate and cvp.sts1 and cvp.vendorcode = foo.vendorcode", myselectedDate)
        myret = (DbAdapter1.ExecuteNonQuery(sqlstr, message:=_ErrorMessage))
        'Update agv2
        sqlstr = String.Format("update cmmfvendorprice cvp set agv2  = foo.value from (select at.material,at.postingdate,ag.value,ag.vendorcode from agreementtx at " &
                 " left join agvalue ag on ag.agreement = at.agreement )foo " &
                 " where myyear = {0:yyyy} and cvp.cmmf = foo.material and cvp.lasttx = foo.postingdate and cvp.sts2 and cvp.vendorcode = foo.vendorcode", myselectedDate)
        myret = (DbAdapter1.ExecuteNonQuery(sqlstr, message:=_ErrorMessage))

        Return myret
    End Function

    Private Function validSAPDate(ByVal mydate As String) As String

        If mydate = "" Then
            Return "Null"
        Else
            '"yyyy-MM-dd"
            Dim myStr = mydate.Split(".")
            Return String.Format("{0}-{1}-{2}", myStr(2), myStr(1), myStr(0))
        End If
    End Function

    Private Function validNum(ByVal myNum As String) As String
        If myNum = "" Then
            Return "Null"
        Else
            Return myNum.Replace(",", "")
        End If

    End Function

    Private Function validStr(ByVal myStr As String) As String
        If myStr = "" Then
            Return "Null"
        Else
            Return String.Format("{0}", myStr)
        End If
    End Function

End Class
