
Imports System.Data
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Windows.Forms

Public Class clsTallyDet

    Public Class clsCompanyDet

        Dim port_no_val As String
        Dim dsn_val As String

        Public Sub New()
            port_no_val = "9000"
        End Sub

        Public Property Port() As String

            Get
                Port = port_no_val
            End Get

            Set(ByVal Value As String)
                port_no_val = Value
                dsn_val = "TallyODBC_" & port_no_val
            End Set

        End Property

        Public ReadOnly Property DSN() As String

            Get
                DSN = dsn_val
            End Get

        End Property

    End Class

End Class

<Serializable()>
Public Class clsGroupMaster_Tally

    Implements ICloneable

    Public XmlDet As clsXmlDet

    <Serializable()>
    Class clsXmlDet

        Implements ICloneable

        Dim _tagCurCompany As clsXmlTag

        Public _tagGrpName As clsXmlTag
        Public _tagParent As clsXmlTag
        Public _tagName As clsXmlTag
        Public _tagAlias As clsXmlTag

        Public _lstTagDet As New List(Of clsXmlTag)

        Public XmlParseTag As Object

        Dim _xmlImport As String
        Dim _xmlExport As String

        Public Sub New()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            XmlParseTag = {"GROUP NAME", "GROUP"}

            _tagGrpName = New clsXmlTag
            _tagGrpName = tagdefnGrpName.Clone
            _lstTagDet.Add(_tagGrpName)

            _tagParent = New clsXmlTag
            _tagParent = tagdefnParent.Clone
            _lstTagDet.Add(_tagParent)

            _tagName = New clsXmlTag
            _tagName = tagdefnName.Clone
            _lstTagDet.Add(_tagName)

            _tagAlias = New clsXmlTag
            _tagAlias = tagdefnAlias.Clone
            _lstTagDet.Add(_tagAlias)

            InitOthDet()
            InitCreateDet()
            InitFindDet()

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitOthDet()

            _tagCurCompany = New clsXmlTag
            With _tagCurCompany
                .Name = "tag_cur_company"
                .Id = "svcurrentcompany"
                .DataType = EnumDataType.Text
            End With

        End Sub

        Private Sub InitCreateDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         " <IMPORTDATA>" & vbCrLf &
                         "  <REQUESTDESC>" & vbCrLf &
                         "   <REPORTNAME>All Masters</REPORTNAME>" & vbCrLf &
                         "   <STATICVARIABLES>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   </STATICVARIABLES>" & vbCrLf &
                         "  </REQUESTDESC>" & vbCrLf &
                         "  <REQUESTDATA>" & vbCrLf

            '            _xmlImport = _xmlImport &
            '                         "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImport = _xmlImport &
                         "   <TALLYMESSAGE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   <GROUP NAME=""" & _tagGrpName.Name & """ RESERVEDNAME="""" Action = ""Create"">" & vbCrLf &
                         "    <PARENT>" & _tagParent.Name & "</PARENT>" & vbCrLf &
                         "     <LANGUAGENAME.LIST>" & vbCrLf &
                         "      <NAME.LIST TYPE=""String"">" & vbCrLf &
                         "       <NAME>" & _tagName.Name & "</NAME>" & vbCrLf &
                         "       <NAME>" & _tagAlias.Name & "</NAME>" & vbCrLf &
                         "      </NAME.LIST>" & vbCrLf &
                         "     </LANGUAGENAME.LIST>" & vbCrLf &
                         "    </GROUP>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   </TALLYMESSAGE>" & vbCrLf &
                         "  </REQUESTDATA>" & vbCrLf &
                         " </IMPORTDATA>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlImport = UCase(_xmlImport)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitFindDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlExport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <VERSION>1</VERSION>" & vbCrLf &
                         " <TALLYREQUEST>Export</TALLYREQUEST>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <TYPE>OBJECT</TYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <SUBTYPE>Group</SUBTYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <ID TYPE=""Name"">" & _tagGrpName.Name & "</ID>" & vbCrLf

            _xmlExport = _xmlExport &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         "  <DESC>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <STATICVARIABLES>" & vbCrLf &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf &
                         "    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>" & vbCrLf &
                         "   </STATICVARIABLES>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <FETCHLIST>" & vbCrLf &
                         "    <FETCH>" & _tagName.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagParent.Id & "</FETCH>" & vbCrLf &
                         "   </FETCHLIST>" & vbCrLf

            _xmlExport = _xmlExport &
                         "  </DESC>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlExport = UCase(_xmlExport)

            Cursor.Current = mpointer_lcl

        End Sub

        Public Function CreateDet(ByVal CurCompany_par As String,
                                  ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim create_xml_det As String = _xmlImport

            _tagCurCompany.Data = CurCompany_par

            create_xml_det = PrepXmlFromText(create_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            CreateDet = PostXml(PrepXmlFromText(create_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function FindDet(ByVal CurCompany_par As String,
                                ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim find_xml_det As String = _xmlExport

            _tagCurCompany.Data = CurCompany_par

            find_xml_det = PrepXmlFromText(find_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            FindDet = PostXml(PrepXmlFromText(find_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Sub New()

        XmlDet = New clsXmlDet

    End Sub

    Public Function CloneOld() As clsGroupMaster_Tally

        Dim temp = DirectCast(Me.MemberwiseClone, clsGroupMaster_Tally)

        With temp
            '            .XmlDet = DirectCast(Me.XmlDet.Clone, clsXmlDet)
        End With

        Return temp

    End Function

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim mem_stream As New MemoryStream()
        Dim bin_formatter As New BinaryFormatter()

        bin_formatter.Serialize(mem_stream, Me)
        mem_stream.Seek(0, SeekOrigin.Begin)

        Return bin_formatter.Deserialize(mem_stream)

    End Function

End Class

<Serializable()>
Public Class clsLedgerMaster_Tally

    Implements ICloneable

    Public XmlDet As clsXmlDet

    <Serializable()>
    Class clsXmlDet

        Implements ICloneable

        Dim _tagCurCompany As clsXmlTag

        Public _tagLedgName As clsXmlTag
        Public _tagName As clsXmlTag
        Public _tagAlias As clsXmlTag
        Public _tagMailName As clsXmlTag
        Public _tagParent As clsXmlTag
        Public _tagAddr1 As clsXmlTag
        Public _tagAddr2 As clsXmlTag
        Public _tagAddr3 As clsXmlTag
        Public _tagStateName As clsXmlTag
        Public _tagPinCode As clsXmlTag
        Public _tagLedgPhone As clsXmlTag
        Public _tagIncomeTaxNo As clsXmlTag
        Public _tagAffectsStock As clsXmlTag
        Public _tagUseForVat As clsXmlTag
        Public _tagOpeningBaln As clsXmlTag
        Public _tagIsCostCentresOn As clsXmlTag

        Public _lstTagDet As New List(Of clsXmlTag)

        Public XmlParseTag As Object

        Dim _xmlImport As String
        Dim _xmlExport As String

        Public Sub New()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            XmlParseTag = {"LEDGER NAME", "LEDGER"}

            _tagLedgName = New clsXmlTag
            _tagLedgName = tagdefnLedgName.Clone
            _lstTagDet.Add(_tagLedgName)

            _tagName = New clsXmlTag
            _tagName = tagdefnName.Clone
            _lstTagDet.Add(_tagName)

            _tagAlias = New clsXmlTag
            _tagAlias = tagdefnAlias.Clone
            _lstTagDet.Add(_tagAlias)

            _tagMailName = New clsXmlTag
            _tagMailName = tagdefnMailName.Clone
            _lstTagDet.Add(_tagMailName)

            _tagParent = New clsXmlTag
            _tagParent = tagdefnParent.Clone
            _lstTagDet.Add(_tagParent)

            _tagAddr1 = New clsXmlTag
            _tagAddr1 = tagdefnAddr1.Clone
            _lstTagDet.Add(_tagAddr1)

            _tagAddr2 = New clsXmlTag
            _tagAddr2 = tagdefnAddr2.Clone
            _lstTagDet.Add(_tagAddr2)

            _tagAddr3 = New clsXmlTag
            _tagAddr3 = tagdefnAddr3.Clone
            _lstTagDet.Add(_tagAddr3)

            _tagStateName = New clsXmlTag
            _tagStateName = tagdefnStateName.Clone
            _lstTagDet.Add(_tagStateName)

            _tagPinCode = New clsXmlTag
            _tagPinCode = tagdefnPinCode.Clone
            _lstTagDet.Add(_tagPinCode)

            _tagLedgPhone = New clsXmlTag
            _tagLedgPhone = tagdefnLedgPhone.Clone
            _lstTagDet.Add(_tagLedgPhone)

            _tagIncomeTaxNo = New clsXmlTag
            _tagIncomeTaxNo = tagdefnIncomeTaxNo.Clone
            _lstTagDet.Add(_tagIncomeTaxNo)

            _tagAffectsStock = New clsXmlTag
            _tagAffectsStock = tagdefnAffectsStock.Clone
            _lstTagDet.Add(_tagAffectsStock)

            _tagUseForVat = New clsXmlTag
            _tagUseForVat = tagdefnUseForVat.Clone
            _lstTagDet.Add(_tagUseForVat)

            _tagOpeningBaln = New clsXmlTag
            _tagOpeningBaln = tagdefnOpeningBaln.Clone
            _lstTagDet.Add(_tagOpeningBaln)

            _tagIsCostCentresOn = New clsXmlTag
            _tagIsCostCentresOn = tagdefnIsCostCentresOn.Clone
            _lstTagDet.Add(_tagIsCostCentresOn)

            InitOthDet()
            InitCreateDet()
            InitFindDet()

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitOthDet()

            _tagCurCompany = New clsXmlTag
            With _tagCurCompany
                .Name = "tag_cur_company"
                .Id = "svcurrentcompany"
                .DataType = EnumDataType.Text
            End With

        End Sub

        Private Sub InitCreateDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         " <IMPORTDATA>" & vbCrLf &
                         "  <REQUESTDESC>" & vbCrLf &
                         "   <REPORTNAME>All Masters</REPORTNAME>" & vbCrLf &
                         "   <STATICVARIABLES>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   </STATICVARIABLES>" & vbCrLf &
                         "  </REQUESTDESC>" & vbCrLf &
                         "  <REQUESTDATA>" & vbCrLf

            '            _xmlImport = _xmlImport &
            '                         "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImport = _xmlImport &
                         "   <TALLYMESSAGE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <LEDGER NAME=""" & _tagLedgName.Name & """ RESERVEDNAME="""" Action = ""Create"">" & vbCrLf &
                         "     <ADDRESS.LIST TYPE=""String"">" & vbCrLf &
                         "      <ADDRESS>" & _tagAddr1.Name & "</ADDRESS>" & vbCrLf &
                         "      <ADDRESS>" & _tagAddr2.Name & "</ADDRESS>" & vbCrLf &
                         "      <ADDRESS>" & _tagAddr3.Name & "</ADDRESS>" & vbCrLf &
                         "     </ADDRESS.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <MAILINGNAME.LIST TYPE=""String"">" & vbCrLf &
                         "      <MAILINGNAME>" & _tagMailName.Name & "</MAILINGNAME>" & vbCrLf &
                         "     </MAILINGNAME.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <ALTEREDON/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <CURRENCYNAME/>" & vbCrLf &
                         "     <PINCODE/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <INCOMETAXNUMBER/>" & vbCrLf &
                         "     <INTERSTATESTNUMBER/>" & vbCrLf &
                         "     <VATTINNUMBER/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <PARENT>" & _tagParent.Name & "</PARENT>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <LEDGERPHONE>" & _tagLedgPhone.Name & "</LEDGERPHONE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <ALTEREDBY/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <LEDGERPHONE/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <TAXCLASSIFICATIONNAME/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <TAXTYPE/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <GSTTYPE/>" & vbCrLf &
                         "     <APPROPRIATEFOR/>" & vbCrLf &
                         "     <SERVICECATEGORY/>" & vbCrLf &
                         "     <EXCISELEDGERCLASSIFICATION/>" & vbCrLf &
                         "     <EXCISEDUTYTYPE/>" & vbCrLf &
                         "     <EXCISEALLOCTYPE/>" & vbCrLf &
                         "     <EXCISENATUREOFPURCHASE/>" & vbCrLf &
                         "     <LEDGERFBTCATEGORY/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <ISBILLWISEON/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <ISCOSTCENTRESON/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <AFFECTSSTOCK/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <USEFORVAT/>" & vbCrLf

            _xmlImport = _xmlImport &
                          "     <OPENINGBALANCE/>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <LANGUAGENAME.LIST>" & vbCrLf &
                         "      <NAME.LIST TYPE=""String"">" & vbCrLf &
                         "       <NAME>" & _tagName.Name & "</NAME>" & vbCrLf &
                         "       <NAME>" & _tagAlias.Name & "</NAME>" & vbCrLf &
                         "      </NAME.LIST>" & vbCrLf &
                         "     </LANGUAGENAME.LIST>" & vbCrLf &
                         "    </LEDGER>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   </TALLYMESSAGE>" & vbCrLf &
                         "  </REQUESTDATA>" & vbCrLf &
                         " </IMPORTDATA>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlImport = UCase(_xmlImport)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitFindDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlExport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <VERSION>1</VERSION>" & vbCrLf &
                         " <TALLYREQUEST>Export</TALLYREQUEST>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <TYPE>OBJECT</TYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <SUBTYPE>Ledger</SUBTYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <ID TYPE=""Name"">" & _tagLedgName.Name & "</ID>" & vbCrLf

            _xmlExport = _xmlExport &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         "  <DESC>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <STATICVARIABLES>" & vbCrLf &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf &
                         "    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>" & vbCrLf &
                         "   </STATICVARIABLES>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <FETCHLIST>" & vbCrLf &
                         "    <FETCH>" & _tagName.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagParent.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagAddr1.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagLedgPhone.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagOpeningBaln.Id & "</FETCH>" & vbCrLf &
                         "   </FETCHLIST>" & vbCrLf

            _xmlExport = _xmlExport &
                         "  </DESC>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlExport = UCase(_xmlExport)

            Cursor.Current = mpointer_lcl

        End Sub

        Public Function CreateDet(ByVal CurCompany_par As String,
                                  ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim create_xml_det As String = _xmlImport

            _tagCurCompany.Data = CurCompany_par

            create_xml_det = PrepXmlFromText(create_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            CreateDet = PostXml(PrepXmlFromText(create_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function FindDet(ByVal CurCompany_par As String,
                                ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim find_xml_det As String = _xmlExport

            _tagCurCompany.Data = CurCompany_par

            find_xml_det = PrepXmlFromText(find_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            FindDet = PostXml(PrepXmlFromText(find_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Sub New()

        XmlDet = New clsXmlDet

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim mem_stream As New MemoryStream()
        Dim bin_formatter As New BinaryFormatter()

        bin_formatter.Serialize(mem_stream, Me)
        mem_stream.Seek(0, SeekOrigin.Begin)

        Return bin_formatter.Deserialize(mem_stream)

    End Function

End Class

<Serializable()>
Public Class clsJournal_Tally

    Implements ICloneable

    Public XmlDet As clsXmlDet

    <Serializable()>
    Class clsXmlDet

        Implements ICloneable

        Dim _tagCurCompany As clsXmlTag

        Public _tagVoucherTypeOrigName As clsXmlTag
        Public _tagVoucherNo As clsXmlTag
        Public _tagRefNo As clsXmlTag
        Public _tagDate As clsXmlTag
        Public _tagEffectiveDate As clsXmlTag
        Public _tagPartyLedger As clsXmlTag
        Public _tagByParty As clsXmlTag
        Public _tagToParty As clsXmlTag
        Public _tagAmount As clsXmlTag
        Public _tagNarration As clsXmlTag
        Public _tagMasterId As clsXmlTag

        Public _lstTagDet As New List(Of clsXmlTag)

        Public XmlParseTag As Object

        Dim _xmlImport As String
        Dim _xmlExport As String

        Public Sub New()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            XmlParseTag = {"", ""}

            _tagVoucherTypeOrigName = New clsXmlTag
            _tagVoucherTypeOrigName = tagdefnVoucherTypeOrigName.Clone
            _lstTagDet.Add(_tagVoucherTypeOrigName)

            _tagVoucherNo = New clsXmlTag
            _tagVoucherNo = tagdefnVoucherNo.Clone
            _lstTagDet.Add(_tagVoucherNo)

            _tagRefNo = New clsXmlTag
            _tagRefNo = tagdefnRefNo.Clone
            With _tagRefNo
                .DataPfx_1 = "JOUR"
                .DataPfx_2 = ""
            End With
            _lstTagDet.Add(_tagRefNo)

            _tagDate = New clsXmlTag
            _tagDate = tagdefnDate.Clone
            _lstTagDet.Add(_tagDate)

            _tagEffectiveDate = New clsXmlTag
            _tagEffectiveDate = tagdefnEffectiveDate.Clone
            _lstTagDet.Add(_tagEffectiveDate)

            _tagPartyLedger = New clsXmlTag
            _tagPartyLedger = tagdefnPartyLedger.Clone
            _lstTagDet.Add(_tagPartyLedger)

            _tagByParty = New clsXmlTag
            _tagByParty = tagdefnByParty.Clone
            _lstTagDet.Add(_tagByParty)

            _tagToParty = New clsXmlTag
            _tagToParty = tagdefnToParty.Clone
            _lstTagDet.Add(_tagToParty)

            _tagAmount = New clsXmlTag
            _tagAmount = tagdefnAmount.Clone
            _lstTagDet.Add(_tagAmount)

            _tagNarration = New clsXmlTag
            _tagNarration = tagdefnNarration.Clone
            _lstTagDet.Add(_tagNarration)

            _tagMasterId = New clsXmlTag
            _tagMasterId = tagdefnMasterId.Clone
            _lstTagDet.Add(_tagMasterId)

            InitOthDet()
            InitCreateDet()
            InitFindDet()

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitOthDet()

            _tagCurCompany = New clsXmlTag
            With _tagCurCompany
                .Name = "tag_cur_company"
                .Id = "svcurrentcompany"
                .DataType = EnumDataType.Text
            End With

        End Sub

        Private Sub InitCreateDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         " <IMPORTDATA>" & vbCrLf &
                         "  <REQUESTDESC>" & vbCrLf &
                         "   <REPORTNAME>Vouchers</REPORTNAME>" & vbCrLf &
                         "   <STATICVARIABLES>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   </STATICVARIABLES>" & vbCrLf &
                         "  </REQUESTDESC>" & vbCrLf &
                         "  <REQUESTDATA>" & vbCrLf

            '            _xmlImport = _xmlImport & _
            '                         "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImport = _xmlImport &
                         "   <TALLYMESSAGE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <VOUCHER VCHTYPE=""Journal"" ACTION=""Create"" OBJView=""Accounting Voucher View"">" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <DATE>" & _tagDate.Name & "</DATE>" & vbCrLf &
                         "     <REFERENCE>" & _tagRefNo.Name & "</REFERENCE>" & vbCrLf &
                         "     <NARRATION>" & _tagNarration.Name & "</NARRATION>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <VOUCHERTYPENAME>Journal</VOUCHERTYPENAME>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <PARTYLEDGERNAME>" & _tagPartyLedger.Name & "</PARTYLEDGERNAME>" & vbCrLf &
                         "     <EFFECTIVEDATE>" & _tagEffectiveDate.Name & "</EFFECTIVEDATE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                         "       <LEDGERNAME>" & _tagByParty.Name & "</LEDGERNAME>" & vbCrLf &
                         "       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                         "       <AMOUNT>-" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                         "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                         "      <LEDGERNAME>" & _tagToParty.Name & "</LEDGERNAME>" & vbCrLf &
                         "      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                         "      <AMOUNT>" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                         "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    </VOUCHER>" & vbCrLf &
                         "   </TALLYMESSAGE>" & vbCrLf &
                         "  </REQUESTDATA>" & vbCrLf &
                         " </IMPORTDATA>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlImport = UCase(_xmlImport)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitFindDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlExport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <VERSION>1</VERSION>" & vbCrLf &
                         " <TALLYREQUEST>Export</TALLYREQUEST>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <TYPE>OBJECT</TYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <SUBTYPE>Voucher</SUBTYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <ID TYPE=""Name"">" & _tagMasterId.Name & "</ID>" & vbCrLf

            _xmlExport = _xmlExport &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         "  <DESC>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <STATICVARIABLES>" & vbCrLf &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf &
                         "    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>" & vbCrLf &
                         "   </STATICVARIABLES>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <FETCHLIST>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherNo.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagDate.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherTypeOrigName.Id & "</FETCH>" & vbCrLf &
                         "   </FETCHLIST>" & vbCrLf

            _xmlExport = _xmlExport &
                         "  </DESC>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlExport = UCase(_xmlExport)

            Cursor.Current = mpointer_lcl

        End Sub

        Public Function CreateDet(ByVal CurCompany_par As String,
                                  ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim create_xml_det As String = _xmlImport

            _tagCurCompany.Data = CurCompany_par

            create_xml_det = PrepXmlFromText(create_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            CreateDet = PostXml(PrepXmlFromText(create_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function FindDet(ByVal CurCompany_par As String,
                                ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim find_xml_det As String = _xmlExport

            _tagCurCompany.Data = CurCompany_par

            find_xml_det = PrepXmlFromText(find_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            FindDet = PostXml(PrepXmlFromText(find_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Sub New()

        XmlDet = New clsXmlDet

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim mem_stream As New MemoryStream()
        Dim bin_formatter As New BinaryFormatter()

        bin_formatter.Serialize(mem_stream, Me)
        mem_stream.Seek(0, SeekOrigin.Begin)

        Return bin_formatter.Deserialize(mem_stream)

    End Function

End Class

<Serializable()>
Public Class clsReceipt_Tally

    Implements ICloneable

    Public XmlDet As clsXmlDet

    <Serializable()>
    Class clsXmlDet

        Implements ICloneable

        Dim _tagCurCompany As clsXmlTag

        Public IsSingleEntryMode As Boolean

        Public _tagVoucherTypeOrigName As clsXmlTag
        Public _tagVoucherNo As clsXmlTag
        Public _tagRefNo As clsXmlTag
        Public _tagDate As clsXmlTag
        Public _tagEffectiveDate As clsXmlTag
        Public _tagAccount As clsXmlTag
        Public _tagAmount As clsXmlTag
        Public _tagPartyLedger As clsXmlTag
        Public _tagParticulars As clsXmlTag
        Public _tagNarration As clsXmlTag
        Public _tagMasterId As clsXmlTag

        Public _lstTagDet As New List(Of clsXmlTag)

        Public XmlParseTag As Object

        ' Particulars

        Public _lstParticulars As New List(Of List(Of clsXmlTag))

        Dim name_parti_index As Integer = 0
        Dim amt_parti_index As Integer = 1

        Dim _xmlImportSingleMode As String
        Dim _xmlImportJournalMode As String
        Dim _xmlImportParticulars As String

        Dim _xmlImport As String
        Dim _xmlExport As String

        Public Sub New()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            XmlParseTag = {"", ""}

            _tagVoucherTypeOrigName = New clsXmlTag
            _tagVoucherTypeOrigName = tagdefnVoucherTypeOrigName.Clone
            _lstTagDet.Add(_tagVoucherTypeOrigName)

            _tagVoucherNo = New clsXmlTag
            _tagVoucherNo = tagdefnVoucherNo.Clone
            _lstTagDet.Add(_tagVoucherNo)

            _tagRefNo = New clsXmlTag
            _tagRefNo = tagdefnRefNo.Clone
            With _tagRefNo
                .DataPfx_1 = "RECI"
                .DataPfx_2 = ""
            End With
            _lstTagDet.Add(_tagRefNo)

            _tagDate = New clsXmlTag
            _tagDate = tagdefnDate.Clone
            _lstTagDet.Add(_tagDate)

            _tagEffectiveDate = New clsXmlTag
            _tagEffectiveDate = tagdefnEffectiveDate.Clone
            _lstTagDet.Add(_tagEffectiveDate)

            _tagAccount = New clsXmlTag
            _tagAccount = tagdefnAccount.Clone
            _lstTagDet.Add(_tagAccount)

            _tagAmount = New clsXmlTag
            _tagAmount = tagdefnAmount.Clone
            _lstTagDet.Add(_tagAmount)

            _tagPartyLedger = New clsXmlTag
            _tagPartyLedger = tagdefnPartyLedger.Clone
            _lstTagDet.Add(_tagPartyLedger)

            _tagParticulars = New clsXmlTag
            _tagParticulars = tagdefnParticulars.Clone
            _lstTagDet.Add(_tagParticulars)

            _tagNarration = New clsXmlTag
            _tagNarration = tagdefnNarration.Clone
            _lstTagDet.Add(_tagNarration)

            _tagMasterId = New clsXmlTag
            _tagMasterId = tagdefnMasterId.Clone
            _lstTagDet.Add(_tagMasterId)

            InitOthDet()
            InitCreateDet()
            InitFindDet()

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitOthDet()

            _tagCurCompany = New clsXmlTag
            With _tagCurCompany
                .Name = "tag_cur_company"
                .Id = "svcurrentcompany"
                .DataType = EnumDataType.Text
            End With

            IsSingleEntryMode = False

        End Sub

        Private Sub InitCreateSingleModeDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImportParticulars = _xmlImportParticulars &
                                     "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                     "      <LEDGERNAME>" & _tagParticulars.Name & "</LEDGERNAME>" & vbCrLf &
                                     "      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                     "      <AMOUNT>" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                     "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportParticulars = UCase(_xmlImportParticulars)


            _xmlImportSingleMode = "<ENVELOPE>" & vbCrLf &
                                   "<HEADER>" & vbCrLf &
                                   " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                                   "</HEADER>" & vbCrLf &
                                   "<BODY>" & vbCrLf &
                                   " <IMPORTDATA>" & vbCrLf &
                                   "  <REQUESTDESC>" & vbCrLf &
                                   "   <REPORTNAME>Vouchers</REPORTNAME>" & vbCrLf &
                                   "   <STATICVARIABLES>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "   </STATICVARIABLES>" & vbCrLf &
                                   "  </REQUESTDESC>" & vbCrLf &
                                   "  <REQUESTDATA>" & vbCrLf

            '            _xmlImportSingleMode = _xmlImportSingleMode & _
            '                                   "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "   <TALLYMESSAGE>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "    <VOUCHER VCHTYPE=""Receipt"" ACTION=""Create"" OBJView=""Accounting Voucher View"">" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <DATE>" & _tagDate.Name & "</DATE>" & vbCrLf &
                                   "     <REFERENCE>" & _tagRefNo.Name & "</REFERENCE>" & vbCrLf &
                                   "     <NARRATION>" & _tagNarration.Name & "</NARRATION>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <VOUCHERTYPENAME>Receipt</VOUCHERTYPENAME>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <PARTYLEDGERNAME>" & _tagPartyLedger.Name & "</PARTYLEDGERNAME>" & vbCrLf &
                                   "     <EFFECTIVEDATE>" & _tagEffectiveDate.Name & "</EFFECTIVEDATE>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "_xmlImportParticulars"

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                   "       <LEDGERNAME>" & _tagAccount.Name & "</LEDGERNAME>" & vbCrLf &
                                   "       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                                   "       <AMOUNT>-" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                   "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "    </VOUCHER>" & vbCrLf &
                                   "   </TALLYMESSAGE>" & vbCrLf &
                                   "  </REQUESTDATA>" & vbCrLf &
                                   " </IMPORTDATA>" & vbCrLf &
                                   "</BODY>" & vbCrLf &
                                   "</ENVELOPE>"

            _xmlImportSingleMode = UCase(_xmlImportSingleMode)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitCreateJournalModeDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImportJournalMode = "<ENVELOPE>" & vbCrLf &
                                    "<HEADER>" & vbCrLf &
                                    " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                                    "</HEADER>" & vbCrLf &
                                    "<BODY>" & vbCrLf &
                                    " <IMPORTDATA>" & vbCrLf &
                                    "  <REQUESTDESC>" & vbCrLf &
                                    "   <REPORTNAME>Vouchers</REPORTNAME>" & vbCrLf &
                                    "   <STATICVARIABLES>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "   </STATICVARIABLES>" & vbCrLf &
                                    "  </REQUESTDESC>" & vbCrLf &
                                    "  <REQUESTDATA>" & vbCrLf

            '            _xmlImportJournalMode = _xmlImportJournalMode & _
            '                                    "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "   <TALLYMESSAGE>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "    <VOUCHER VCHTYPE=""Receipt"" ACTION=""Create"" OBJView=""Accounting Voucher View"">" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <DATE>" & _tagDate.Name & "</DATE>" & vbCrLf &
                                    "     <REFERENCE>" & _tagRefNo.Name & "</REFERENCE>" & vbCrLf &
                                    "     <NARRATION>" & _tagNarration.Name & "</NARRATION>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <VOUCHERTYPENAME>Receipt</VOUCHERTYPENAME>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <PARTYLEDGERNAME>" & _tagPartyLedger.Name & "</PARTYLEDGERNAME>" & vbCrLf &
                                    "     <EFFECTIVEDATE>" & _tagEffectiveDate.Name & "</EFFECTIVEDATE>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                    "      <LEDGERNAME>" & _tagParticulars.Name & "</LEDGERNAME>" & vbCrLf &
                                    "      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                    "      <AMOUNT>" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                    "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                    "       <LEDGERNAME>" & _tagAccount.Name & "</LEDGERNAME>" & vbCrLf &
                                    "       <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                                    "       <AMOUNT>-" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                    "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "    </VOUCHER>" & vbCrLf &
                                    "   </TALLYMESSAGE>" & vbCrLf &
                                    "  </REQUESTDATA>" & vbCrLf &
                                    " </IMPORTDATA>" & vbCrLf &
                                    "</BODY>" & vbCrLf &
                                    "</ENVELOPE>"

            _xmlImportJournalMode = UCase(_xmlImportJournalMode)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitCreateDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            InitCreateSingleModeDet()
            InitCreateJournalModeDet()

            If (IsSingleEntryMode) Then
                _xmlImport = _xmlImportSingleMode
            Else
                _xmlImport = _xmlImportJournalMode
            End If

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitFindDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlExport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <VERSION>1</VERSION>" & vbCrLf &
                         " <TALLYREQUEST>Export</TALLYREQUEST>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <TYPE>OBJECT</TYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <SUBTYPE>Voucher</SUBTYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <ID TYPE=""Name"">" & _tagMasterId.Name & "</ID>" & vbCrLf

            _xmlExport = _xmlExport &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         "  <DESC>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <STATICVARIABLES>" & vbCrLf &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf &
                         "    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>" & vbCrLf &
                         "   </STATICVARIABLES>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <FETCHLIST>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherNo.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagDate.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherTypeOrigName.Id & "</FETCH>" & vbCrLf &
                         "   </FETCHLIST>" & vbCrLf

            _xmlExport = _xmlExport &
                         "  </DESC>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlExport = UCase(_xmlExport)

            Cursor.Current = mpointer_lcl

        End Sub

        Public Function AddParticulars(ByVal Name_par As Object, ByVal Amount_par As Object) As Boolean

            Dim ret_val As Boolean = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            If (Len(Name_par) = 0 Or Amount_par = 0) Then
                '                GoTo end_func
            End If


            Try

                With _lstParticulars

                    .Add(New List(Of clsXmlTag))

                    With .Item(.Count - 1)

                        .Add(New clsXmlTag)

                        .Item(.Count - 1) = _tagParticulars.Clone

                        With .Item(.Count - 1)
                            If (Len(Name_par) > 0) Then .Data = Name_par
                        End With


                        .Add(New clsXmlTag)

                        .Item(.Count - 1) = _tagAmount.Clone

                        With .Item(.Count - 1)
                            If (Len(Amount_par) > 0) Then .Data = Amount_par
                        End With

                    End With

                End With

                ret_val = True

            Catch

            End Try

end_func:
            Cursor.Current = mpointer_lcl

            AddParticulars = ret_val

        End Function

        Public Function CreateDet(ByVal CurCompany_par As String,
                                  ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim create_xml_det As String = _xmlImport
            Dim create_particulars_xml_det As String = ""

            Dim party_ledger_amt As Single = 0


            With _lstParticulars

                For parti_no As Integer = 0 To .Count - 1

                    create_particulars_xml_det = create_particulars_xml_det &
                                                 PrepXmlFromText(_xmlImportParticulars, .Item(parti_no))

                    party_ledger_amt += .Item(parti_no).Item(amt_parti_index).Data

                Next


                If (.Count > 0) Then

                    _tagPartyLedger.Data = .Item(0).Item(name_parti_index).Data
                    _tagAmount.Data = party_ledger_amt

                    If (Not IsSingleEntryMode) Then
                        _tagParticulars.DataPfx_2 = _tagPartyLedger.Data
                    End If

                End If

            End With


            _tagCurCompany.Data = CurCompany_par

            create_xml_det = PrepXmlFromText(create_xml_det, _tagCurCompany)

            create_xml_det = Replace(create_xml_det, "_xmlImportParticulars",
                                     create_particulars_xml_det, 1, -1, CompareMethod.Text)

            frm_master_det.TextBox1.Text = PrepXmlFromText(create_xml_det, _lstTagDet)

            Cursor.Current = mpointer_lcl

            '            CreateDet = PostXml(PrepXmlFromText(create_xml_det, _lstTagDet), Url_par, Port_par)

            _lstParticulars.Clear()

        End Function

        Public Function FindDet(ByVal CurCompany_par As String,
                                ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim find_xml_det As String = _xmlExport

            _tagCurCompany.Data = CurCompany_par

            find_xml_det = PrepXmlFromText(find_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            FindDet = PostXml(PrepXmlFromText(find_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Sub New()

        XmlDet = New clsXmlDet

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim mem_stream As New MemoryStream()
        Dim bin_formatter As New BinaryFormatter()

        bin_formatter.Serialize(mem_stream, Me)
        mem_stream.Seek(0, SeekOrigin.Begin)

        Return bin_formatter.Deserialize(mem_stream)

    End Function

End Class

<Serializable()>
Public Class clsPayment_Tally

    Implements ICloneable

    Public XmlDet As clsXmlDet

    <Serializable()>
    Class clsXmlDet

        Implements ICloneable

        Dim _tagCurCompany As clsXmlTag

        Public IsSingleEntryMode As Boolean

        Public _tagVoucherTypeOrigName As clsXmlTag
        Public _tagVoucherNo As clsXmlTag
        Public _tagRefNo As clsXmlTag
        Public _tagDate As clsXmlTag
        Public _tagEffectiveDate As clsXmlTag
        Public _tagAccount As clsXmlTag
        Public _tagAmount As clsXmlTag
        Public _tagPartyLedger As clsXmlTag
        Public _tagParticulars As clsXmlTag
        Public _tagNarration As clsXmlTag
        Public _tagMasterId As clsXmlTag

        Public _lstTagDet As New List(Of clsXmlTag)

        Public XmlParseTag As Object

        ' Particulars

        Public _lstParticulars As New List(Of List(Of clsXmlTag))

        Dim name_parti_index As Integer = 0
        Dim amt_parti_index As Integer = 1

        Dim _xmlImportSingleMode As String
        Dim _xmlImportJournalMode As String
        Dim _xmlImportParticulars As String

        Dim _xmlImport As String
        Dim _xmlExport As String

        Public Sub New()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            XmlParseTag = {"", ""}

            _tagVoucherTypeOrigName = New clsXmlTag
            _tagVoucherTypeOrigName = tagdefnVoucherTypeOrigName.Clone
            _lstTagDet.Add(_tagVoucherTypeOrigName)

            _tagVoucherNo = New clsXmlTag
            _tagVoucherNo = tagdefnVoucherNo.Clone
            _lstTagDet.Add(_tagVoucherNo)

            _tagRefNo = New clsXmlTag
            _tagRefNo = tagdefnRefNo.Clone
            With _tagRefNo
                .DataPfx_1 = "PAID"
                .DataPfx_2 = ""
            End With
            _lstTagDet.Add(_tagRefNo)

            _tagDate = New clsXmlTag
            _tagDate = tagdefnDate.Clone
            _lstTagDet.Add(_tagDate)

            _tagEffectiveDate = New clsXmlTag
            _tagEffectiveDate = tagdefnEffectiveDate.Clone
            _lstTagDet.Add(_tagEffectiveDate)

            _tagAccount = New clsXmlTag
            _tagAccount = tagdefnAccount.Clone
            _lstTagDet.Add(_tagAccount)

            _tagAmount = New clsXmlTag
            _tagAmount = tagdefnAmount.Clone
            _lstTagDet.Add(_tagAmount)

            _tagPartyLedger = New clsXmlTag
            _tagPartyLedger = tagdefnPartyLedger.Clone
            _lstTagDet.Add(_tagPartyLedger)

            _tagParticulars = New clsXmlTag
            _tagParticulars = tagdefnParticulars.Clone
            _lstTagDet.Add(_tagParticulars)

            _tagNarration = New clsXmlTag
            _tagNarration = tagdefnNarration.Clone
            _lstTagDet.Add(_tagNarration)

            _tagMasterId = New clsXmlTag
            _tagMasterId = tagdefnMasterId.Clone
            _lstTagDet.Add(_tagMasterId)

            InitOthDet()
            InitCreateDet()
            InitFindDet()

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitOthDet()

            _tagCurCompany = New clsXmlTag
            With _tagCurCompany
                .Name = "tag_cur_company"
                .Id = "svcurrentcompany"
                .DataType = EnumDataType.Text
            End With

            IsSingleEntryMode = False

        End Sub

        Private Sub InitCreateSingleModeDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImportParticulars = _xmlImportParticulars &
                                     "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                     "      <LEDGERNAME>" & _tagParticulars.Name & "</LEDGERNAME>" & vbCrLf &
                                     "      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                                     "      <AMOUNT>" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                     "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportParticulars = UCase(_xmlImportParticulars)


            _xmlImportSingleMode = "<ENVELOPE>" & vbCrLf &
                                   "<HEADER>" & vbCrLf &
                                   " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                                   "</HEADER>" & vbCrLf &
                                   "<BODY>" & vbCrLf &
                                   " <IMPORTDATA>" & vbCrLf &
                                   "  <REQUESTDESC>" & vbCrLf &
                                   "   <REPORTNAME>Vouchers</REPORTNAME>" & vbCrLf &
                                   "   <STATICVARIABLES>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "   </STATICVARIABLES>" & vbCrLf &
                                   "  </REQUESTDESC>" & vbCrLf &
                                   "  <REQUESTDATA>" & vbCrLf

            '            _xmlImportSingleMode = _xmlImportSingleMode & _
            '                                   "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "   <TALLYMESSAGE>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "    <VOUCHER VCHTYPE=""Payment"" ACTION=""Create"" OBJView=""Accounting Voucher View"">" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <DATE>" & _tagDate.Name & "</DATE>" & vbCrLf &
                                   "     <REFERENCE>" & _tagRefNo.Name & "</REFERENCE>" & vbCrLf &
                                   "     <NARRATION>" & _tagNarration.Name & "</NARRATION>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <PARTYLEDGERNAME>" & _tagPartyLedger.Name & "</PARTYLEDGERNAME>" & vbCrLf &
                                   "     <EFFECTIVEDATE>" & _tagEffectiveDate.Name & "</EFFECTIVEDATE>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "_xmlImportParticulars"

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                   "       <LEDGERNAME>" & _tagAccount.Name & "</LEDGERNAME>" & vbCrLf &
                                   "       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                   "       <AMOUNT>-" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                   "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportSingleMode = _xmlImportSingleMode &
                                   "    </VOUCHER>" & vbCrLf &
                                   "   </TALLYMESSAGE>" & vbCrLf &
                                   "  </REQUESTDATA>" & vbCrLf &
                                   " </IMPORTDATA>" & vbCrLf &
                                   "</BODY>" & vbCrLf &
                                   "</ENVELOPE>"

            _xmlImportSingleMode = UCase(_xmlImportSingleMode)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitCreateJournalModeDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImportJournalMode = "<ENVELOPE>" & vbCrLf &
                                    "<HEADER>" & vbCrLf &
                                    " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                                    "</HEADER>" & vbCrLf &
                                    "<BODY>" & vbCrLf &
                                    " <IMPORTDATA>" & vbCrLf &
                                    "  <REQUESTDESC>" & vbCrLf &
                                    "   <REPORTNAME>Vouchers</REPORTNAME>" & vbCrLf &
                                    "   <STATICVARIABLES>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "   </STATICVARIABLES>" & vbCrLf &
                                    "  </REQUESTDESC>" & vbCrLf &
                                    "  <REQUESTDATA>" & vbCrLf

            '            _xmlImportJournalMode = _xmlImportJournalMode & _
            '                                    "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "   <TALLYMESSAGE>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "    <VOUCHER VCHTYPE=""Payment"" ACTION=""Create"" OBJView=""Accounting Voucher View"">" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <DATE>" & _tagDate.Name & "</DATE>" & vbCrLf &
                                    "     <REFERENCE>" & _tagRefNo.Name & "</REFERENCE>" & vbCrLf &
                                    "     <NARRATION>" & _tagNarration.Name & "</NARRATION>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <VOUCHERTYPENAME>Payment</VOUCHERTYPENAME>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <PARTYLEDGERNAME>" & _tagPartyLedger.Name & "</PARTYLEDGERNAME>" & vbCrLf &
                                    "     <EFFECTIVEDATE>" & _tagEffectiveDate.Name & "</EFFECTIVEDATE>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                    "      <LEDGERNAME>" & _tagParticulars.Name & "</LEDGERNAME>" & vbCrLf &
                                    "      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                                    "      <AMOUNT>" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                    "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "     <ALLLEDGERENTRIES.LIST>" & vbCrLf &
                                    "       <LEDGERNAME>" & _tagAccount.Name & "</LEDGERNAME>" & vbCrLf &
                                    "       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                    "       <AMOUNT>-" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                                    "     </ALLLEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportJournalMode = _xmlImportJournalMode &
                                    "    </VOUCHER>" & vbCrLf &
                                    "   </TALLYMESSAGE>" & vbCrLf &
                                    "  </REQUESTDATA>" & vbCrLf &
                                    " </IMPORTDATA>" & vbCrLf &
                                    "</BODY>" & vbCrLf &
                                    "</ENVELOPE>"

            _xmlImportJournalMode = UCase(_xmlImportJournalMode)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitCreateDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            InitCreateSingleModeDet()
            InitCreateJournalModeDet()

            If (IsSingleEntryMode) Then
                _xmlImport = _xmlImportSingleMode
            Else
                _xmlImport = _xmlImportJournalMode
            End If

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitFindDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlExport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <VERSION>1</VERSION>" & vbCrLf &
                         " <TALLYREQUEST>Export</TALLYREQUEST>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <TYPE>OBJECT</TYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <SUBTYPE>Voucher</SUBTYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <ID TYPE=""Name"">" & _tagMasterId.Name & "</ID>" & vbCrLf

            _xmlExport = _xmlExport &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         "  <DESC>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <STATICVARIABLES>" & vbCrLf &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf &
                         "    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>" & vbCrLf &
                         "   </STATICVARIABLES>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <FETCHLIST>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherNo.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagDate.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherTypeOrigName.Id & "</FETCH>" & vbCrLf &
                         "   </FETCHLIST>" & vbCrLf

            _xmlExport = _xmlExport &
                         "  </DESC>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlExport = UCase(_xmlExport)

            Cursor.Current = mpointer_lcl

        End Sub

        Public Function AddParticulars(ByVal Name_par As Object, ByVal Amount_par As Object) As Boolean

            Dim ret_val As Boolean = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            If (Len(Name_par) = 0 Or Amount_par = 0) Then
                '                GoTo end_func
            End If


            Try

                With _lstParticulars

                    .Add(New List(Of clsXmlTag))

                    With .Item(.Count - 1)

                        .Add(New clsXmlTag)

                        .Item(.Count - 1) = _tagParticulars.Clone

                        With .Item(.Count - 1)
                            If (Len(Name_par) > 0) Then .Data = Name_par
                        End With


                        .Add(New clsXmlTag)

                        .Item(.Count - 1) = _tagAmount.Clone

                        With .Item(.Count - 1)
                            If (Len(Amount_par) > 0) Then .Data = Amount_par
                        End With

                    End With

                End With

                ret_val = True

            Catch

            End Try

end_func:
            Cursor.Current = mpointer_lcl

            AddParticulars = ret_val

        End Function

        Public Function CreateDet(ByVal CurCompany_par As String,
                                  ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim create_xml_det As String = _xmlImport
            Dim create_particulars_xml_det As String = ""

            Dim party_ledger_amt As Single = 0


            With _lstParticulars

                For parti_no As Integer = 0 To .Count - 1

                    create_particulars_xml_det = create_particulars_xml_det &
                                                 PrepXmlFromText(_xmlImportParticulars, .Item(parti_no))

                    party_ledger_amt += .Item(parti_no).Item(amt_parti_index).Data

                Next


                If (.Count > 0) Then

                    _tagPartyLedger.Data = .Item(0).Item(name_parti_index).Data
                    _tagAmount.Data = party_ledger_amt

                    If (Not IsSingleEntryMode) Then
                        _tagParticulars.DataPfx_2 = _tagPartyLedger.Data
                    End If

                End If

            End With


            _tagCurCompany.Data = CurCompany_par

            create_xml_det = PrepXmlFromText(create_xml_det, _tagCurCompany)

            create_xml_det = Replace(create_xml_det, "_xmlImportParticulars",
                                     create_particulars_xml_det, 1, -1, CompareMethod.Text)

            Cursor.Current = mpointer_lcl

            CreateDet = PostXml(PrepXmlFromText(create_xml_det, _lstTagDet), Url_par, Port_par)

            _lstParticulars.Clear()

        End Function

        Public Function FindDet(ByVal CurCompany_par As String,
                                ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim find_xml_det As String = _xmlExport

            _tagCurCompany.Data = CurCompany_par

            find_xml_det = PrepXmlFromText(find_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            FindDet = PostXml(PrepXmlFromText(find_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Sub New()

        XmlDet = New clsXmlDet

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim mem_stream As New MemoryStream()
        Dim bin_formatter As New BinaryFormatter()

        bin_formatter.Serialize(mem_stream, Me)
        mem_stream.Seek(0, SeekOrigin.Begin)

        Return bin_formatter.Deserialize(mem_stream)

    End Function

End Class

<Serializable()>
Public Class clsPurchase_Tally

    Implements ICloneable

    Public XmlDet As clsXmlDet

    <Serializable()>
    Class clsXmlDet

        Implements ICloneable

        Dim _tagCurCompany As clsXmlTag

        Public _tagVoucherNo As clsXmlTag
        Public _tagRefNo As clsXmlTag
        Public _tagBasicOrderRefNo As clsXmlTag
        Public _tagDate As clsXmlTag
        Public _tagEffectiveDate As clsXmlTag
        Public _tagAccount As clsXmlTag
        Public _tagAmount As clsXmlTag
        Public _tagPartyLedger As clsXmlTag
        Public _tagLedger As clsXmlTag
        Public _tagLedgName As clsXmlTag
        Public _tagPartyName As clsXmlTag
        Public _tagAddr1 As clsXmlTag
        Public _tagAddr2 As clsXmlTag
        Public _tagNarration As clsXmlTag
        Public _tagMasterId As clsXmlTag

        Public _tagItemName As clsXmlTag
        Public _tagItemRate As clsXmlTag
        Public _tagItemAmount As clsXmlTag
        Public _tagItemActQty As clsXmlTag
        Public _tagItemBillQty As clsXmlTag
        Public _tagItemBatchGodown As clsXmlTag
        Public _tagItemBatchIndentNo As clsXmlTag
        Public _tagItemBatchOrderNo As clsXmlTag
        Public _tagItemBatchTrackingNo As clsXmlTag
        Public _tagItemBatchAmount As clsXmlTag
        Public _tagItemBatchActQty As clsXmlTag
        Public _tagItemBatchBillQty As clsXmlTag
        Public _tagItemAcntAllocTaxClassName As clsXmlTag
        Public _tagItemAcntAllocLedger As clsXmlTag
        Public _tagItemAcntAllocAmount As clsXmlTag

        Public _lstTagDet As New List(Of clsXmlTag)

        Public XmlParseTag As Object

        ' Items Particulars

        Public _lstItemParticulars As New List(Of List(Of clsXmlTag))

        Dim name_parti_index As Integer = 0
        Dim qty_parti_index As Integer = 1
        Dim rate_parti_index As Integer = 2
        Dim amt_parti_index As Integer = 3
        Dim godown_parti_index As Integer = 4
        Dim acnt_class_parti_index As Integer = 5
        Dim acnt_ledger_index As Integer = 6

        Dim _xmlImport As String
        Dim _xmlImportItemParticulars As String

        Dim _xmlExport As String

        Public Sub New()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            XmlParseTag = {"", ""}

            _tagVoucherNo = New clsXmlTag
            _tagVoucherNo = tagdefnVoucherNo.Clone
            _lstTagDet.Add(_tagVoucherNo)

            _tagRefNo = New clsXmlTag
            _tagRefNo = tagdefnRefNo.Clone
            With _tagRefNo
                .DataPfx_1 = "PURC"
                .DataPfx_2 = ""
            End With
            _lstTagDet.Add(_tagRefNo)

            _tagBasicOrderRefNo = New clsXmlTag
            _tagBasicOrderRefNo = tagdefnBasicOrderRefNo.Clone
            _lstTagDet.Add(_tagBasicOrderRefNo)

            _tagDate = New clsXmlTag
            _tagDate = tagdefnDate.Clone
            _lstTagDet.Add(_tagDate)

            _tagEffectiveDate = New clsXmlTag
            _tagEffectiveDate = tagdefnEffectiveDate.Clone
            _lstTagDet.Add(_tagEffectiveDate)

            _tagAccount = New clsXmlTag
            _tagAccount = tagdefnAccount.Clone
            _lstTagDet.Add(_tagAccount)

            _tagAmount = New clsXmlTag
            _tagAmount = tagdefnAmount.Clone
            _lstTagDet.Add(_tagAmount)

            _tagPartyLedger = New clsXmlTag
            _tagPartyLedger = tagdefnPartyLedger.Clone
            _lstTagDet.Add(_tagPartyLedger)

            _tagLedger = New clsXmlTag
            _tagLedger = tagdefnLedger.Clone
            _lstTagDet.Add(_tagLedger)

            _tagLedgName = New clsXmlTag
            _tagLedgName = tagdefnLedger.Clone
            _lstTagDet.Add(_tagLedgName)

            _tagPartyName = New clsXmlTag
            _tagPartyName = tagdefnPartyName.Clone
            _lstTagDet.Add(_tagPartyName)

            _tagAddr1 = New clsXmlTag
            _tagAddr1 = tagdefnAddr1.Clone
            _lstTagDet.Add(_tagAddr1)

            _tagAddr2 = New clsXmlTag
            _tagAddr2 = tagdefnAddr2.Clone
            _lstTagDet.Add(_tagAddr2)

            _tagNarration = New clsXmlTag
            _tagNarration = tagdefnNarration.Clone
            _lstTagDet.Add(_tagNarration)

            _tagMasterId = New clsXmlTag
            _tagMasterId = tagdefnMasterId.Clone
            _lstTagDet.Add(_tagMasterId)

            _tagItemName = New clsXmlTag
            _tagItemName = tagdefnItemName.Clone
            _lstTagDet.Add(_tagItemName)

            _tagItemRate = New clsXmlTag
            _tagItemRate = tagdefnItemRate.Clone
            _lstTagDet.Add(_tagItemRate)

            _tagItemAmount = New clsXmlTag
            _tagItemAmount = tagdefnItemAmount.Clone
            _lstTagDet.Add(_tagItemAmount)

            _tagItemActQty = New clsXmlTag
            _tagItemActQty = tagdefnItemActQty.Clone
            _lstTagDet.Add(_tagItemActQty)

            _tagItemBillQty = New clsXmlTag
            _tagItemBillQty = tagdefnItemBillQty.Clone
            _lstTagDet.Add(_tagItemBillQty)

            _tagItemBatchGodown = New clsXmlTag
            _tagItemBatchGodown = tagdefnItemBatchGodown.Clone
            _lstTagDet.Add(_tagItemBatchGodown)

            _tagItemBatchIndentNo = New clsXmlTag
            _tagItemBatchIndentNo = tagdefnItemBatchIndentNo.Clone
            _lstTagDet.Add(_tagItemBatchIndentNo)

            _tagItemBatchOrderNo = New clsXmlTag
            _tagItemBatchOrderNo = tagdefnItemBatchOrderNo.Clone
            _lstTagDet.Add(_tagItemBatchOrderNo)

            _tagItemBatchTrackingNo = New clsXmlTag
            _tagItemBatchTrackingNo = tagdefnItemBatchTrackingNo.Clone
            _lstTagDet.Add(_tagItemBatchTrackingNo)

            _tagItemBatchAmount = New clsXmlTag
            _tagItemBatchAmount = tagdefnItemBatchAmount.Clone
            _lstTagDet.Add(_tagItemBatchAmount)

            _tagItemBatchActQty = New clsXmlTag
            _tagItemBatchActQty = tagdefnItemBatchActQty.Clone
            _lstTagDet.Add(_tagItemBatchActQty)

            _tagItemBatchBillQty = New clsXmlTag
            _tagItemBatchBillQty = tagdefnItemBatchBillQty.Clone
            _lstTagDet.Add(_tagItemBatchBillQty)

            _tagItemAcntAllocTaxClassName = New clsXmlTag
            _tagItemAcntAllocTaxClassName = tagdefnItemAcntAllocTaxClassName.Clone
            _lstTagDet.Add(_tagItemAcntAllocTaxClassName)

            _tagItemAcntAllocLedger = New clsXmlTag
            _tagItemAcntAllocLedger = tagdefnItemAcntAllocLedger.Clone
            _lstTagDet.Add(_tagItemAcntAllocLedger)

            _tagItemAcntAllocAmount = New clsXmlTag
            _tagItemAcntAllocAmount = tagdefnItemAcntAllocAmount.Clone
            _lstTagDet.Add(_tagItemAcntAllocAmount)

            InitOthDet()
            InitCreateDet()
            InitFindDet()

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitOthDet()

            _tagCurCompany = New clsXmlTag
            With _tagCurCompany
                .Name = "tag_cur_company"
                .Id = "svcurrentcompany"
                .DataType = EnumDataType.Text
            End With

        End Sub

        Private Sub InitCreateDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImportItemParticulars = _xmlImportItemParticulars &
                                        "<ALLINVENTORYENTRIES.LIST>" & vbCrLf &
                                        " <STOCKITEMNAME>" & "" & "</STOCKITEMNAME>" & vbCrLf &
                                        " <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                                        " <RATE>" & "" & "/Grm" & "</RATE>" & vbCrLf &
                                        " <AMOUNT>-" & "" & "</AMOUNT>" & vbCrLf &
                                        " <ACTUALQTY> " & "" & " Grm" & "</ACTUALQTY>" & vbCrLf &
                                        " <BILLEDQTY> " & "" & " Grm" & "</BILLEDQTY>" & vbCrLf &
                                        " <BATCHALLOCATIONS.LIST>" & vbCrLf &
                                        "  <GODOWNNAME>" & "" & "</GODOWNNAME>" & vbCrLf &
                                        "  <INDENTNO/>" & vbCrLf &
                                        "  <ORDERNO/>" & vbCrLf &
                                        "  <TRACKINGNUMBER/>" & vbCrLf &
                                        "  <AMOUNT>-" & "" & "</AMOUNT>" & vbCrLf &
                                        "  <ACTUALQTY> " & "" & " Grm" & "</ACTUALQTY>" & vbCrLf &
                                        "  <BILLEDQTY> " & "" & " Grm" & "</BILLEDQTY>" & vbCrLf &
                                        " </BATCHALLOCATIONS.LIST>" & vbCrLf &
                                        " <ACCOUNTINGALLOCATIONS.LIST>" & vbCrLf &
                                        "  <TAXCLASSIFICATIONNAME/>" & vbCrLf &
                                        "  <LEDGERNAME>" & "" & "</LEDGERNAME>" & vbCrLf &
                                        "  <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                                        "  <AMOUNT>-" & "" & "</AMOUNT>" & vbCrLf &
                                        " </ACCOUNTINGALLOCATIONS.LIST>" & vbCrLf &
                                        "</ALLINVENTORYENTRIES.LIST>" & vbCrLf

            _xmlImportItemParticulars = UCase(_xmlImportItemParticulars)


            _xmlImport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         " <IMPORTDATA>" & vbCrLf &
                         "  <REQUESTDESC>" & vbCrLf &
                         "   <REPORTNAME>Vouchers</REPORTNAME>" & vbCrLf &
                         "   <STATICVARIABLES>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   </STATICVARIABLES>" & vbCrLf &
                         "  </REQUESTDESC>" & vbCrLf &
                         "  <REQUESTDATA>" & vbCrLf

            '            _xmlImport = _xmlImport & _
            '                         "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImport = _xmlImport &
                         "   <TALLYMESSAGE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <VOUCHER VCHTYPE=""Purchase"" ACTION=""Create"" OBJView=""Invoice Voucher View"">" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <DATE>" & _tagDate.Name & "</DATE>" & vbCrLf &
                         "     <REFERENCE>" & _tagRefNo.Name & "</REFERENCE>" & vbCrLf &
                         "     <NARRATION>" & _tagNarration.Name & "</NARRATION>" & vbCrLf &
                         "     <PARTYNAME>" & _tagPartyName.Name & "</PARTYNAME>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <VOUCHERTYPENAME>Purchase</VOUCHERTYPENAME>" & vbCrLf &
                         "     <VOUCHERNUMBER>" & _tagVoucherNo.Name & "</VOUCHERNUMBER>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <PARTYLEDGERNAME>" & _tagPartyLedger.Name & "</PARTYLEDGERNAME>" & vbCrLf &
                         "     <PERSISTEDView>Invoice Voucher View</PERSISTEDView>" & vbCrLf &
                         "     <BASICORDERREF>" & _tagBasicOrderRefNo.Name & "</BASICORDERREF>" & vbCrLf &
                         "     <EFFECTIVEDATE>" & _tagEffectiveDate.Name & "</EFFECTIVEDATE>" & vbCrLf &
                         "     <ISINVOICE>Yes</ISINVOICE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <LEDGERENTRIES.LIST>" & vbCrLf &
                         "      <LEDGERNAME>" & _tagLedger.Name & "</LEDGERNAME>" & vbCrLf &
                         "      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                         "      <AMOUNT>" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                         "     </LEDGERENTRIES.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "_xmlImportItemParticulars"

            _xmlImport = _xmlImport &
                         "    </VOUCHER>" & vbCrLf &
                         "   </TALLYMESSAGE>" & vbCrLf &
                         "  </REQUESTDATA>" & vbCrLf &
                         " </IMPORTDATA>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlImport = UCase(_xmlImport)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitFindDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlExport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <VERSION>1</VERSION>" & vbCrLf &
                         " <TALLYREQUEST>Export</TALLYREQUEST>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <TYPE>OBJECT</TYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <SUBTYPE>Voucher</SUBTYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <ID TYPE=""Name"">" & _tagMasterId.Name & "</ID>" & vbCrLf

            _xmlExport = _xmlExport &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         "  <DESC>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <STATICVARIABLES>" & vbCrLf &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf &
                         "    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>" & vbCrLf &
                         "   </STATICVARIABLES>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <FETCHLIST>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherNo.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagDate.Id & "</FETCH>" & vbCrLf &
                         "   </FETCHLIST>" & vbCrLf

            _xmlExport = _xmlExport &
                         "  </DESC>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlExport = UCase(_xmlExport)

            Cursor.Current = mpointer_lcl

        End Sub

        Public Function AddItemParticulars(ByVal Name_par As Object,
                                           ByVal ActQty_par As Object,
                                           ByVal Rate_par As Object,
                                           ByVal Amount_par As Object,
                                           ByVal Godown_par As Object,
                                           ByVal AcntAllocTaxClassName_par As Object,
                                           ByVal AcntAllocLedger_par As Object) As Boolean

            Dim ret_val As Boolean = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            If (Len(Name_par) = 0 Or Amount_par = 0) Then
                '                GoTo end_func
            End If


            Try

                With _lstItemParticulars

                    .Add(New List(Of clsXmlTag))

                    With .Item(.Count - 1)

                        .Add(New clsXmlTag)

                        .Item(.Count - 1) = _tagItemName.Clone

                        With .Item(.Count - 1)
                            If (Len(Name_par) > 0) Then .Data = Name_par
                        End With


                        .Add(New clsXmlTag)

                        .Item(.Count - 1) = _tagAmount.Clone

                        With .Item(.Count - 1)
                            If (Len(Amount_par) > 0) Then .Data = Amount_par
                        End With

                    End With

                End With

                ret_val = True

            Catch

            End Try

end_func:
            Cursor.Current = mpointer_lcl

            AddItemParticulars = ret_val

        End Function

        Public Function CreateDet(ByVal CurCompany_par As String,
                                  ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim create_xml_det As String = _xmlImport
            Dim create_particulars_xml_det As String = ""

            Dim party_ledger_amt As Single = 0


            With _lstItemParticulars

                For parti_no As Integer = 0 To .Count - 1

                    create_particulars_xml_det = create_particulars_xml_det &
                                                 PrepXmlFromText(_xmlImportItemParticulars, .Item(parti_no))

                    party_ledger_amt = party_ledger_amt + .Item(parti_no).Item(amt_parti_index).Data

                Next


                If (.Count > 0) Then
                    _tagAmount.Data = party_ledger_amt
                End If

            End With


            _tagCurCompany.Data = CurCompany_par

            create_xml_det = PrepXmlFromText(create_xml_det, _tagCurCompany)

            create_xml_det = Replace(create_xml_det, "_xmlImportItemParticulars",
                                     create_particulars_xml_det, 1, -1, CompareMethod.Text)

            Cursor.Current = mpointer_lcl

            CreateDet = PostXml(PrepXmlFromText(create_xml_det, _lstTagDet), Url_par, Port_par)

            _lstItemParticulars.Clear()

        End Function

        Public Function FindDet(ByVal CurCompany_par As String,
                                ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim find_xml_det As String = _xmlExport

            _tagCurCompany.Data = CurCompany_par

            find_xml_det = PrepXmlFromText(find_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            FindDet = PostXml(PrepXmlFromText(find_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Sub New()

        XmlDet = New clsXmlDet

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim mem_stream As New MemoryStream()
        Dim bin_formatter As New BinaryFormatter()

        bin_formatter.Serialize(mem_stream, Me)
        mem_stream.Seek(0, SeekOrigin.Begin)

        Return bin_formatter.Deserialize(mem_stream)

    End Function

End Class

<Serializable()>
Public Class clsSales_Tally

    Implements ICloneable

    Public XmlDet As clsXmlDet

    <Serializable()>
    Class clsXmlDet

        Implements ICloneable

        Dim _tagCurCompany As clsXmlTag

        Public _tagVoucherNo As clsXmlTag
        Public _tagRefNo As clsXmlTag
        Public _tagDate As clsXmlTag
        Public _tagEffectiveDate As clsXmlTag
        Public _tagAccount As clsXmlTag
        Public _tagAmount As clsXmlTag
        Public _tagPartyLedger As clsXmlTag
        Public _tagLedger As clsXmlTag
        Public _tagLedgName As clsXmlTag
        Public _tagLedgAmount As clsXmlTag
        Public _tagLedgIsPositive As clsXmlTag
        Public _tagPartyName As clsXmlTag
        Public _tagAddr1 As clsXmlTag
        Public _tagAddr2 As clsXmlTag
        Public _tagNarration As clsXmlTag
        Public _tagMasterId As clsXmlTag

        Public _tagBasicBuyerName As clsXmlTag
        Public _tagBasicBuyerAddr1 As clsXmlTag
        Public _tagBasicOrderDate As clsXmlTag
        Public _tagBasicPurchaseOrderNo As clsXmlTag
        Public _tagBasicOrderRefNo As clsXmlTag
        Public _tagBasicOrderTerms As clsXmlTag
        Public _tagBasicBuyerSalesTaxNo As clsXmlTag
        Public _tagBuyerCstNo As clsXmlTag
        Public _tagBasicShippedBy As clsXmlTag
        Public _tagBasicDueDateOfPymt As clsXmlTag
        Public _tagBasicShippingDate As clsXmlTag
        Public _tagBasicShipDeliveryNote As clsXmlTag
        Public _tagBasicShipDocumentNo As clsXmlTag
        Public _tagBasicFinalDestination As clsXmlTag

        Public _tagTaxClassLedgName As clsXmlTag
        Public _tagTaxClassLedgTaxClassName As clsXmlTag
        Public _tagTaxClassLedgBasicRateOfInvoiceTax As clsXmlTag
        Public _tagTaxClassLedgAmount As clsXmlTag
        Public _tagTaxClassLedgVatAssessableValue As clsXmlTag
        Public _tagTaxClassLedgCategory As clsXmlTag
        Public _tagTaxClassLedgTaxType As clsXmlTag
        Public _tagTaxClassLedgTaxName As clsXmlTag
        Public _tagTaxClassLedgPartyLedger As clsXmlTag
        Public _tagTaxClassLedgStockItemName As clsXmlTag
        Public _tagTaxClassLedgSubCategory As clsXmlTag
        Public _tagTaxClassLedgDutyLedger As clsXmlTag
        Public _tagTaxClassLedgTaxRate As clsXmlTag
        Public _tagTaxClassLedgAssessableAmount As clsXmlTag
        Public _tagTaxClassLedgTax As clsXmlTag
        Public _tagTaxClassLedgBilledQty As clsXmlTag

        Public _tagItemName As clsXmlTag
        Public _tagItemRate As clsXmlTag
        Public _tagItemAmount As clsXmlTag
        Public _tagItemActQty As clsXmlTag
        Public _tagItemBillQty As clsXmlTag
        Public _tagItemBatchGodown As clsXmlTag
        Public _tagItemBatchIndentNo As clsXmlTag
        Public _tagItemBatchOrderNo As clsXmlTag
        Public _tagItemBatchTrackingNo As clsXmlTag
        Public _tagItemBatchAmount As clsXmlTag
        Public _tagItemBatchActQty As clsXmlTag
        Public _tagItemBatchBillQty As clsXmlTag
        Public _tagItemAcntAllocTaxClassName As clsXmlTag
        Public _tagItemAcntAllocLedger As clsXmlTag
        Public _tagItemAcntAllocAmount As clsXmlTag

        Public _lstTagDet As New List(Of clsXmlTag)

        Public XmlParseTag As Object

        ' Other Ledgers Particulars

        Public _lstLedgParticulars As New List(Of List(Of clsXmlTag))

        Dim name_ledgparti_index As Integer = 0
        Dim amt_ledgparti_index As Integer = 1
        Dim ispositive_ledgparti_index As Integer = 2

        ' Items Particulars

        Public _lstItemParticulars As New List(Of List(Of clsXmlTag))

        Dim name_parti_index As Integer = 0
        Dim qty_parti_index As Integer = 1
        Dim rate_parti_index As Integer = 2
        Dim amt_parti_index As Integer = 3
        Dim godown_parti_index As Integer = 4
        Dim acnt_class_parti_index As Integer = 5
        Dim acnt_ledger_index As Integer = 6

        Dim _xmlImport As String
        Dim _xmlImportTaxClassLedger As String
        Dim _xmlImportTaxClassLedgerSubCategory As String
        Dim _xmlImportLedgParticulars As String
        Dim _xmlImportItemParticulars As String

        Dim _xmlExport As String

        Public Sub New()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            XmlParseTag = {"", ""}

            _tagVoucherNo = New clsXmlTag
            _tagVoucherNo = tagdefnVoucherNo.Clone
            _lstTagDet.Add(_tagVoucherNo)

            _tagRefNo = New clsXmlTag
            _tagRefNo = tagdefnRefNo.Clone
            With _tagRefNo
                .DataPfx_1 = "SALE"
                .DataPfx_2 = ""
            End With
            _lstTagDet.Add(_tagRefNo)

            _tagDate = New clsXmlTag
            _tagDate = tagdefnDate.Clone
            _lstTagDet.Add(_tagDate)

            _tagEffectiveDate = New clsXmlTag
            _tagEffectiveDate = tagdefnEffectiveDate.Clone
            _lstTagDet.Add(_tagEffectiveDate)

            _tagAccount = New clsXmlTag
            _tagAccount = tagdefnAccount.Clone
            _lstTagDet.Add(_tagAccount)

            _tagAmount = New clsXmlTag
            _tagAmount = tagdefnAmount.Clone
            _lstTagDet.Add(_tagAmount)

            _tagPartyLedger = New clsXmlTag
            _tagPartyLedger = tagdefnPartyLedger.Clone
            _lstTagDet.Add(_tagPartyLedger)

            _tagLedger = New clsXmlTag
            _tagLedger = tagdefnLedger.Clone
            _lstTagDet.Add(_tagLedger)

            _tagLedgName = New clsXmlTag
            _tagLedgName = tagdefnLedger.Clone
            _lstTagDet.Add(_tagLedgName)

            _tagLedgAmount = New clsXmlTag
            _tagLedgAmount = tagdefnLedgAmount.Clone
            _lstTagDet.Add(_tagLedgAmount)

            _tagLedgIsPositive = New clsXmlTag
            _tagLedgIsPositive = tagdefnLedgIsPositive.Clone
            _lstTagDet.Add(_tagLedgIsPositive)

            _tagPartyName = New clsXmlTag
            _tagPartyName = tagdefnPartyName.Clone
            _lstTagDet.Add(_tagPartyName)

            _tagAddr1 = New clsXmlTag
            _tagAddr1 = tagdefnAddr1.Clone
            _lstTagDet.Add(_tagAddr1)

            _tagAddr2 = New clsXmlTag
            _tagAddr2 = tagdefnAddr2.Clone
            _lstTagDet.Add(_tagAddr2)

            _tagNarration = New clsXmlTag
            _tagNarration = tagdefnNarration.Clone
            _lstTagDet.Add(_tagNarration)

            _tagMasterId = New clsXmlTag
            _tagMasterId = tagdefnMasterId.Clone
            _lstTagDet.Add(_tagMasterId)

            '******************************************************************************************************************

            _tagBasicBuyerName = New clsXmlTag
            _tagBasicBuyerName = tagdefnBasicBuyerName.Clone
            _lstTagDet.Add(_tagBasicBuyerName)

            _tagBasicBuyerAddr1 = New clsXmlTag
            _tagBasicBuyerAddr1 = tagdefnBasicBuyerAddr1.Clone
            _lstTagDet.Add(_tagBasicBuyerAddr1)

            _tagBasicOrderDate = New clsXmlTag
            _tagBasicOrderDate = tagdefnBasicOrderDate.Clone
            _lstTagDet.Add(_tagBasicOrderDate)

            _tagBasicPurchaseOrderNo = New clsXmlTag
            _tagBasicPurchaseOrderNo = tagdefnBasicPurchaseOrderNo.Clone
            _lstTagDet.Add(_tagBasicPurchaseOrderNo)

            _tagBasicOrderRefNo = New clsXmlTag
            _tagBasicOrderRefNo = tagdefnBasicOrderRefNo.Clone
            _lstTagDet.Add(_tagBasicOrderRefNo)

            _tagBasicOrderTerms = New clsXmlTag
            _tagBasicOrderTerms = tagdefnBasicOrderTerms.Clone
            _lstTagDet.Add(_tagBasicOrderTerms)

            _tagBasicBuyerSalesTaxNo = New clsXmlTag
            _tagBasicBuyerSalesTaxNo = tagdefnBasicBuyerSalesTaxNo.Clone
            _lstTagDet.Add(_tagBasicBuyerSalesTaxNo)

            _tagBuyerCstNo = New clsXmlTag
            _tagBuyerCstNo = tagdefnBuyerCstNo.Clone
            _lstTagDet.Add(_tagBuyerCstNo)

            _tagBasicShippedBy = New clsXmlTag
            _tagBasicShippedBy = tagdefnBasicShippedBy.Clone
            _lstTagDet.Add(_tagBasicShippedBy)

            _tagBasicDueDateOfPymt = New clsXmlTag
            _tagBasicDueDateOfPymt = tagdefnBasicDueDateOfPymt.Clone
            _lstTagDet.Add(_tagBasicDueDateOfPymt)

            _tagBasicShippingDate = New clsXmlTag
            _tagBasicShippingDate = tagdefnBasicShippingDate.Clone
            _lstTagDet.Add(_tagBasicShippingDate)

            _tagBasicShipDeliveryNote = New clsXmlTag
            _tagBasicShipDeliveryNote = tagdefnBasicShipDeliveryNote.Clone
            _lstTagDet.Add(_tagBasicShipDeliveryNote)

            _tagBasicShipDocumentNo = New clsXmlTag
            _tagBasicShipDocumentNo = tagdefnBasicShipDocumentNo.Clone
            _lstTagDet.Add(_tagBasicShipDocumentNo)

            _tagBasicFinalDestination = New clsXmlTag
            _tagBasicFinalDestination = tagdefnBasicFinalDestination.Clone
            _lstTagDet.Add(_tagBasicFinalDestination)

            '******************************************************************************************************************

            _tagTaxClassLedgName = New clsXmlTag
            _tagTaxClassLedgName = tagdefnTaxClassLedgName.Clone
            _lstTagDet.Add(_tagTaxClassLedgName)

            _tagTaxClassLedgTaxClassName = New clsXmlTag
            _tagTaxClassLedgTaxClassName = tagdefnTaxClassLedgTaxClassName.Clone
            _lstTagDet.Add(_tagTaxClassLedgTaxClassName)

            _tagTaxClassLedgBasicRateOfInvoiceTax = New clsXmlTag
            _tagTaxClassLedgBasicRateOfInvoiceTax = tagdefnTaxClassLedgBasicRateOfInvoiceTax.Clone
            _lstTagDet.Add(_tagTaxClassLedgBasicRateOfInvoiceTax)

            _tagTaxClassLedgAmount = New clsXmlTag
            _tagTaxClassLedgAmount = tagdefnTaxClassLedgAmount.Clone
            _lstTagDet.Add(_tagTaxClassLedgAmount)

            _tagTaxClassLedgVatAssessableValue = New clsXmlTag
            _tagTaxClassLedgVatAssessableValue = tagdefnTaxClassLedgVatAssessableValue.Clone
            _lstTagDet.Add(_tagTaxClassLedgVatAssessableValue)

            _tagTaxClassLedgCategory = New clsXmlTag
            _tagTaxClassLedgCategory = tagdefnTaxClassLedgCategory.Clone
            _lstTagDet.Add(_tagTaxClassLedgCategory)

            _tagTaxClassLedgTaxType = New clsXmlTag
            _tagTaxClassLedgTaxType = tagdefnTaxClassLedgTaxType.Clone
            _lstTagDet.Add(_tagTaxClassLedgTaxType)

            _tagTaxClassLedgTaxName = New clsXmlTag
            _tagTaxClassLedgTaxName = tagdefnTaxClassLedgTaxName.Clone
            _lstTagDet.Add(_tagTaxClassLedgTaxName)

            _tagTaxClassLedgPartyLedger = New clsXmlTag
            _tagTaxClassLedgPartyLedger = tagdefnTaxClassLedgPartyLedger.Clone
            _lstTagDet.Add(_tagTaxClassLedgPartyLedger)

            _tagTaxClassLedgStockItemName = New clsXmlTag
            _tagTaxClassLedgStockItemName = tagdefnTaxClassLedgStockItemName.Clone
            _lstTagDet.Add(_tagTaxClassLedgStockItemName)

            _tagTaxClassLedgSubCategory = New clsXmlTag
            _tagTaxClassLedgSubCategory = tagdefnTaxClassLedgSubCategory.Clone
            _lstTagDet.Add(_tagTaxClassLedgSubCategory)

            _tagTaxClassLedgDutyLedger = New clsXmlTag
            _tagTaxClassLedgDutyLedger = tagdefnTaxClassLedgDutyLedger.Clone
            _lstTagDet.Add(_tagTaxClassLedgDutyLedger)

            _tagTaxClassLedgTaxRate = New clsXmlTag
            _tagTaxClassLedgTaxRate = tagdefnTaxClassLedgTaxRate.Clone
            _lstTagDet.Add(_tagTaxClassLedgTaxRate)

            _tagTaxClassLedgAssessableAmount = New clsXmlTag
            _tagTaxClassLedgAssessableAmount = tagdefnTaxClassLedgAssessableAmount.Clone
            _lstTagDet.Add(_tagTaxClassLedgAssessableAmount)

            _tagTaxClassLedgTax = New clsXmlTag
            _tagTaxClassLedgTax = tagdefnTaxClassLedgTax.Clone
            _lstTagDet.Add(_tagTaxClassLedgTax)

            _tagTaxClassLedgBilledQty = New clsXmlTag
            _tagTaxClassLedgBilledQty = tagdefnTaxClassLedgBilledQty.Clone
            _lstTagDet.Add(_tagTaxClassLedgBilledQty)

            '******************************************************************************************************************

            _tagItemName = New clsXmlTag
            _tagItemName = tagdefnItemName.Clone
            _lstTagDet.Add(_tagItemName)

            _tagItemRate = New clsXmlTag
            _tagItemRate = tagdefnItemRate.Clone
            _lstTagDet.Add(_tagItemRate)

            _tagItemAmount = New clsXmlTag
            _tagItemAmount = tagdefnItemAmount.Clone
            _lstTagDet.Add(_tagItemAmount)

            _tagItemActQty = New clsXmlTag
            _tagItemActQty = tagdefnItemActQty.Clone
            _lstTagDet.Add(_tagItemActQty)

            _tagItemBillQty = New clsXmlTag
            _tagItemBillQty = tagdefnItemBillQty.Clone
            _lstTagDet.Add(_tagItemBillQty)

            _tagItemBatchGodown = New clsXmlTag
            _tagItemBatchGodown = tagdefnItemBatchGodown.Clone
            _lstTagDet.Add(_tagItemBatchGodown)

            _tagItemBatchIndentNo = New clsXmlTag
            _tagItemBatchIndentNo = tagdefnItemBatchIndentNo.Clone
            _lstTagDet.Add(_tagItemBatchIndentNo)

            _tagItemBatchOrderNo = New clsXmlTag
            _tagItemBatchOrderNo = tagdefnItemBatchOrderNo.Clone
            _lstTagDet.Add(_tagItemBatchOrderNo)

            _tagItemBatchTrackingNo = New clsXmlTag
            _tagItemBatchTrackingNo = tagdefnItemBatchTrackingNo.Clone
            _lstTagDet.Add(_tagItemBatchTrackingNo)

            _tagItemBatchAmount = New clsXmlTag
            _tagItemBatchAmount = tagdefnItemBatchAmount.Clone
            _lstTagDet.Add(_tagItemBatchAmount)

            _tagItemBatchActQty = New clsXmlTag
            _tagItemBatchActQty = tagdefnItemBatchActQty.Clone
            _lstTagDet.Add(_tagItemBatchActQty)

            _tagItemBatchBillQty = New clsXmlTag
            _tagItemBatchBillQty = tagdefnItemBatchBillQty.Clone
            _lstTagDet.Add(_tagItemBatchBillQty)

            _tagItemAcntAllocTaxClassName = New clsXmlTag
            _tagItemAcntAllocTaxClassName = tagdefnItemAcntAllocTaxClassName.Clone
            _lstTagDet.Add(_tagItemAcntAllocTaxClassName)

            _tagItemAcntAllocLedger = New clsXmlTag
            _tagItemAcntAllocLedger = tagdefnItemAcntAllocLedger.Clone
            _lstTagDet.Add(_tagItemAcntAllocLedger)

            _tagItemAcntAllocAmount = New clsXmlTag
            _tagItemAcntAllocAmount = tagdefnItemAcntAllocAmount.Clone
            _lstTagDet.Add(_tagItemAcntAllocAmount)

            InitOthDet()
            InitCreateDet()
            InitFindDet()

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitOthDet()

            _tagCurCompany = New clsXmlTag
            With _tagCurCompany
                .Name = "tag_cur_company"
                .Id = "svcurrentcompany"
                .DataType = EnumDataType.Text
            End With


            _xmlImport = ""
            _xmlImportTaxClassLedger = ""
            _xmlImportTaxClassLedgerSubCategory = ""
            _xmlImportLedgParticulars = ""
            _xmlImportItemParticulars = ""

            _xmlExport = ""

        End Sub

        Private Sub InitCreateDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlImportTaxClassLedgerSubCategory = _xmlImportTaxClassLedgerSubCategory &
                                                  "  <SUBCATEGORYALLOCATION.LIST>" & vbCrLf &
                                                  "   <STOCKITEMNAME>" & _tagTaxClassLedgStockItemName.Name & "</STOCKITEMNAME>" & vbCrLf &
                                                  "   <SUBCATEGORY>" & _tagTaxClassLedgSubCategory.Name & "</SUBCATEGORY>" & vbCrLf &
                                                  "   <DUTYLEDGER>" & _tagTaxClassLedgDutyLedger.Name & "</DUTYLEDGER>" & vbCrLf &
                                                  "   <TAXRATE> " & _tagTaxClassLedgTaxRate.Name & "</TAXRATE>" & vbCrLf &
                                                  "   <ASSESSABLEAMOUNT>" & _tagTaxClassLedgAssessableAmount.Name & "</ASSESSABLEAMOUNT>" & vbCrLf &
                                                  "   <TAX>" & _tagTaxClassLedgTax.Name & "</TAX>" & vbCrLf &
                                                  "   <BILLEDQTY> " & _tagTaxClassLedgBilledQty.Name & "</BILLEDQTY>" & vbCrLf &
                                                  "  </SUBCATEGORYALLOCATION.LIST>" & vbCrLf

            _xmlImportTaxClassLedgerSubCategory = UCase(_xmlImportTaxClassLedgerSubCategory)


            _xmlImportTaxClassLedger = _xmlImportTaxClassLedger &
                                       "<LEDGERENTRIES.LIST>" & vbCrLf &
                                       " <BASICRATEOFINVOICETAX.LIST TYPE=""Number"">" & vbCrLf &
                                       "  <BASICRATEOFINVOICETAX> " & _tagTaxClassLedgBasicRateOfInvoiceTax.Name & "</BASICRATEOFINVOICETAX>" & vbCrLf &
                                       " </BASICRATEOFINVOICETAX.LIST>" & vbCrLf &
                                       " <TAXCLASSIFICATIONNAME>" & _tagTaxClassLedgTaxClassName.Name & "</TAXCLASSIFICATIONNAME>" & vbCrLf &
                                       " <LEDGERNAME>" & _tagTaxClassLedgName.Name & "</LEDGERNAME>" & vbCrLf &
                                       " <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                       " <AMOUNT>" & _tagTaxClassLedgAmount.Name & "</AMOUNT>" & vbCrLf &
                                       " <VATASSESSABLEVALUE>" & _tagTaxClassLedgVatAssessableValue.Name & "</VATASSESSABLEVALUE>" & vbCrLf &
                                       " <TAXOBJECTALLOCATIONS.LIST>" & vbCrLf &
                                       " <CATEGORY>" & _tagTaxClassLedgCategory.Name & "</CATEGORY>" & vbCrLf &
                                       " <TAXTYPE>" & _tagTaxClassLedgTaxType.Name & "</TAXTYPE>" & vbCrLf &
                                       " <PARTYLEDGER>" & _tagTaxClassLedgPartyLedger.Name & "</PARTYLEDGER>" & vbCrLf

            _xmlImportTaxClassLedger = _xmlImportTaxClassLedger &
                                       "_xmlImportTaxClassLedgerSubCategory"

            _xmlImportTaxClassLedger = _xmlImportTaxClassLedger &
                                       " </TAXOBJECTALLOCATIONS.LIST>" & vbCrLf &
                                       "</LEDGERENTRIES.LIST>"

            _xmlImportTaxClassLedger = UCase(_xmlImportTaxClassLedger)

            '******************************************************************************************************************

            _xmlImportLedgParticulars = _xmlImportLedgParticulars &
                                        "     <LEDGERENTRIES.LIST>" & vbCrLf &
                                        "      <LEDGERNAME>" & _tagLedgName.Name & "</LEDGERNAME>" & vbCrLf &
                                        "      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                        "      <AMOUNT>" & _tagLedgAmount.Name & "</AMOUNT>" & vbCrLf &
                                        "     </LEDGERENTRIES.LIST>" & vbCrLf

            _xmlImportLedgParticulars = UCase(_xmlImportLedgParticulars)

            '******************************************************************************************************************

            _xmlImportItemParticulars = _xmlImportItemParticulars &
                                        "     <ALLINVENTORYENTRIES.LIST>" & vbCrLf &
                                        "      <STOCKITEMNAME>" & _tagItemName.Name & "</STOCKITEMNAME>" & vbCrLf &
                                        "      <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                        "      <RATE>" & _tagItemRate.Name & "</RATE>" & vbCrLf &
                                        "      <AMOUNT>" & _tagItemAmount.Name & "</AMOUNT>" & vbCrLf &
                                        "      <ACTUALQTY> " & _tagItemActQty.Name & "</ACTUALQTY>" & vbCrLf &
                                        "      <BILLEDQTY> " & _tagItemBillQty.Name & "</BILLEDQTY>" & vbCrLf &
                                        "      <BATCHALLOCATIONS.LIST>" & vbCrLf &
                                        "       <GODOWNNAME>" & _tagItemBatchGodown.Name & "</GODOWNNAME>" & vbCrLf &
                                        "       <INDENTNO>" & _tagItemBatchIndentNo.Name & "</INDENTNO>" & vbCrLf &
                                        "       <ORDERNO>" & _tagItemBatchOrderNo.Name & "</ORDERNO>" & vbCrLf &
                                        "       <TRACKINGNUMBER>" & _tagItemBatchTrackingNo.Name & "</TRACKINGNUMBER>" & vbCrLf &
                                        "       <AMOUNT>" & _tagItemBatchAmount.Name & "</AMOUNT>" & vbCrLf &
                                        "       <ACTUALQTY> " & _tagItemBatchActQty.Name & "</ACTUALQTY>" & vbCrLf &
                                        "       <BILLEDQTY> " & _tagItemBatchBillQty.Name & "</BILLEDQTY>" & vbCrLf &
                                        "      </BATCHALLOCATIONS.LIST>" & vbCrLf &
                                        "      <ACCOUNTINGALLOCATIONS.LIST>" & vbCrLf &
                                        "       <TAXCLASSIFICATIONNAME>" & _tagItemAcntAllocTaxClassName.Name & "</TAXCLASSIFICATIONNAME>" & vbCrLf &
                                        "       <LEDGERNAME>" & _tagItemAcntAllocLedger.Name & "</LEDGERNAME>" & vbCrLf &
                                        "       <ISDEEMEDPOSITIVE>No</ISDEEMEDPOSITIVE>" & vbCrLf &
                                        "       <AMOUNT>" & _tagItemAcntAllocAmount.Name & "</AMOUNT>" & vbCrLf &
                                        "      </ACCOUNTINGALLOCATIONS.LIST>" & vbCrLf &
                                        "     </ALLINVENTORYENTRIES.LIST>" & vbCrLf

            _xmlImportItemParticulars = UCase(_xmlImportItemParticulars)

            '******************************************************************************************************************

            _xmlImport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <TALLYREQUEST>Import Data</TALLYREQUEST>" & vbCrLf &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         " <IMPORTDATA>" & vbCrLf &
                         "  <REQUESTDESC>" & vbCrLf &
                         "   <REPORTNAME>Vouchers</REPORTNAME>" & vbCrLf &
                         "   <STATICVARIABLES>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf

            _xmlImport = _xmlImport &
                         "   </STATICVARIABLES>" & vbCrLf &
                         "  </REQUESTDESC>" & vbCrLf &
                         "  <REQUESTDATA>" & vbCrLf

            '            _xmlImport = _xmlImport & _
            '                         "   <TALLYMESSAGE xmlns:UDF=""TallyUDF"">" & vbCrLf
            _xmlImport = _xmlImport &
                         "   <TALLYMESSAGE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "    <VOUCHER VCHTYPE=""Sales"" ACTION=""Create"" OBJView=""Invoice Voucher View"">" & vbCrLf

            '            _xmlImport = _xmlImport & _
            '                         "     <BASICBUYERADDRESS.LIST TYPE=""String"">" & vbCrLf & _
            '                         "      <BASICBUYERADDRESS>" & _tagBasicBuyerAddr1.Name & "</BASICBUYERADDRESS>" & vbCrLf & _
            '                         "     </BASICBUYERADDRESS.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <BASICORDERTERMS.LIST TYPE=""String"">" & vbCrLf &
                         "      <BASICORDERTERMS>" & _tagBasicOrderTerms.Name & "</BASICORDERTERMS>" & vbCrLf &
                         "     </BASICORDERTERMS.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <BUYERSCSTNUMBER>" & _tagBuyerCstNo.Name & "</BUYERSCSTNUMBER>" & vbCrLf &
                         "     <BASICSHIPPEDBY>" & _tagBasicShippedBy.Name & "</BASICSHIPPEDBY>" & vbCrLf &
                         "     <BASICBUYERNAME>" & _tagBasicBuyerName.Name & "</BASICBUYERNAME>" & vbCrLf &
                         "     <BASICSHIPDOCUMENTNO>" & _tagBasicShipDocumentNo.Name & "</BASICSHIPDOCUMENTNO>" & vbCrLf &
                         "     <BASICFINALDESTINATION>" & _tagBasicFinalDestination.Name & "</BASICFINALDESTINATION>" & vbCrLf &
                         "     <BASICORDERREF>" & _tagBasicOrderRefNo.Name & "</BASICORDERREF>" & vbCrLf &
                         "     <BASICBUYERSSALESTAXNO>" & _tagBasicBuyerSalesTaxNo.Name & "</BASICBUYERSSALESTAXNO>" & vbCrLf &
                         "     <BASICDUEDATEOFPYMT>" & _tagBasicDueDateOfPymt.Name & "</BASICDUEDATEOFPYMT>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <INVOICEDELNOTES.LIST>" & vbCrLf &
                         "      <BASICSHIPPINGDATE>" & _tagBasicShippingDate.Name & "</BASICSHIPPINGDATE>" & vbCrLf &
                         "      <BASICSHIPDELIVERYNOTE>" & _tagBasicShipDeliveryNote.Name & "</BASICSHIPDELIVERYNOTE>" & vbCrLf &
                         "     </INVOICEDELNOTES.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <INVOICEORDERLIST.LIST>" & vbCrLf &
                         "      <BASICORDERDATE>" & _tagBasicOrderDate.Name & "</BASICORDERDATE>" & vbCrLf &
                         "      <BASICPURCHASEORDERNO>" & _tagBasicPurchaseOrderNo.Name & "</BASICPURCHASEORDERNO>" & vbCrLf &
                         "     </INVOICEORDERLIST.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <DATE>" & _tagDate.Name & "</DATE>" & vbCrLf &
                         "     <REFERENCE>" & _tagRefNo.Name & "</REFERENCE>" & vbCrLf &
                         "     <NARRATION>" & _tagNarration.Name & "</NARRATION>" & vbCrLf &
                         "     <PARTYNAME>" & _tagPartyName.Name & "</PARTYNAME>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <VOUCHERTYPENAME>Sales</VOUCHERTYPENAME>" & vbCrLf &
                         "     <VOUCHERNUMBER>" & _tagVoucherNo.Name & "</VOUCHERNUMBER>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <PARTYLEDGERNAME>" & _tagPartyLedger.Name & "</PARTYLEDGERNAME>" & vbCrLf &
                         "     <PERSISTEDView>Invoice Voucher View</PERSISTEDView>" & vbCrLf &
                         "     <EFFECTIVEDATE>" & _tagEffectiveDate.Name & "</EFFECTIVEDATE>" & vbCrLf &
                         "     <ISINVOICE>Yes</ISINVOICE>" & vbCrLf

            _xmlImport = _xmlImport &
                         "     <LEDGERENTRIES.LIST>" & vbCrLf &
                         "      <LEDGERNAME>" & _tagLedger.Name & "</LEDGERNAME>" & vbCrLf &
                         "      <ISDEEMEDPOSITIVE>Yes</ISDEEMEDPOSITIVE>" & vbCrLf &
                         "      <AMOUNT>-" & _tagAmount.Name & "</AMOUNT>" & vbCrLf &
                         "     </LEDGERENTRIES.LIST>" & vbCrLf

            _xmlImport = _xmlImport &
                         "_xmlImportTaxClassLedger"

            _xmlImport = _xmlImport &
                         "_xmlImportLedgParticulars"

            _xmlImport = _xmlImport &
                         "_xmlImportItemParticulars"

            _xmlImport = _xmlImport &
                         "    </VOUCHER>" & vbCrLf &
                         "   </TALLYMESSAGE>" & vbCrLf &
                         "  </REQUESTDATA>" & vbCrLf &
                         " </IMPORTDATA>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlImport = UCase(_xmlImport)

            Cursor.Current = mpointer_lcl

        End Sub

        Private Sub InitFindDet()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _xmlExport = "<ENVELOPE>" & vbCrLf &
                         "<HEADER>" & vbCrLf &
                         " <VERSION>1</VERSION>" & vbCrLf &
                         " <TALLYREQUEST>Export</TALLYREQUEST>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <TYPE>OBJECT</TYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <SUBTYPE>Voucher</SUBTYPE>" & vbCrLf

            _xmlExport = _xmlExport &
                         " <ID TYPE=""Name"">" & _tagMasterId.Name & "</ID>" & vbCrLf

            _xmlExport = _xmlExport &
                         "</HEADER>" & vbCrLf &
                         "<BODY>" & vbCrLf &
                         "  <DESC>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <STATICVARIABLES>" & vbCrLf &
                         "    <SVCURRENTCOMPANY>" & _tagCurCompany.Name & "</SVCURRENTCOMPANY>" & vbCrLf &
                         "    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>" & vbCrLf &
                         "   </STATICVARIABLES>" & vbCrLf

            _xmlExport = _xmlExport &
                         "   <FETCHLIST>" & vbCrLf &
                         "    <FETCH>" & _tagVoucherNo.Id & "</FETCH>" & vbCrLf &
                         "    <FETCH>" & _tagDate.Id & "</FETCH>" & vbCrLf &
                         "   </FETCHLIST>" & vbCrLf

            _xmlExport = _xmlExport &
                         "  </DESC>" & vbCrLf &
                         "</BODY>" & vbCrLf &
                         "</ENVELOPE>"

            _xmlExport = UCase(_xmlExport)

            Cursor.Current = mpointer_lcl

        End Sub

        Public Function AddLedgerParticulars(ByVal Name_par As Object,
                                             ByVal Amount_par As Object) As Boolean

            Dim ret_val As Boolean = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            If (Len(Name_par) = 0 Or Amount_par = 0) Then
                '                GoTo end_func
            End If


            '            Try

            With _lstLedgParticulars

                .Add(New List(Of clsXmlTag))

                With .Item(.Count - 1)

                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagLedgName.Clone

                    With .Item(.Count - 1)
                        If (Len(Name_par) > 0) Then .Data = Name_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagLedgAmount.Clone

                    With .Item(.Count - 1)
                        If (Len(Amount_par) > 0) Then .Data = Amount_par
                    End With

                End With

            End With

            ret_val = True

            '            Catch

            '            End Try

end_func:
            Cursor.Current = mpointer_lcl

            AddLedgerParticulars = ret_val

        End Function

        Public Function AddItemParticulars(ByVal ItemName_par As Object,
                                           ByVal ActQty_par As Object,
                                           ByVal Rate_par As Object,
                                           ByVal Amount_par As Object,
                                           ByVal Godown_par As Object,
                                           ByVal AcntAllocTaxClassName_par As Object,
                                           ByVal AcntAllocLedger_par As Object) As Boolean

            Dim ret_val As Boolean = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            If (Len(ItemName_par) = 0 Or ActQty_par = 0 Or Rate_par = 0 Or Amount_par = 0 Or
                Len(Godown_par) = 0 Or Len(AcntAllocLedger_par)) Then

                '                GoTo end_func

            End If


            '            Try

            With _lstItemParticulars

                .Add(New List(Of clsXmlTag))

                With .Item(.Count - 1)

                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemName.Clone

                    With .Item(.Count - 1)
                        If (Len(ItemName_par) > 0) Then .Data = ItemName_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemActQty.Clone

                    With .Item(.Count - 1)
                        If (Len(ActQty_par) > 0) Then .Data = ActQty_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemRate.Clone

                    With .Item(.Count - 1)
                        If (Len(Rate_par) > 0) Then .Data = Rate_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemAmount.Clone

                    With .Item(.Count - 1)
                        If (Len(Amount_par) > 0) Then .Data = Amount_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemBatchGodown.Clone

                    With .Item(.Count - 1)
                        If (Len(Godown_par) > 0) Then .Data = Godown_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemAcntAllocTaxClassName.Clone

                    With .Item(.Count - 1)
                        If (Len(AcntAllocTaxClassName_par) > 0) Then .Data = AcntAllocTaxClassName_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemAcntAllocLedger.Clone

                    With .Item(.Count - 1)
                        If (Len(AcntAllocLedger_par) > 0) Then .Data = AcntAllocLedger_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemAcntAllocAmount.Clone

                    With .Item(.Count - 1)
                        If (Len(Amount_par) > 0) Then .Data = Amount_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemBillQty.Clone

                    With .Item(.Count - 1)
                        If (Len(ActQty_par) > 0) Then .Data = ActQty_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemBatchActQty.Clone

                    With .Item(.Count - 1)
                        If (Len(ActQty_par) > 0) Then .Data = ActQty_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemBatchBillQty.Clone

                    With .Item(.Count - 1)
                        If (Len(ActQty_par) > 0) Then .Data = ActQty_par
                    End With


                    .Add(New clsXmlTag)

                    .Item(.Count - 1) = _tagItemBatchAmount.Clone

                    With .Item(.Count - 1)
                        If (Len(Amount_par) > 0) Then .Data = Amount_par
                    End With

                End With

            End With

            ret_val = True

            '            Catch

            '            End Try

end_func:
            Cursor.Current = mpointer_lcl

            AddItemParticulars = ret_val

        End Function

        Public Function CreateDet(ByVal CurCompany_par As String,
                                  ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim create_xml_det As String = _xmlImport

            Dim create_taxclass_ledg_xml_det As String = ""
            Dim create_taxclass_ledg_subcategory_xml_det = ""

            Dim create_item_particulars_xml_det As String = ""
            Dim create_ledg_particulars_xml_det As String = ""

            Dim party_ledger_amt As Single = 0


            _tagPartyLedger.Data = _tagLedger.Data


            With _lstItemParticulars

                For parti_no As Integer = 0 To .Count - 1

                    create_item_particulars_xml_det = create_item_particulars_xml_det &
                                                      PrepXmlFromText(_xmlImportItemParticulars, .Item(parti_no))

                    party_ledger_amt += .Item(parti_no).Item(amt_parti_index).Data

                Next


                If (.Count > 0) Then
                    _tagAmount.Data = party_ledger_amt
                End If

            End With


            With _lstLedgParticulars

                For parti_no As Integer = 0 To .Count - 1

                    create_ledg_particulars_xml_det = create_ledg_particulars_xml_det &
                                                      PrepXmlFromText(_xmlImportLedgParticulars, .Item(parti_no))

                    party_ledger_amt += .Item(parti_no).Item(amt_ledgparti_index).Data

                Next


                If (.Count > 0) Then
                    _tagAmount.Data = party_ledger_amt
                End If

            End With


            _tagCurCompany.Data = CurCompany_par

            create_xml_det = PrepXmlFromText(create_xml_det, _tagCurCompany)

            create_xml_det = Replace(create_xml_det, "_xmlImportTaxClassLedgerSubCategory",
                                     create_taxclass_ledg_subcategory_xml_det, 1, -1, CompareMethod.Text)

            create_xml_det = Replace(create_xml_det, "_xmlImportTaxClassLedger",
                                     create_taxclass_ledg_xml_det, 1, -1, CompareMethod.Text)

            create_xml_det = Replace(create_xml_det, "_xmlImportLedgParticulars",
                                     create_ledg_particulars_xml_det, 1, -1, CompareMethod.Text)

            create_xml_det = Replace(create_xml_det, "_xmlImportItemParticulars",
                                     create_item_particulars_xml_det, 1, -1, CompareMethod.Text)

            frm_master_det.TextBox1.Text = PrepXmlFromText(create_xml_det, _lstTagDet)

            Cursor.Current = mpointer_lcl

            '            CreateDet = PostXml(PrepXmlFromText(create_xml_det, _lstTagDet), Url_par, Port_par)

            _lstItemParticulars.Clear()

        End Function

        Public Function FindDet(ByVal CurCompany_par As String,
                                ByVal Url_par As String, ByVal Port_par As String) As String

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim find_xml_det As String = _xmlExport

            _tagCurCompany.Data = CurCompany_par

            find_xml_det = PrepXmlFromText(find_xml_det, _tagCurCompany)

            Cursor.Current = mpointer_lcl

            FindDet = PostXml(PrepXmlFromText(find_xml_det, _lstTagDet), Url_par, Port_par)

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Sub New()

        XmlDet = New clsXmlDet

    End Sub

    Public Function Clone() As Object Implements System.ICloneable.Clone

        Dim mem_stream As New MemoryStream()
        Dim bin_formatter As New BinaryFormatter()

        bin_formatter.Serialize(mem_stream, Me)
        mem_stream.Seek(0, SeekOrigin.Begin)

        Return bin_formatter.Deserialize(mem_stream)

    End Function

End Class

Module TallySoftwareRelated

#Region "Tags"

    Public Const XmlParseByTallyMessage = "tallymessage"
    Public Const XmlParseByName_List = "name.list"
    Public Const XmlParseByAddress_List = "address.list"

    Public tagdefnGrpName As New clsXmlTag
    Public tagdefnName As New clsXmlTag
    Public tagdefnAlias As New clsXmlTag
    Public tagdefnLedgName As New clsXmlTag
    Public tagdefnMailName As New clsXmlTag
    Public tagdefnParent As New clsXmlTag
    Public tagdefnAddr1 As New clsXmlTag
    Public tagdefnAddr2 As New clsXmlTag
    Public tagdefnAddr3 As New clsXmlTag
    Public tagdefnStateName As New clsXmlTag
    Public tagdefnPinCode As New clsXmlTag
    Public tagdefnLedgPhone As New clsXmlTag
    Public tagdefnIncomeTaxNo As New clsXmlTag
    Public tagdefnAffectsStock As New clsXmlTag
    Public tagdefnUseForVat As New clsXmlTag
    Public tagdefnOpeningBaln As New clsXmlTag
    Public tagdefnIsCostCentresOn As New clsXmlTag
    Public tagdefnVoucherTypeOrigName As New clsXmlTag
    Public tagdefnVoucherTypeName As New clsXmlTag
    Public tagdefnVoucherNo As New clsXmlTag
    Public tagdefnRefNo As New clsXmlTag
    Public tagdefnDate As New clsXmlTag
    Public tagdefnEffectiveDate As New clsXmlTag
    Public tagdefnLedger As New clsXmlTag
    Public tagdefnPartyLedger As New clsXmlTag
    Public tagdefnPartyName As New clsXmlTag
    Public tagdefnByParty As New clsXmlTag
    Public tagdefnToParty As New clsXmlTag
    Public tagdefnAccount As New clsXmlTag
    Public tagdefnParticulars As New clsXmlTag
    Public tagdefnAmount As New clsXmlTag
    Public tagdefnLedgAmount As New clsXmlTag
    Public tagdefnIsPositive As New clsXmlTag
    Public tagdefnLedgIsPositive As New clsXmlTag
    Public tagdefnNarration As New clsXmlTag
    Public tagdefnMasterId As New clsXmlTag

    '******************************************************************************************************************

    Public tagdefnBasicBuyerName As New clsXmlTag
    Public tagdefnBasicBuyerAddr1 As New clsXmlTag
    Public tagdefnBasicOrderDate As New clsXmlTag
    Public tagdefnBasicPurchaseOrderNo As New clsXmlTag
    Public tagdefnBasicOrderRefNo As New clsXmlTag
    Public tagdefnBasicOrderTerms As New clsXmlTag
    Public tagdefnBasicBuyerSalesTaxNo As New clsXmlTag
    Public tagdefnBuyerCstNo As New clsXmlTag
    Public tagdefnBasicShippedBy As New clsXmlTag
    Public tagdefnBasicDueDateOfPymt As New clsXmlTag
    Public tagdefnBasicShippingDate As New clsXmlTag
    Public tagdefnBasicShipDeliveryNote As New clsXmlTag
    Public tagdefnBasicShipDocumentNo As New clsXmlTag
    Public tagdefnBasicFinalDestination As New clsXmlTag

    '******************************************************************************************************************

    Public tagdefnTaxClassLedgName As New clsXmlTag
    Public tagdefnTaxClassLedgTaxClassName As New clsXmlTag
    Public tagdefnTaxClassLedgBasicRateOfInvoiceTax As New clsXmlTag
    Public tagdefnTaxClassLedgAmount As New clsXmlTag
    Public tagdefnTaxClassLedgVatAssessableValue As New clsXmlTag
    Public tagdefnTaxClassLedgCategory As New clsXmlTag
    Public tagdefnTaxClassLedgTaxType As New clsXmlTag
    Public tagdefnTaxClassLedgTaxName As New clsXmlTag
    Public tagdefnTaxClassLedgPartyLedger As New clsXmlTag
    Public tagdefnTaxClassLedgStockItemName As New clsXmlTag
    Public tagdefnTaxClassLedgSubCategory As New clsXmlTag
    Public tagdefnTaxClassLedgDutyLedger As New clsXmlTag
    Public tagdefnTaxClassLedgTaxRate As New clsXmlTag
    Public tagdefnTaxClassLedgAssessableAmount As New clsXmlTag
    Public tagdefnTaxClassLedgTax As New clsXmlTag
    Public tagdefnTaxClassLedgBilledQty As New clsXmlTag

    '******************************************************************************************************************

    Public tagdefnItemName As New clsXmlTag
    Public tagdefnItemRate As New clsXmlTag
    Public tagdefnItemAmount As New clsXmlTag
    Public tagdefnItemActQty As New clsXmlTag
    Public tagdefnItemBillQty As New clsXmlTag
    Public tagdefnItemBatchGodown As New clsXmlTag
    Public tagdefnItemBatchIndentNo As New clsXmlTag
    Public tagdefnItemBatchOrderNo As New clsXmlTag
    Public tagdefnItemBatchTrackingNo As New clsXmlTag
    Public tagdefnItemBatchAmount As New clsXmlTag
    Public tagdefnItemBatchActQty As New clsXmlTag
    Public tagdefnItemBatchBillQty As New clsXmlTag
    Public tagdefnItemAcntAllocTaxClassName As New clsXmlTag
    Public tagdefnItemAcntAllocLedger As New clsXmlTag
    Public tagdefnItemAcntAllocAmount As New clsXmlTag

    Public Sub InitTallyTags()

        Dim mpointer_lcl = Cursor.Current
        Cursor.Current = Cursors.WaitCursor

        tagdefnGrpName = New clsXmlTag
        With tagdefnGrpName
            .Name = "tagGrpName"
            .Name = "group name"
            .DataType = EnumDataType.Text
        End With

        tagdefnName = New clsXmlTag
        With tagdefnName
            .Name = "tagName"
            .Id = "name"
            .DataType = EnumDataType.Text
        End With

        tagdefnAlias = New clsXmlTag
        With tagdefnAlias
            .Name = "tagAlias"
            .Id = "name"
            .DataType = EnumDataType.Text
        End With

        tagdefnLedgName = New clsXmlTag
        With tagdefnLedgName
            .Name = "tagLedgName"
            .DataType = EnumDataType.Text
        End With

        tagdefnMailName = New clsXmlTag
        With tagdefnMailName
            .Name = "tagMailName"
            .Id = "mailingname"
            .DataType = EnumDataType.Text
        End With

        tagdefnParent = New clsXmlTag
        With tagdefnParent
            .Name = "tagParent"
            .Id = "parent"
            .DataType = EnumDataType.Text
        End With

        tagdefnAddr1 = New clsXmlTag
        With tagdefnAddr1
            .Name = "tagAddr1"
            .Id = "address"
            .DataType = EnumDataType.Text
        End With

        tagdefnAddr2 = New clsXmlTag
        With tagdefnAddr2
            .Name = "tagAddr2"
            .Id = "address"
            .DataType = EnumDataType.Text
        End With

        tagdefnAddr3 = New clsXmlTag
        With tagdefnAddr3
            .Name = "tagAddr3"
            .Id = "address"
            .DataType = EnumDataType.Text
        End With

        tagdefnStateName = New clsXmlTag
        With tagdefnStateName
            .Name = "tagStateName"
            .Id = "statename"
            .DataType = EnumDataType.Text
        End With

        tagdefnPinCode = New clsXmlTag
        With tagdefnPinCode
            .Name = "tagPinCode"
            .Id = "pincode"
            .DataType = EnumDataType.Text
        End With

        tagdefnLedgPhone = New clsXmlTag
        With tagdefnLedgPhone
            .Name = "tagLedgPhone"
            .Id = "ledgerphone"
            .DataType = EnumDataType.Text
        End With

        tagdefnIncomeTaxNo = New clsXmlTag
        With tagdefnIncomeTaxNo
            .Name = "tagIncomeTaxNo"
            .Id = "incometaxnumber"
            .DataType = EnumDataType.Text
        End With

        tagdefnAffectsStock = New clsXmlTag
        With tagdefnAffectsStock
            .Name = "tagAffectsStock"
            .Id = "affectsstock"
            .DefaData = "No"
            .DataType = EnumDataType.Bool
        End With

        tagdefnUseForVat = New clsXmlTag
        With tagdefnUseForVat
            .Name = "tagUseForVat"
            .Id = "useforvat"
            .DefaData = "No"
            .DataType = EnumDataType.Bool
        End With

        tagdefnOpeningBaln = New clsXmlTag
        With tagdefnOpeningBaln
            .Name = "tagOpeningBaln"
            .Id = "openingbalance"
            .DataType = EnumDataType.Amount
        End With

        tagdefnIsCostCentresOn = New clsXmlTag
        With tagdefnIsCostCentresOn
            .Name = "tagIsCostCentresOn"
            .Id = "iscostcentreson"
            .DefaData = "No"
            .DataType = EnumDataType.Bool
        End With

        tagdefnVoucherTypeOrigName = New clsXmlTag
        With tagdefnVoucherTypeOrigName
            .Name = "tagVoucherTypeOrigName"
            .Id = "vouchertypeorigname"
            .DataType = EnumDataType.Text
        End With

        tagdefnVoucherTypeName = New clsXmlTag
        With tagdefnVoucherTypeName
            .Name = "tagVoucherTypeName"
            .Id = "vouchertypename"
            .DataType = EnumDataType.Text
        End With

        tagdefnVoucherNo = New clsXmlTag
        With tagdefnVoucherNo
            .Name = "tagVoucherNumber"
            .Id = "vouchernumber"
            .DataType = EnumDataType.Text
        End With

        tagdefnRefNo = New clsXmlTag
        With tagdefnRefNo
            .Name = "tagRefNo"
            .Id = "reference"
            .DataType = EnumDataType.Text
        End With

        tagdefnDate = New clsXmlTag
        With tagdefnDate
            .Name = "tagDate"
            .Id = "date"
            .DataType = EnumDataType.Dt
        End With

        tagdefnEffectiveDate = New clsXmlTag
        With tagdefnEffectiveDate
            .Name = "tagEffectiveDate"
            .Id = "effectivedate"
            .DataType = EnumDataType.Dt
        End With

        tagdefnLedger = New clsXmlTag
        With tagdefnLedger
            .Name = "tagLedger"
            .Id = "ledgername"
            .DataType = EnumDataType.Text
        End With

        tagdefnPartyLedger = New clsXmlTag
        With tagdefnPartyLedger
            .Name = "tagPartyLedger"
            .Id = "partyledgername"
            .DataType = EnumDataType.Text
        End With

        tagdefnPartyName = New clsXmlTag
        With tagdefnPartyName
            .Name = "tagPartyName"
            .Id = "partyname"
            .DataType = EnumDataType.Text
        End With

        tagdefnByParty = New clsXmlTag
        With tagdefnByParty
            .Name = "tagByParty"
            .Id = "ledgername"
            .DataType = EnumDataType.Text
        End With

        tagdefnToParty = New clsXmlTag
        With tagdefnToParty
            .Name = "tagToParty"
            .Id = "ledgername"
            .DataType = EnumDataType.Text
        End With

        tagdefnAccount = New clsXmlTag
        With tagdefnAccount
            .Name = "tagAccount"
            .Id = "ledgername"
            .DataType = EnumDataType.Text
        End With

        With tagdefnParticulars
            .Name = "tag_particulars"
            .Id = "ledgername"
            .DataType = EnumDataType.Text
        End With

        tagdefnAmount = New clsXmlTag
        With tagdefnAmount
            .Name = "tagAmount"
            .Id = "amount"
            .DataType = EnumDataType.Amount
        End With

        tagdefnLedgAmount = New clsXmlTag
        With tagdefnLedgAmount
            .Name = "tagLedgAmount"
            .Id = "amount"
            .DataType = EnumDataType.Amount
        End With

        tagdefnIsPositive = New clsXmlTag
        With tagdefnIsPositive
            .Name = "tagIsPositive"
            .Id = "Isdeemedpositive"
            .DataType = EnumDataType.Bool
        End With

        tagdefnLedgIsPositive = New clsXmlTag
        With tagdefnLedgIsPositive
            .Name = "tagLedgIsPositive"
            .Id = "Isdeemedpositive"
            .DataType = EnumDataType.Bool
        End With

        tagdefnNarration = New clsXmlTag
        With tagdefnNarration
            .Name = "tagNarration"
            .Id = "narration"
            .DataType = EnumDataType.Text
        End With

        tagdefnMasterId = New clsXmlTag
        With tagdefnMasterId
            .Name = "tagMasterId"
            .Id = "masterid"
            .DataType = EnumDataType.Text
        End With

        '******************************************************************************************************************

        tagdefnBasicBuyerName = New clsXmlTag
        With tagdefnBasicBuyerName
            .Name = "tagBasicBuyerName"
            .Id = "BASICBUYERNAME"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicBuyerAddr1 = New clsXmlTag
        With tagdefnBasicBuyerAddr1
            .Name = "tagBasicBuyerAddr1"
            .Id = "BASICBUYERADDRESS"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicOrderDate = New clsXmlTag
        With tagdefnBasicOrderDate
            .Name = "tagBasicOrderDate"
            .Id = "BASICORDERDATE"
            .DataType = EnumDataType.Dt
        End With

        tagdefnBasicPurchaseOrderNo = New clsXmlTag
        With tagdefnBasicPurchaseOrderNo
            .Name = "tagBasicPurchaseOrderNo"
            .Id = "BASICPURCHASEORDERNO"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicOrderRefNo = New clsXmlTag
        With tagdefnBasicOrderRefNo
            .Name = "tagBasicOrderRefNo"
            .Id = "basicorderref"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicOrderTerms = New clsXmlTag
        With tagdefnBasicOrderTerms
            .Name = "tagBasicOrderTerms"
            .Id = "BASICORDERTERMS"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicBuyerSalesTaxNo = New clsXmlTag
        With tagdefnBasicBuyerSalesTaxNo
            .Name = "tagBasicBuyerSalesTaxNo"
            .Id = "BASICBUYERSSALESTAXNO"
            .DataType = EnumDataType.Text
        End With

        tagdefnBuyerCstNo = New clsXmlTag
        With tagdefnBuyerCstNo
            .Name = "tagBuyerCstNo"
            .Id = "BUYERSCSTNUMBER"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicShippedBy = New clsXmlTag
        With tagdefnBasicShippedBy
            .Name = "tagBasicShippedBy"
            .Id = "BASICSHIPPEDBY"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicDueDateOfPymt = New clsXmlTag   ' Mode/Terms of Payment
        With tagdefnBasicDueDateOfPymt
            .Name = "tagBasicDueDateOfPymt"
            .Id = "BASICDUEDATEOFPYMT"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicShippingDate = New clsXmlTag
        With tagdefnBasicShippingDate
            .Name = "tagBasicShippingDate"
            .Id = "BASICSHIPPINGDATE"
            .DataType = EnumDataType.Dt
        End With

        tagdefnBasicShipDeliveryNote = New clsXmlTag
        With tagdefnBasicShipDeliveryNote
            .Name = "tagBasicShipDeliveryNote"
            .Id = "BASICSHIPDELIVERYNOTE"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicShipDocumentNo = New clsXmlTag
        With tagdefnBasicShipDocumentNo
            .Name = "tagBasicShipDocumentNo"
            .Id = "BASICSHIPDOCUMENTNO"
            .DataType = EnumDataType.Text
        End With

        tagdefnBasicFinalDestination = New clsXmlTag
        With tagdefnBasicFinalDestination
            .Name = "tagBasicFinalDestination"
            .Id = "BASICFINALDESTINATION"
            .DataType = EnumDataType.Text
        End With

        '******************************************************************************************************************

        tagdefnTaxClassLedgName = New clsXmlTag
        With tagdefnTaxClassLedgName
            .Name = "tagTaxClassLedgName"
            .Id = "ledgername"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgTaxClassName = New clsXmlTag
        With tagdefnTaxClassLedgTaxClassName
            .Name = "tagTaxClassLedgTaxClassName"
            .Id = "TAXCLASSIFICATIONNAME"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgBasicRateOfInvoiceTax = New clsXmlTag
        With tagdefnTaxClassLedgBasicRateOfInvoiceTax
            .Name = "tagTaxClassLedgBasicRateOfInvoiceTax"
            .Id = "BASICRATEOFINVOICETAX"
            .DataType = EnumDataType.Int
        End With

        tagdefnTaxClassLedgAmount = New clsXmlTag
        With tagdefnTaxClassLedgAmount
            .Name = "tagTaxClassLedgAmount"
            .Id = "AMOUNT"
            .DataType = EnumDataType.Amount
        End With

        tagdefnTaxClassLedgVatAssessableValue = New clsXmlTag
        With tagdefnTaxClassLedgVatAssessableValue
            .Name = "tagTaxClassLedgVatAssessableValue"
            .Id = "VATASSESSABLEVALUE"
            .DataType = EnumDataType.Amount
        End With

        tagdefnTaxClassLedgCategory = New clsXmlTag
        With tagdefnTaxClassLedgCategory
            .Name = "tagTaxClassLedgCategory"
            .Id = "CATEGORY"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgTaxType = New clsXmlTag
        With tagdefnTaxClassLedgTaxType
            .Name = "tagTaxClassLedgTaxType"
            .Id = "TAXTYPE"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgTaxName = New clsXmlTag
        With tagdefnTaxClassLedgTaxName
            .Name = "tagTaxClassLedgTaxName"
            .Id = "TAXNAME"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgPartyLedger = New clsXmlTag
        With tagdefnTaxClassLedgPartyLedger
            .Name = "tagTaxClassLedgPartyLedger"
            .Id = "PARTYLEDGER"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgStockItemName = New clsXmlTag
        With tagdefnTaxClassLedgStockItemName
            .Name = "tagTaxClassLedgStockItemName"
            .Id = "STOCKITEMNAME"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgSubCategory = New clsXmlTag
        With tagdefnTaxClassLedgSubCategory
            .Name = "tagTaxClassLedgSubCategory"
            .Id = "SUBCATEGORY"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgDutyLedger = New clsXmlTag
        With tagdefnTaxClassLedgDutyLedger
            .Name = "tagTaxClassLedgDutyLedger"
            .Id = "DUTYLEDGER"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgTaxRate = New clsXmlTag
        With tagdefnTaxClassLedgTaxRate
            .Name = "tagTaxClassLedgTaxRate"
            .Id = "TAXRATE"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgAssessableAmount = New clsXmlTag
        With tagdefnTaxClassLedgAssessableAmount
            .Name = "tagTaxClassLedgAssessableAmount"
            .Id = "ASSESSABLEAMOUNT"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgTax = New clsXmlTag
        With tagdefnTaxClassLedgTax
            .Name = "tagTaxClassLedgTax"
            .Id = "TAX"
            .DataType = EnumDataType.Text
        End With

        tagdefnTaxClassLedgBilledQty = New clsXmlTag
        With tagdefnTaxClassLedgBilledQty
            .Name = "tagTaxClassLedgBilledQty"
            .Id = "BILLEDQTY"
            .DataType = EnumDataType.Weight
        End With

        '******************************************************************************************************************

        tagdefnItemName = New clsXmlTag
        With tagdefnItemName
            .Name = "tagItemName"
            .Id = "stockitemname"
            .DataType = EnumDataType.Text
        End With

        tagdefnItemRate = New clsXmlTag
        With tagdefnItemRate
            .Name = "tagItemRate"
            .Id = "rate"
            .DataType = EnumDataType.Amount
        End With

        tagdefnItemAmount = New clsXmlTag
        With tagdefnItemAmount
            .Name = "tagItemAmount"
            .Id = "amount"
            .DataType = EnumDataType.Amount
        End With

        tagdefnItemActQty = New clsXmlTag
        With tagdefnItemActQty
            .Name = "tagItemActQty"
            .Id = "actualqty"
            .DataType = EnumDataType.Weight
        End With

        tagdefnItemBillQty = New clsXmlTag
        With tagdefnItemBillQty
            .Name = "tagItemBillQty"
            .Id = "billedqty"
            .DataType = EnumDataType.Weight
        End With

        tagdefnItemBatchGodown = New clsXmlTag
        With tagdefnItemBatchGodown
            .Name = "tagItemBatchGodown"
            .Id = "godownname"
            .DataType = EnumDataType.Text
        End With

        tagdefnItemBatchIndentNo = New clsXmlTag
        With tagdefnItemBatchIndentNo
            .Name = "tagItemBatchIndentNo"
            .Id = "indentno"
            .DataType = EnumDataType.Text
        End With

        tagdefnItemBatchOrderNo = New clsXmlTag
        With tagdefnItemBatchOrderNo
            .Name = "tagItemBatchOrderNo"
            .Id = "orderno"
            .DataType = EnumDataType.Text
        End With

        tagdefnItemBatchTrackingNo = New clsXmlTag
        With tagdefnItemBatchTrackingNo
            .Name = "tagItemBatchTrackingNo"
            .Id = "trackingnumber"
            .DataType = EnumDataType.Text
        End With

        tagdefnItemBatchAmount = New clsXmlTag
        With tagdefnItemBatchAmount
            .Name = "tagItemBatchAmount"
            .Id = "amount"
            .DataType = EnumDataType.Amount
        End With

        tagdefnItemBatchActQty = New clsXmlTag
        With tagdefnItemBatchActQty
            .Name = "tagItemBatchActQty"
            .Id = "actualqty"
            .DataType = EnumDataType.Weight
        End With

        tagdefnItemBatchBillQty = New clsXmlTag
        With tagdefnItemBatchBillQty
            .Name = "tagItemBatchBillQty"
            .Id = "billedqty"
            .DataType = EnumDataType.Weight
        End With

        tagdefnItemAcntAllocTaxClassName = New clsXmlTag
        With tagdefnItemAcntAllocTaxClassName
            .Name = "tagItemAcntAllocTaxClassName"
            .Id = "taxclassificationname"
            .DataType = EnumDataType.Text
        End With

        tagdefnItemAcntAllocLedger = New clsXmlTag
        With tagdefnItemAcntAllocLedger
            .Name = "tagItemAcntAllocLedger"
            .Id = "ledgername"
            .DataType = EnumDataType.Text
        End With

        tagdefnItemAcntAllocAmount = New clsXmlTag
        With tagdefnItemAcntAllocAmount
            .Name = "tagItemAcntAllocAmount"
            .Id = "amount"
            .DataType = EnumDataType.Amount
        End With


        Cursor.Current = mpointer_lcl

    End Sub

#End Region

End Module

Module TallyBridgeDatabaseRelated

    <Serializable()>
    Public Class clsMasterDet_TallyBridge

        Implements ICloneable

        Private _IsInitialized As Boolean
        Private _IsBusy As Boolean

        Public _fldCode As String
        Public _fldOthCode As String
        Public _fldEntryTypeNo As String
        Public _fldName As String
        Public _fldName_2 As String
        Public _fldOthName As String
        Public _fldAddr1 As String
        Public _fldAddr2 As String
        Public _fldAddr3 As String
        Public _fldCity As String
        Public _fldPinNo As String
        Public _fldState As String
        Public _fldPhNo As String
        Public _fldCellNo As String
        Public _fldPanNo As String
        Public _fldAccHeadNo As String
        Public _fldDelStat As String

        Dim _Pkey() As Object

        Public _UniqKeyStat As Boolean = False

        Public _DataCount
        Public _lstData As New List(Of clsDbFieldDet)

        Public MsgSaveDet As String
        Public MsgDelDet As String

        Public DlgDelDet As String

        Public _DbTblName As String

        Private _FieldPfx As String

        Public View As String
        Public ViewOrderByDet

        Private _Conn As Object

        Public _oRst As New clsRecordset

        Sub New(ByVal Conn_par As Object)

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _Conn = Conn_par

            _DbTblName = "sss_master_det"
            _FieldPfx = "mast_"

            View = ""
            ViewOrderByDet = ""

            MsgSaveDet = "Master Details Saved......"
            MsgDelDet = "Master Details Deleted . Please Check ......"

            DlgDelDet = "Want to Delete Master Details ?"

            CreateDataList()

            _Pkey = {_fldCode}

            Open()

            _UniqKeyStat = CheckUniqKeyDet()

            InitInstance()

            Cursor.Current = mpointer_lcl

        End Sub

        Protected Overrides Sub Finalize()

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _Conn.Close()

            Cursor.Current = mpointer_lcl

        End Sub

        Public Property IsInitialized() As Boolean

            Get
                '                IsInitialized = _IsInitialized
                IsInitialized = _oRst.IsInitialized
            End Get

            Set(ByVal value As Boolean)
                _IsInitialized = value
            End Set

        End Property

        Public Property IsBusy() As Boolean

            Get
                IsBusy = _IsBusy
            End Get

            Set(ByVal value As Boolean)

                _IsBusy = value
            End Set

        End Property

        Public ReadOnly Property Da() As Object

            Get
                Da = _oRst.Da
            End Get

        End Property

        Public ReadOnly Property Dset() As System.Data.DataSet

            Get
                Dset = _oRst.Dset
            End Get

        End Property

        Public ReadOnly Property CommBuilder() As Object

            Get
                CommBuilder = _oRst.CommBuilder
            End Get

        End Property

        Public ReadOnly Property Dtable() As System.Data.DataTable

            Get
                Dtable = _oRst.Dtable
            End Get

        End Property

        Public ReadOnly Property Drow() As System.Data.DataRow

            Get
                Drow = _oRst.Drow
            End Get

        End Property

        Public ReadOnly Property Bof() As Boolean

            Get
                Bof = _oRst.Bof
            End Get

        End Property

        Public ReadOnly Property Eof() As Boolean

            Get
                Eof = _oRst.Eof
            End Get

        End Property

        Public Property RecNo() As Object

            Get
                RecNo = _oRst.RecNo
            End Get

            Set(ByVal value As Object)
                _oRst.RecNo = value
            End Set

        End Property

        Public Sub Open()

            _oRst.Open(_Conn, _DbTblName, _FieldPfx)

        End Sub

        Public Sub Reopen()

            _oRst.Reopen(_DbTblName, _FieldPfx)

        End Sub

        Public Sub Refresh()

            _oRst.Refresh()

        End Sub

        Private Function CreateDataList() As Boolean

            Dim ret_val As Boolean : ret_val = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            _lstData.Clear()

            _fldCode = _FieldPfx & fldCode
            AddListItem(_lstData, _fldCode)

            _fldOthCode = _FieldPfx & fldOthCode
            AddListItem(_lstData, _fldOthCode)
            SetData(_fldOthCode, 50, EnumDbFieldDet.Length)

            _fldEntryTypeNo = _FieldPfx & fldEntryTypeNo
            AddListItem(_lstData, _fldEntryTypeNo)

            _fldName = _FieldPfx & fldName
            AddListItem(_lstData, _fldName)
            SetData(_fldName, True, EnumDbFieldDet.ImpStat)

            _fldName_2 = _FieldPfx & fldName_2
            AddListItem(_lstData, _fldName_2)

            _fldOthName = _FieldPfx & fldOthName
            AddListItem(_lstData, _fldOthName)

            _fldAddr1 = _FieldPfx & fldAddr1
            AddListItem(_lstData, _fldAddr1)

            _fldAddr2 = _FieldPfx & fldAddr2
            AddListItem(_lstData, _fldAddr2)

            _fldAddr3 = _FieldPfx & fldAddr3
            AddListItem(_lstData, _fldAddr3)

            _fldCity = _FieldPfx & fldCity
            AddListItem(_lstData, _fldCity)

            _fldPinNo = _FieldPfx & fldPinNo
            AddListItem(_lstData, _fldPinNo)

            _fldState = _FieldPfx & fldState
            AddListItem(_lstData, _fldState)

            _fldPhNo = _FieldPfx & fldPhNo
            AddListItem(_lstData, _fldPhNo)

            _fldCellNo = _FieldPfx & fldCellNo
            AddListItem(_lstData, _fldCellNo)

            _fldPanNo = _FieldPfx & fldPanNo
            AddListItem(_lstData, _fldPanNo)

            _fldAccHeadNo = _FieldPfx & fldAccHeadNo
            AddListItem(_lstData, _fldAccHeadNo)
            SetData(_fldAccHeadNo, True, EnumDbFieldDet.ImpStat)

            _fldDelStat = _FieldPfx & fldDelStat
            AddListItem(_lstData, _fldDelStat)

            _DataCount = _lstData.Count

            ret_val = True

end_sub:
            Cursor.Current = mpointer_lcl

            CreateDataList = ret_val

        End Function

        Public Function InitInstance() As Object

            Dim ret_val : ret_val = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            InitData()

            ret_val = True

end_sub:
            Cursor.Current = mpointer_lcl

            InitInstance = ret_val

        End Function

        Private Function CheckUniqKeyDet() As Object

            Dim ret_val : ret_val = False

            Dim mpointer_lcl = Cursor.Current

            '            Try

            If (Not _oRst.IsInitialized) Then GoTo end_sub

            Cursor.Current = Cursors.WaitCursor

            Dim rdr_lcl = Da.SelectCommand.ExecuteReader(CommandBehavior.KeyInfo)

            Dim dtblSchema = rdr_lcl.GetSchemaTable()


            With dtblSchema

                For col_no = 0 To .Rows.Count - 1
                    If (.Rows(col_no).Item("iskey")) Then
                        ret_val = True
                        Exit For
                    End If
                Next

            End With

            rdr_lcl.Close()


            If (Not ret_val) Then

                If (IsNothing(_Pkey)) Then GoTo end_sub

                With Da.SelectCommand

                    '                    Dim _Pkey() = {_FieldPfx & "ID"}
                    '                    .CommandText = "Alter Table " & _DbTblName & " Add Column " & _Pkey(0) & " AutoIncrement"
                    '                    .ExecuteNonQuery(strSQL)

                    .CommandText = "Alter Table " & _DbTblName & " Add Constraint" &
                                   " PrimaryKey Primary Key(" & _Pkey(0) & ")"
                    .ExecuteNonQuery()

                    Reopen()

                    ret_val = _oRst.IsInitialized

                End With

            End If

            '            Catch

            '            End Try

end_sub:
            If (Not ret_val) Then Dim _Pkey() = Nothing

            Cursor.Current = mpointer_lcl

            CheckUniqKeyDet = ret_val

        End Function

        Public Function CheckMandatoryData() As Object

            Dim ret_val : ret_val = False

            Dim mpointer_lcl = Cursor.Current
            Cursor.Current = Cursors.WaitCursor

            Dim Dset_lcl As New System.Data.DataSet

            Dim lstReqdData As New List(Of clsGenData)


            lstReqdData.Add(New clsGenData)

            With lstReqdData(lstReqdData.Count - 1)

                .Code = lstReqdData.Count
                .Text = UCase(Trim(""))
                .OthDet_1 = ""
                .OthDet_2 = True

                Dim DefaVal = Val(.Code)

            End With


            lstReqdData.Add(New clsGenData)

            With lstReqdData(lstReqdData.Count - 1)

                .Code = lstReqdData.Count
                .Text = UCase(Trim("1"))
                .OthDet_1 = ""
                .OthDet_2 = True

                Dim DefaVal = Val(.Code)

            End With


            '            ret_val = CheckDbTblWithCommFieldsForUniqRec(_Conn, _
            '                                                         Dset_lcl, _
            '                                                         _DbTblName, _
            '                                                         lstReqdData, _
            '                                                         _FieldPfx)

end_sub:
            Dset_lcl.Dispose()

            Cursor.Current = mpointer_lcl

            CheckMandatoryData = ret_val

        End Function

        '******************************************************************************************************************

        Public Sub InitData()

            InitDbField(_lstData)

        End Sub

        Public Function RecCount() As Object

            RecCount = _oRst.RecCount

        End Function

        Public Sub MoveFirst()

            _oRst.MoveFirst()

        End Sub

        Public Sub MoveLast()

            _oRst.MoveLast()

        End Sub

        Public Sub MoveNext()

            _oRst.MoveNext()

        End Sub

        Public Sub MovePrevious()

            _oRst.MovePrevious()

        End Sub

        Public Sub MoveTo(ByVal RecNo_par As Object)

            _oRst.MoveTo(RecNo_par)

        End Sub

        Public Function Find(ByVal Key_par As Object, _
                             ByVal KeyVal_par As Object, _
                             Optional ByVal OthOpn_par As EnumDataOpn = EnumDataOpn.None, _
                             Optional ByVal OthOpnMsg_par As Object = True, _
                             Optional ByVal AsyncRunStat_par As Object = False) As Object

            If (AsyncRunStat_par) Then
                Find = RunDelegate(AddressOf Me.Find_Arg, _
                                   {Key_par, KeyVal_par, OthOpn_par, OthOpnMsg_par}, _
                                   AsyncRunStat_par)
            Else
                Find = Me.Find_Arg({Key_par, KeyVal_par, OthOpn_par, OthOpnMsg_par})
            End If

        End Function

        Private Function Find_Arg(ByVal Arg_par As Object) As Object

            Dim proc_stat : proc_stat = False

            Dim Key_lcl As Object = Arg_par(0)
            Dim KeyVal_lcl As Object = Arg_par(1)
            Dim OthOpn_lcl As EnumDataOpn = Arg_par(2)
            Dim OthOpnMsg_lcl As Object = Arg_par(3)

            Reopen()

            _oRst.Find(Key_lcl, KeyVal_lcl)

            proc_stat = Not Eof


            If (OthOpn_lcl = EnumDataOpn.NewItem) Then

                If (Not Eof) Then

                    If (OthOpnMsg_lcl = True) Then DispAppMsg(EnumAppMsgType.DuplicateItem)

                    proc_stat = EnumResult.Found

                Else

                    proc_stat = Add()

                    If (proc_stat) Then
                        _oRst.SetKeyVal(KeyVal_lcl)
                        proc_stat = Update()
                    End If

                End If

            End If

end_sub:
            Find_Arg = proc_stat

        End Function

        Public Function PrepRelCond(ByVal FieldName_par As Object, ByVal RelOprtr_par As Object, _
                                    ByVal FieldVal_par As Object) As Object

            PrepRelCond = _oRst.PrepRelCond(FieldName_par, RelOprtr_par, FieldVal_par)

        End Function

        Public Function SetFilter(ByVal Sql_par As Object) As Object

            '            ReCreate()

            SetFilter = _oRst.SetFilter(Sql_par)

        End Function

        Public Function ResetFilter() As Object

            '            ResetFilter = ReCreate()
            ResetFilter = _oRst.ResetFilter()

        End Function

        Public Function FindInvalidData() As clsDbFieldDet

            Dim ret_val As clsDbFieldDet = Nothing


            With _lstData

                For ele_no = 0 To .Count - 1

                    With .Item(ele_no)

                        If (.ImpStat And IfNull(.Value.Text)) Then

                            ret_val = _lstData.Item(ele_no)

                            Exit For

                        End If

                    End With

                Next ele_no

            End With

end_sub:
            FindInvalidData = ret_val

        End Function

        Public Function GetData(ByVal FieldName_par As Object, _
                                Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

            GetData = GetDbField(_lstData, FieldName_par, DetType_par)

        End Function

        Public Function SetData(ByVal FieldName_par As Object, _
                                ByVal DetVal_par As Object, _
                                Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

            SetData = SetDbField(_lstData, FieldName_par, DetVal_par, DetType_par)

        End Function

        Public Function GetField(ByVal FieldName_par As Object, _
                                 Optional ByVal Drow_par As System.Data.DataRow = Nothing, _
                                 Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

            GetField = _oRst.GetDet(FieldName_par, Drow_par, DetType_par)

        End Function

        Public Function SetField(ByVal LstField_par As List(Of clsDbFieldDet), _
                                 Optional ByVal ExcludeField_par As Object = Nothing, _
                                 Optional ByRef Drow_par As System.Data.DataRow = Nothing, _
                                 Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

            Dim ret_val : ret_val = False

            If (IsNothing(LstField_par)) Then GoTo end_sub


            With LstField_par

                Dim lstFieldName As New List(Of Object)
                Dim lstFieldVal As New List(Of Object)

                Dim lstExcludeField As New List(Of Object)

                If (Not IsNothing(ExcludeField_par)) Then

                    If (IsArray(ExcludeField_par)) Then
                        lstExcludeField.AddRange(ExcludeField_par)
                    Else
                        If (Len(Trim(ExcludeField_par)) > 0) Then lstExcludeField.Add(ExcludeField_par)
                    End If

                End If


                For field_no = 0 To .Count - 1

                    With .Item(field_no)

                        If (lstFieldName.IndexOf(.Name) < 0 And _
                            (lstExcludeField.Count = 0 Or _
                             (lstExcludeField.Count > 0 And lstExcludeField.IndexOf(.Name) < 0))) Then

                            lstFieldName.Add(.Name)
                            lstFieldVal.Add(.Value.Text)

                        End If

                    End With

                Next field_no


                If (lstFieldName.Count > 0) Then
                    ret_val = SetField(lstFieldName.ToArray, lstFieldVal.ToArray, Drow_par, DetType_par)
                End If

            End With

end_sub:
            SetField = ret_val

        End Function

        Public Function SetField(ByVal FieldName_par As Object, _
                                 ByRef FieldVal_par As Object, _
                                 Optional ByRef Drow_par As System.Data.DataRow = Nothing, _
                                 Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

            SetField = _oRst.SetDet(FieldName_par, FieldVal_par, Drow_par, DetType_par)

        End Function

        Public Function Update(Optional ByVal MsgDisp_par As Object = True) As Object

            Dim ret_val : ret_val = False

            If (_oRst.Update()) Then

                If (MsgDisp_par = True) Then DispAppMsg(EnumAppMsgType.General, msg_det_par:=MsgSaveDet)

                ret_val = True

            End If

end_sub:
            Update = ret_val

        End Function

        Public Function Add() As Object

            Dim ret_val : ret_val = False

            Dim FieldName_par = ""
            Dim FieldVal_par = ""

            If (Not IsNothing(_Pkey)) Then
                FieldName_par = fldCode
                FieldVal_par = GetUniqueVal(_FieldPfx)
            End If

            ret_val = _oRst.Add(FieldName_par, FieldVal_par)

end_sub:
            Add = ret_val

        End Function

        Public Function Add(ByVal FieldName_par As Object, _
                            ByVal FieldVal_par As Object) As Object

            Add = _oRst.Add(FieldName_par, FieldVal_par)

        End Function

        Public Function Delete(Optional ByVal MsgDisp_par As Object = True) As Object

            Dim ret_val : ret_val = False

            If (Eof) Then GoTo end_sub

            If (MsgDisp_par = True) Then
                If (DispAppDlg(EnumAppDlgType.YesNo, DlgDelDet) = vbNo) Then GoTo end_sub
            End If

            SetField(_fldDelStat, EnumDataStat.Deleted)

            Delete()
            Update()

            If (MsgDisp_par = True) Then DispAppMsg(EnumAppMsgType.General, msg_det_par:=MsgDelDet)

            ret_val = True

end_sub:
            Delete = ret_val

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

End Module