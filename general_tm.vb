
Imports System.IO
Imports System.ComponentModel
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Xml

Imports System.Threading
Imports System.Threading.Tasks
Imports System.Runtime.Remoting.Messaging

Imports Microsoft.Office.Interop

Imports System.Windows.Forms
Imports System.Drawing
Imports System.Data

Public Module Application_Related

    Public validity_date

    Public computer_name

    Public firm_server

    Public firm_no As Integer
    Public firm_code As Integer
    Public firm_name As String
    Public firm_short_name As String
    Public firm_addr_1 As String
    Public firm_addr_2 As String
    Public firm_phone_1 As String
    Public firm_mesg_id As String

    Public firm_name_caption
    Public firm_name_caption_short

    Public our_name
    Public our_ph_1
    Public our_ph_2
    Public our_mail

    Enum EnumAppMsgType
        ForLicence
        ForModuleActivation
        Validity
        Trial
        Maintanance
        DemoExtension
        General
        DuplicateItem
        IncompleteDetails
        InvalidDetails
    End Enum

    Enum EnumAppDlgType
        YesNo
        ExitDlg
        Quit
        Confirmation
        SaveDetails
        SaveChanges
    End Enum

    Public Sub SetOurDet()

        validity_date = "30/APR/2016"

        our_name = "SRI SOFTWARE SOLUTIONS"
        our_ph_1 = "09885251518"
        our_ph_2 = "09849197707"
        our_mail = "sri.softwaresolutions@yahoo.com"

    End Sub

    Public Sub DispAppMsg(ByVal msg_no_par As EnumAppMsgType,
                          Optional ByVal msg_box_style_par As Object = 0,
                          Optional ByVal msg_title_par As Object = "",
                          Optional ByVal msg_det_par As Object = "")

        Dim mpointer_lcl = Cursor.Current
        Cursor.Current = Cursors.Default

        Dim msg_det = ""

        If (msg_no_par = EnumAppMsgType.ForLicence) Then

            msg_det = "Please Contact the Vendor for Licensed Version" & vbCrLf & vbCrLf &
                      our_name & vbCrLf &
                      "PH   : " & our_ph_1 & " , " & our_ph_2 & vbCrLf &
                      "Mail : " & our_mail

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbCritical
            If (Len(msg_title_par) = 0) Then msg_title_par = "UNAUTHORIZED  USER"


        ElseIf (msg_no_par = EnumAppMsgType.ForModuleActivation) Then

            msg_det = "Please Contact the Vendor for this Operation/Option" & vbCrLf & vbCrLf &
                      our_name & vbCrLf &
                      "PH   : " & our_ph_1 & " , " & our_ph_2 & vbCrLf &
                      "Mail : " & our_mail

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbInformation


        ElseIf (msg_no_par = EnumAppMsgType.Validity) Then

            msg_det = "This Version Is Valid Till : " & Format(msg_det_par, "dd'mmm yyyy") & vbCrLf & vbCrLf &
                      our_name & vbCrLf &
                      "PH   : " & our_ph_1 & " , " & our_ph_2 & vbCrLf &
                      "Mail : " & our_mail

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbInformation


        ElseIf (msg_no_par = EnumAppMsgType.Trial) Then

            msg_det = "This Is A Demo Version And Works For A Limited Period" & vbCrLf & vbCrLf &
                      our_name & vbCrLf &
                      "PH   : " & our_ph_1 & " , " & our_ph_2 & vbCrLf &
                      "Mail : " & our_mail

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbInformation


        ElseIf (msg_no_par = EnumAppMsgType.Maintanance) Then

            msg_det = "This Software Needs Tobe Upgraded . Please Contact the Vendor" & vbCrLf & vbCrLf &
                      our_name & vbCrLf &
                      "PH   : " & our_ph_1 & " , " & our_ph_2 & vbCrLf &
                      "Mail : " & our_mail

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbInformation


        ElseIf (msg_no_par = EnumAppMsgType.DemoExtension) Then

            msg_det = "Please Contact the Vendor For Extension" & vbCrLf & vbCrLf &
                      our_name & vbCrLf &
                      "PH   : " & our_ph_1 & " , " & our_ph_2 & vbCrLf &
                      "Mail : " & our_mail

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbCritical
            If (Len(msg_title_par) = 0) Then msg_title_par = "ALREADY  WORKED  ON  THE  DEMO"


        ElseIf (msg_no_par = EnumAppMsgType.General) Then

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbInformation
            If (Len(msg_title_par) = 0) Then msg_title_par = "ATTENTION"

            msg_det = msg_det_par


        ElseIf (msg_no_par = EnumAppMsgType.DuplicateItem) Then

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbCritical
            If (Len(msg_title_par) = 0) Then msg_title_par = "ATTENTION"

            msg_det = msg_det_par

            If (Len(Trim(msg_det)) = 0) Then msg_det = "Duplicate Entry . Please Check......"


        ElseIf (msg_no_par = EnumAppMsgType.IncompleteDetails) Then

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbCritical
            If (Len(msg_title_par) = 0) Then msg_title_par = "CANNOT  PROCEED"

            msg_det = msg_det_par

            If (Len(Trim(msg_det)) = 0) Then msg_det = "Incomplete Details . Please Check......"


        ElseIf (msg_no_par = EnumAppMsgType.InvalidDetails) Then

            If (msg_box_style_par <= 0) Then msg_box_style_par = vbCritical
            If (Len(msg_title_par) = 0) Then msg_title_par = "CANNOT  PROCEED"

            msg_det = msg_det_par

            If (Len(Trim(msg_det)) = 0) Then msg_det = "Invalid Details . Please Check......"

        End If


        MsgBox(msg_det, msg_box_style_par, msg_title_par)

        Cursor.Current = mpointer_lcl

    End Sub

    Public Function DispAppDlg(ByVal dlg_no_par As EnumAppDlgType,
                               Optional ByVal dlg_det_par As Object = "",
                               Optional ByVal dlg_title_par As Object = "",
                               Optional ByVal defa_button_par As Object = 0) As Object

        Dim mpointer_lcl = Cursor.Current
        Cursor.Current = Cursors.Default

        Dim ret_val = ""


        If (dlg_no_par = EnumAppDlgType.YesNo) Then

            If (Len(Trim(dlg_title_par)) = 0) Then dlg_title_par = "A T T E N T I O N"
            If (defa_button_par = 0) Then defa_button_par = vbDefaultButton2

            ret_val = MsgBox(dlg_det_par, vbQuestion + vbYesNo + defa_button_par, dlg_title_par)


        ElseIf (dlg_no_par = EnumAppDlgType.ExitDlg) Then

            If (Len(Trim(dlg_det_par)) = 0) Then dlg_det_par = "Want To Exit ?"
            If (Len(Trim(dlg_title_par)) = 0) Then dlg_title_par = "A T T E N T I O N"
            If (defa_button_par = 0) Then defa_button_par = vbDefaultButton2

            ret_val = MsgBox(dlg_det_par, vbQuestion + vbYesNo + defa_button_par, dlg_title_par)


        ElseIf (dlg_no_par = EnumAppDlgType.Quit) Then

            If (Len(Trim(dlg_det_par)) = 0) Then dlg_det_par = "Want To Quit ?"
            If (Len(Trim(dlg_title_par)) = 0) Then dlg_title_par = "A T T E N T I O N"
            If (defa_button_par = 0) Then defa_button_par = vbDefaultButton2

            ret_val = MsgBox(dlg_det_par, vbQuestion + vbYesNo + defa_button_par, dlg_title_par)


        ElseIf (dlg_no_par = EnumAppDlgType.Confirmation) Then

            If (Len(Trim(dlg_det_par)) = 0) Then dlg_det_par = "Are You Sure ?"
            If (Len(Trim(dlg_title_par)) = 0) Then dlg_title_par = "A T T E N T I O N"
            If (defa_button_par = 0) Then defa_button_par = vbDefaultButton2

            ret_val = MsgBox(dlg_det_par, vbQuestion + vbYesNo + defa_button_par, dlg_title_par)


        ElseIf (dlg_no_par = EnumAppDlgType.SaveDetails) Then

            If (Len(Trim(dlg_det_par)) = 0) Then dlg_det_par = "Save Details ?"
            If (Len(Trim(dlg_title_par)) = 0) Then dlg_title_par = "A T T E N T I O N"
            If (defa_button_par = 0) Then defa_button_par = vbDefaultButton1

            ret_val = MsgBox(dlg_det_par, vbQuestion + vbYesNo + defa_button_par, dlg_title_par)


        ElseIf (dlg_no_par = EnumAppDlgType.SaveChanges) Then

            If (Len(Trim(dlg_det_par)) = 0) Then dlg_det_par = "Save Changes ?"
            If (Len(Trim(dlg_title_par)) = 0) Then dlg_title_par = "A T T E N T I O N"
            If (defa_button_par = 0) Then defa_button_par = vbDefaultButton1

            ret_val = MsgBox(dlg_det_par, vbQuestion + vbYesNo + defa_button_par, dlg_title_par)

        End If


        Cursor.Current = mpointer_lcl

        DispAppDlg = ret_val

    End Function

    Public Sub GetErrorDet(ByVal Reset_par As Boolean)

        Dim err_no, err_details

        If (Not Reset_par) Then
            err_no = Err.Number
            err_details = Err.Description
        Else
            err_no = 0
            err_details = ""
        End If

    End Sub

End Module

Public Module Threads_Related

    Public Delegate Sub SubWithoutParDelegateType()
    Public Delegate Sub SubWithParDelegateType(ByVal Arg_par As Object)

    Public Delegate Function FuncWithoutParDelegateType() As Object
    Public Delegate Function FuncWithParDelegateType(ByVal Arg_par As Object) As Object

    Public Class clsThread

        Implements ICloneable

        Public _Thread As Thread

        Public _StartData As Object

        Public _DefaultBackgroundStatus As Boolean

        Sub New(Optional ByVal DefaultBackgroundStatus_par As Boolean = True)

            _Thread = Nothing

            _StartData = Nothing

            _DefaultBackgroundStatus = DefaultBackgroundStatus_par

        End Sub

        Protected Overrides Sub Finalize()

        End Sub

        Public ReadOnly Property IfValid() As Boolean

            Get
                IfValid = Not IsNothing(_Thread)
            End Get

        End Property

        Public ReadOnly Property IfValidStartData() As Boolean

            Get
                IfValidStartData = Not IsNothing(_StartData)
            End Get

        End Property

        Public Sub Create(ByRef ThreadStart_par As Thread,
                          Optional ByVal StartData_par As Object = Nothing)

            '            Try

            If (IfValid) Then
                If (_Thread.IsAlive) Then GoTo end_sub
            End If

            _Thread = ThreadStart_par
            With _Thread
                .IsBackground = _DefaultBackgroundStatus
            End With

            _StartData = StartData_par

            '            Catch

            '            End Try
end_sub:

        End Sub

        Public Sub Create(ByRef ThreadStart_par As ThreadStart)

            '            Try

            If (IfValid) Then
                If (_Thread.IsAlive) Then GoTo end_sub
            End If

            _Thread = New Thread(ThreadStart_par)
            With _Thread
                .IsBackground = _DefaultBackgroundStatus
            End With

            _StartData = Nothing

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub Create(ByRef ThreadStart_par As ParameterizedThreadStart,
                          ByRef StartData_par As Object)

            '            Try

            If (IfValid) Then
                If (_Thread.IsAlive) Then GoTo end_sub
            End If

            _Thread = New Thread(ThreadStart_par)
            With _Thread
                .IsBackground = _DefaultBackgroundStatus
            End With

            _StartData = StartData_par

            '            Catch

            '            End Try

end_sub:
        End Sub

        Public Function _Start() As Object

            Dim ret_val = False

            '            Try

            If (Not IfValid) Then
                GoTo end_sub
            ElseIf (_Thread.IsAlive) Then
                GoTo end_sub
            End If

            If (Not IfValidStartData) Then
                _Thread.Start()
            Else
                _Thread.Start(_StartData)
            End If

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            _Start = ret_val

        End Function

        Public Function _Stop() As Object

            Dim ret_val = False

            '            Try

            If (Not IfValid) Then GoTo end_sub

            With _Thread
                If (.IsAlive) Then .Abort()
            End With

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            _Stop = ret_val

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Class clsThreadQueue

        Implements ICloneable

        Dim _Queue As Queue(Of Thread)
        Dim _lstThreadStartData As List(Of Object)

        Dim _Count As Integer

        Public _DefaultBackgroundStatus As Boolean

        Public _CurrentThread As Thread
        Public _CurrentThreadStartData As Object

        Sub New(Optional ByVal DefaultBackgroundStatus_par As Boolean = True)

            _DefaultBackgroundStatus = DefaultBackgroundStatus_par

            _Queue = New Queue(Of Thread)
            _lstThreadStartData = New List(Of Object)

            _CurrentThread = Nothing
            _CurrentThreadStartData = Nothing

        End Sub

        Protected Overrides Sub Finalize()

            Clear()

        End Sub

        Public ReadOnly Property IfValidCurrentThread() As Boolean

            Get
                IfValidCurrentThread = Not IsNothing(_CurrentThread)
            End Get

        End Property

        Public ReadOnly Property IfValidCurrentThreadStartData() As Boolean

            Get
                IfValidCurrentThreadStartData = Not IsNothing(_CurrentThreadStartData)
            End Get

        End Property

        Public ReadOnly Property Count() As Integer

            Get
                Count = _Queue.Count
            End Get

        End Property

        Public Function _Start() As Object

            Dim ret_val = False

            '            Try

            _CurrentThread = Nothing
            _CurrentThreadStartData = Nothing

            If (Count = 0) Then GoTo end_sub

            With _lstThreadStartData

                _CurrentThread = _Queue(0)
                _CurrentThreadStartData = .Item(0)

                If (Not IfValidCurrentThreadStartData) Then
                    _CurrentThread.Start()
                Else
                    _CurrentThread.Start(_CurrentThreadStartData)
                End If

            End With

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            _Start = ret_val

        End Function

        Public Function _Stop() As Object

            Dim ret_val = False

            '            Try

            If (Not IfValidCurrentThread) Then GoTo end_sub

            With _CurrentThread
                If (.IsAlive) Then .Abort()
            End With

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            _Stop = ret_val

        End Function

        Public Sub Enqueue(ByVal Thread_par As System.Threading.Thread,
                           Optional ByVal StartData_par As Object = Nothing)

            '            Try

            With Thread_par
                .IsBackground = _DefaultBackgroundStatus
            End With

            _Queue.Enqueue(Thread_par)
            _lstThreadStartData.Add(StartData_par)

            '            Catch

            '            End Try

end_sub:
        End Sub

        Public Sub Dequeue()

            '            Try

            _Queue.Dequeue()

            If Count > 0 Then
                _lstThreadStartData.RemoveAt(0)
                _Start()
            Else
                _CurrentThread = Nothing
                Clear()
            End If

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub Clear()

            '            Try

            _Stop()

            _CurrentThread = Nothing
            _CurrentThreadStartData = Nothing

            _Queue.Clear()
            _lstThreadStartData.Clear()

            '            Catch

            '            End Try

        End Sub

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Function RunDelegate(ByVal Ctrl_par As Object,
                                ByVal Delegate_par As SubWithoutParDelegateType,
                                Optional ByVal AsyncRunStat_par As Object = True) As Object

        Dim ret_val As Object = False

        '        Try

        With Ctrl_par

            If (.InvokeRequired) Then

                If (AsyncRunStat_par) Then

                    Dim AsyncRes_lcl As IAsyncResult
                    AsyncRes_lcl = .BeginInvoke(Delegate_par)

                    .EndInvoke(AsyncRes_lcl)

                Else

                    .Invoke(Delegate_par)

                End If

                ret_val = True

            End If

        End With

        '        Catch

        '        End Try

end_sub:

        RunDelegate = ret_val

    End Function

    Public Function RunDelegate(ByVal Ctrl_par As Object,
                                ByVal Delegate_par As SubWithParDelegateType,
                                ByRef Arg_par As Object,
                                Optional ByVal AsyncRunStat_par As Object = True) As Object

        Dim ret_val As Object = False

        '        Try

        With Ctrl_par

            If (.InvokeRequired) Then

                If (AsyncRunStat_par) Then

                    Dim AsyncRes_lcl As IAsyncResult
                    AsyncRes_lcl = .BeginInvoke(Delegate_par, Arg_par)

                    .EndInvoke(AsyncRes_lcl)

                Else

                    .Invoke(Delegate_par, Arg_par)

                End If

                ret_val = True

            End If

        End With

        '        Catch

        '        End Try

end_sub:

        RunDelegate = ret_val

    End Function

    Public Sub RunDelegate(ByVal Delegate_par As SubWithoutParDelegateType,
                           Optional ByRef CallbackMethodArg_par As Object = Nothing,
                           Optional ByVal AsyncRunStat_par As Object = True)

        Dim AsyncRes_lcl As IAsyncResult = Nothing

        '        Try

        With Delegate_par

            If (AsyncRunStat_par) Then
                AsyncRes_lcl = .BeginInvoke(AddressOf CallbackSubWithoutParDelegate, CallbackMethodArg_par)
            Else
                .Invoke()
            End If

        End With

        '        Catch

        '        End Try

end_sub:

    End Sub

    Public Sub CallbackSubWithoutParDelegate(ByVal Ar As IAsyncResult)

        ' Retrieve the delegate.

        Dim result As AsyncResult = CType(Ar, AsyncResult)
        Dim caller As SubWithoutParDelegateType = CType(result.AsyncDelegate, SubWithoutParDelegateType)

        ' Retrieve the format string that was passed as state information.

        Dim formatString As String = CType(Ar.AsyncState, String)

        ' Call EndInvoke to retrieve the results.

        caller.EndInvoke(Ar)

    End Sub

    Public Sub RunDelegate(ByVal Delegate_par As SubWithParDelegateType,
                           ByRef Arg_par As Object,
                           Optional ByRef CallbackMethodArg_par As Object = Nothing,
                           Optional ByVal AsyncRunStat_par As Object = True)

        Dim AsyncRes_lcl As IAsyncResult = Nothing

        '        Try

        With Delegate_par

            If (AsyncRunStat_par) Then
                AsyncRes_lcl = .BeginInvoke(Arg_par, AddressOf CallbackSubWithParDelegate, CallbackMethodArg_par)
            Else
                .Invoke(Arg_par)
            End If

        End With

        '        Catch

        '        End Try

end_sub:

    End Sub

    Public Sub CallbackSubWithParDelegate(ByVal Ar As IAsyncResult)

        ' Retrieve the delegate.

        Dim result As AsyncResult = CType(Ar, AsyncResult)
        Dim caller As SubWithParDelegateType = CType(result.AsyncDelegate, SubWithParDelegateType)

        ' Retrieve the format string that was passed as state information.

        Dim formatString As String = CType(Ar.AsyncState, String)

        ' Call EndInvoke to retrieve the results.

        caller.EndInvoke(Ar)

    End Sub

    Public Function RunDelegate(ByVal Delegate_par As FuncWithoutParDelegateType,
                                Optional ByRef CallbackMethodArg_par As Object = Nothing,
                                Optional ByVal AsyncRunStat_par As Object = True) As Object

        Dim ret_val As Object = Nothing

        Dim AsyncRes_lcl As IAsyncResult = Nothing

        '        Try

        With Delegate_par

            If (AsyncRunStat_par) Then
                '                AsyncRes_lcl = .BeginInvoke(AddressOf CallbackFuncWithoutParDelegate, CallbackMethodArg_par)
                AsyncRes_lcl = .BeginInvoke(Nothing, Nothing)
                ret_val = .EndInvoke(AsyncRes_lcl)
            Else
                ret_val = .Invoke()
            End If

        End With

        '        Catch

        '        End Try

end_sub:

        RunDelegate = ret_val

    End Function

    Public Sub CallbackFuncWithoutParDelegate(ByVal Ar As IAsyncResult)

        ' Retrieve the delegate.

        Dim result As AsyncResult = CType(Ar, AsyncResult)
        Dim caller As FuncWithoutParDelegateType = CType(result.AsyncDelegate, FuncWithoutParDelegateType)

        ' Retrieve the format string that was passed as state information.

        Dim formatString As String = CType(Ar.AsyncState, String)

        ' Call EndInvoke to retrieve the results.

        Dim returnValue As String = caller.EndInvoke(Ar)

        ' Use the format string to format the output message.
        '        Console.WriteLine(formatString, threadId, returnValue)

    End Sub

    Public Function RunDelegate(ByVal Delegate_par As FuncWithParDelegateType,
                                ByRef Arg_par As Object,
                                Optional ByRef CallbackMethodArg_par As Object = Nothing,
                                Optional ByVal AsyncRunStat_par As Object = True)

        Dim ret_val As Object = Nothing

        Dim AsyncRes_lcl As IAsyncResult = Nothing

        '        Try

        With Delegate_par

            If (AsyncRunStat_par) Then
                '                AsyncRes_lcl = .BeginInvoke(Arg_par, AddressOf CallbackFuncWithParDelegate, CallbackMethodArg_par)
                AsyncRes_lcl = .BeginInvoke(Arg_par, Nothing, Nothing)
                ret_val = .EndInvoke(AsyncRes_lcl)
            Else
                ret_val = .Invoke(Arg_par)
            End If

        End With

        '        Catch

        '        End Try

end_sub:

        RunDelegate = ret_val

    End Function

    Public Sub CallbackFuncWithParDelegate(ByVal Ar As IAsyncResult)

        ' Retrieve the delegate.

        Dim result As AsyncResult = CType(Ar, AsyncResult)
        Dim caller As FuncWithParDelegateType = CType(result.AsyncDelegate, FuncWithParDelegateType)

        ' Retrieve the format string that was passed as state information.

        Dim formatString As String = CType(Ar.AsyncState, String)

        ' Call EndInvoke to retrieve the results.

        Dim returnValue As Object = caller.EndInvoke(Ar)

        ' Use the format string to format the output message.
        '        Console.WriteLine(formatString, threadId, returnValue)

    End Sub

    Public Function RunThread(ByVal ThreadStart_par As ThreadStart,
                              Optional ByVal BackgroundStat_par As Object = True) As Object

        Dim ret_val As Object = Nothing

        '        Try

        Dim Thread_lcl As Thread = New Thread(ThreadStart_par)

        With Thread_lcl

            .IsBackground = BackgroundStat_par

            .Start()

            ret_val = .ManagedThreadId

        End With

        '        Catch

        '        End Try

end_sub:

        RunThread = ret_val

    End Function

    Public Function RunThread(ByVal ThreadStart_par As ParameterizedThreadStart,
                              ByRef ThreadStartData_par As Object,
                              Optional ByVal BackgroundStat_par As Object = True) As Object

        Dim ret_val As Object = Nothing

        '        Try

        Dim Thread_lcl As Thread = New Thread(ThreadStart_par)

        With Thread_lcl

            .IsBackground = BackgroundStat_par

            .Start(ThreadStartData_par)

            ret_val = .ManagedThreadId

        End With

        '        Catch

        '        End Try

end_sub:

        RunThread = ret_val

    End Function

End Module

Public Module XML_Related

    Public Const XmlTag = "Det"
    Public Const XmlTagMainStat = "MainStat"
    Public Const XmlTagImpStat = "ImpStat"
    Public Const XmlTagNameVal = "NameVal"
    Public Const XmlTagIdVal = "IdVal"
    Public Const XmlTagDataVal = "DataVal"
    Public Const XmlTagParentNameVal = "ParentNameVal"
    Public Const XmlTagParentIdVal = "ParentIdVal"
    Public Const XmlTagParentDataVal = "ParentDataVal"
    Public Const XmlTagDefaDataVal = "DefaDataVal"
    Public Const XmlTagDataType = "DataType"
    Public Const XmlTagDataPfx_1 = "DataPfx_1"
    Public Const XmlTagDataPfx_2 = "DataPfx_2"
    Public Const XmlTagDataSfx_1 = "DataSfx_1"
    Public Const XmlTagDataSfx_2 = "DataSfx_2"

    <Serializable()>
    Public Class clsXmlTag

        Implements ICloneable

        Dim main_stat As Boolean
        Dim imp_stat As Boolean

        Dim parent_name_val As String ' Used when element has attributes
        Dim parent_id_val As String
        Dim parent_data_val As Object

        Dim name_val As String
        Dim id_val As String
        Dim data_val As Object

        Dim defa_data_val As Object
        Dim data_type As EnumDataType

        Public DataPfx_1 As Object
        Public DataPfx_2 As Object

        Public DataSfx_1 As Object
        Public DataSfx_2 As Object

        Public Sub New()

            main_stat = False
            imp_stat = False

            parent_name_val = ""
            parent_id_val = ""
            parent_data_val = ""

            name_val = ""
            id_val = ""
            data_val = ""

            defa_data_val = ""
            data_type = EnumDataType.Text

            DataPfx_1 = ""
            DataPfx_2 = ""

            DataSfx_1 = ""
            DataSfx_2 = ""

        End Sub

        Public Sub New(ByVal name_par As String, ByVal id_par As String, ByVal data_par As String,
                       Optional ByVal dafa_data_par As String = "")

            name_val = name_par
            id_val = id_par
            data_val = data_par
            defa_data_val = dafa_data_par

        End Sub

        Public Property MainStat() As Boolean

            Get
                MainStat = main_stat
            End Get

            Set(ByVal Value As Boolean)
                main_stat = Value
            End Set

        End Property

        Public Property ImpStat() As Boolean

            Get
                ImpStat = imp_stat
            End Get

            Set(ByVal Value As Boolean)
                imp_stat = Value
            End Set

        End Property

        Public Property ParentName() As String

            Get
                ParentName = parent_name_val
            End Get

            Set(ByVal Value As String)
                parent_name_val = Value
            End Set

        End Property

        Public Property ParentId() As String

            Get
                ParentId = parent_id_val
            End Get

            Set(ByVal Value As String)
                parent_id_val = Value
            End Set

        End Property

        Public Property ParentData() As Object

            Get
                If (DataType = EnumDataType.Dt) Then
                    ParentData = parent_data_val 'Mid(Parent_data_val, 7, 2) & "/" & Mid(Parent_data_val, 5, 2) & "/" & Mid(Parent_data_val, 1, 4)
                ElseIf (DataType = EnumDataType.Amount) Then
                    ParentData = Format(parent_data_val, ".#0")
                ElseIf (DataType = EnumDataType.Weight) Then
                    ParentData = Format(parent_data_val, ".##0")
                ElseIf (DataType = EnumDataType.Volume) Then
                    ParentData = Format(parent_data_val, ".##0")
                Else
                    ParentData = parent_data_val
                End If

                If (InStr(ParentData, DataPfx_1 & DataPfx_2, CompareMethod.Text) = 0) Then
                    ParentData = DataPfx_1 & DataPfx_2 & Data
                End If

            End Get


            Set(ByVal Value As Object)

                Value = Replace(Value, DataPfx_1 & DataPfx_2, "", 1, -1, CompareMethod.Text)

                If (DataType = EnumDataType.Dt) Then

                    '                    Parent_data_val = Format(CDate(Value).Year, "0###") & _
                    '                               Format(CDate(Value).Month, "0#") & _
                    '                    Format(CDate(Value).Day, "0#")

                    parent_data_val = Format(CDate(Value).Day, "0#") & "-" &
                                      CDate(Value).ToString("MMM") & "-" &
                                      Format(CDate(Value).Year, "0###")

                ElseIf (DataType = EnumDataType.Amount Or DataType = EnumDataType.Weight Or
                        DataType = EnumDataType.Volume) Then

                    parent_data_val = Val(Value)

                Else

                    parent_data_val = Value

                End If

            End Set

        End Property

        Public Property Name() As String

            Get
                Name = name_val
            End Get

            Set(ByVal Value As String)
                name_val = Value
            End Set

        End Property

        Public Property Id() As String

            Get
                Id = id_val
            End Get

            Set(ByVal Value As String)
                id_val = Value
            End Set

        End Property

        Public Property Data() As Object

            Get
                If (DataType = EnumDataType.Dt) Then
                    Data = data_val 'Mid(data_val, 7, 2) & "/" & Mid(data_val, 5, 2) & "/" & Mid(data_val, 1, 4)
                ElseIf (DataType = EnumDataType.Amount) Then
                    Data = Format(data_val, ".#0")
                ElseIf (DataType = EnumDataType.Weight) Then
                    Data = Format(data_val, ".##0")
                ElseIf (DataType = EnumDataType.Volume) Then
                    Data = Format(data_val, ".##0")
                Else
                    Data = data_val
                End If

                If (InStr(Data, DataPfx_1 & DataPfx_2, CompareMethod.Text) = 0) Then
                    Data = DataPfx_1 & DataPfx_2 & Data
                End If

            End Get


            Set(ByVal Value As Object)

                Value = Replace(Value, DataPfx_1 & DataPfx_2, "", 1, -1, CompareMethod.Text)

                If (DataType = EnumDataType.Dt) Then

                    '                    data_val = Format(CDate(Value).Year, "0###") & _
                    '                               Format(CDate(Value).Month, "0#") & _
                    '                    Format(CDate(Value).Day, "0#")

                    data_val = Format(CDate(Value).Day, "0#") & "-" &
                               CDate(Value).ToString("MMM") & "-" &
                               Format(CDate(Value).Year, "0###")

                ElseIf (DataType = EnumDataType.Amount Or DataType = EnumDataType.Weight Or
                        DataType = EnumDataType.Volume) Then

                    data_val = Val(Value)

                Else

                    data_val = Value

                End If

            End Set

        End Property

        Public Property DefaData() As Object

            Get
                DefaData = defa_data_val
            End Get

            Set(ByVal Value As Object)
                defa_data_val = Value
            End Set

        End Property

        Public Property DataType() As EnumDataType

            Get
                DataType = data_type
            End Get

            Set(ByVal Value As EnumDataType)
                data_type = Value
            End Set

        End Property

        Public Sub CopyData(ByVal Data_par As clsXmlTag)

            With Data_par

                data_val = .data_val

                DataPfx_1 = .DataPfx_1
                DataPfx_2 = .DataPfx_2

                DataSfx_1 = .DataSfx_1
                DataSfx_2 = .DataSfx_2

            End With

        End Sub

        Public Function CloneOld() As clsXmlTag

            Dim temp = DirectCast(Me.MemberwiseClone(), clsXmlTag)

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

    Public Sub ResetXmlTag(ByVal LstTag_par As List(Of clsXmlTag))

        Try

            With LstTag_par

                For tag_no = 0 To .Count - 1
                    ResetXmlTag(.Item(tag_no))
                Next tag_no

            End With

        Catch

        End Try

end_sub:

    End Sub

    Public Sub ResetXmlTag(ByRef Tag_par As clsXmlTag)

        Try

            With Tag_par

                .Data = ""

            End With

        Catch

        End Try

end_sub:

    End Sub

    Public Function GetXmlTag(ByVal LstTag_par As List(Of clsXmlTag),
                              ByVal GetDataType_par As Object,
                              ByVal FindData_par As Object,
                              Optional ByVal FindBy_par As Object = FindByData) As Object

        Dim ret_val As Object = Nothing

        '        Try

        Dim FindIndex_lcl As Object = FindXmlTag(LstTag_par, FindData_par, FindBy_par)

        If (FindIndex_lcl >= 0) Then

            With LstTag_par.Item(FindIndex_lcl)

                If (GetDataType_par = XmlTag) Then
                    ret_val = LstTag_par.Item(FindIndex_lcl)
                ElseIf (GetDataType_par = XmlTagDataVal) Then
                    ret_val = ValidateData(.Data, .DataType)
                ElseIf (GetDataType_par = XmlTagIdVal) Then
                    ret_val = .Id
                ElseIf (GetDataType_par = XmlTagNameVal) Then
                    ret_val = .Name
                ElseIf (GetDataType_par = XmlTagParentDataVal) Then
                    ret_val = ValidateData(.ParentData, .DataType)
                ElseIf (GetDataType_par = XmlTagParentIdVal) Then
                    ret_val = .ParentId
                ElseIf (GetDataType_par = XmlTagParentNameVal) Then
                    ret_val = .ParentName
                End If

            End With

        End If

        '        Catch

        '        End Try

end_sub:

        GetXmlTag = ret_val

    End Function

    Public Function FindXmlTag(ByVal LstTag_par As List(Of clsXmlTag), ByVal FindData_par As Object,
                               Optional ByVal FindBy_par As Object = FindByData) As Object

        Dim ret_val As Object = -1

        '        Try

        With LstTag_par

            For tag_no = 0 To .Count - 1

                If (CheckXmlTag(.Item(tag_no), FindData_par, FindBy_par)) Then

                    ret_val = tag_no

                    Exit For

                End If

            Next tag_no

        End With

        '        Catch

        '        End Try

end_sub:

        FindXmlTag = ret_val

    End Function

    Public Function CheckXmlTag(ByVal Tag_par As clsXmlTag, ByVal FindData_par As Object,
                                Optional ByVal FindBy_par As Object = FindByData) As Object

        Dim ret_val As Object = False

        '        Try

        With Tag_par

            If (FindBy_par = FindByData) Then
                ret_val = (UCase(.Data) = UCase(FindData_par))
            ElseIf (FindBy_par = FindById) Then
                ret_val = (UCase(.Id) = UCase(FindData_par))
            ElseIf (FindBy_par = FindByName) Then
                ret_val = (UCase(.Name) = UCase(FindData_par))
            ElseIf (FindBy_par = FindByParentData) Then
                ret_val = (UCase(.ParentData) = UCase(FindData_par))
            ElseIf (FindBy_par = FindByParentId) Then
                ret_val = (UCase(.ParentId) = UCase(FindData_par))
            ElseIf (FindBy_par = FindByParentName) Then
                ret_val = (UCase(.ParentName) = UCase(FindData_par))
            End If

        End With

        '        Catch

        '        End Try

end_sub:

        CheckXmlTag = ret_val

    End Function

    Public Function PrepXmlFromText(ByVal Xml_par As Object,
                                    ByVal LstTags_par As List(Of clsXmlTag)) As Object

        Dim ret_val As Object = ""

        '        Try

        With LstTags_par

            For tag_no = 0 To .Count - 1
                Xml_par = PrepXmlFromText(Xml_par, .Item(tag_no))
            Next tag_no

        End With

        ret_val = Xml_par

        '        Catch

        '        End Try


        frm_master_det.TextBox1.Text = ret_val

        PrepXmlFromText = ret_val

    End Function

    Public Function PrepXmlFromText(ByVal Xml_par As Object, ByVal Tag_par As clsXmlTag) As Object

        Dim ret_val As Object = ""

        '        Try

        Dim tag_val As Object
        Dim str_lcl(2) As Object


        With Tag_par

            tag_val = .Data

            If (Len(Trim(tag_val)) = 0) Then
                tag_val = .DefaData
            End If

            '            If (UCase(.Name) = UCase("tagBasicBuyerAddr1")) Then
            '            MsgBox("," & tag_val & ",")
            '            End If


            If (Len(Trim(tag_val)) = 0) Then

                str_lcl(0) = UCase("<" & .Id & ">" & .Name & "</" & .Id & ">")
                str_lcl(1) = UCase("<" & .Id & "/>")

                Xml_par = Replace(Xml_par, str_lcl(0), str_lcl(1), 1, -1, CompareMethod.Text)

                '                If (UCase(.Name) = UCase("tagBasicBuyerAddr1")) Then
                '                MsgBox(Xml_par, , str_lcl(0) & " - " & InStr(1, Xml_par, str_lcl(0), CompareMethod.Text))
                '            End If

            Else

                tag_val = tag_val & .DataSfx_1 & .DataSfx_2
                tag_val = Replace(tag_val, "&", "&amp;")

                If (InStr(1, Xml_par, .Name, CompareMethod.Text) > 0) Then

                    Xml_par = Replace(Xml_par, .Name, tag_val, 1, -1, CompareMethod.Text)

                Else

                    str_lcl(0) = UCase("<" & .Id & "/>")
                    str_lcl(1) = UCase("<" & .Id & ">" & tag_val & "</" & .Id & ">")

                    Xml_par = Replace(Xml_par, str_lcl(0), str_lcl(1), 1, -1, CompareMethod.Text)

                End If

            End If

            ret_val = Xml_par

        End With

        '        Catch

        '        End Try


        PrepXmlFromText = ret_val

    End Function

    Public Function SplitXmlText(ByVal XmlText_par As Object,
                                 ByVal SplitByTagName_par As Object) As List(Of Object)

        Dim lstRet_lcl As New List(Of Object)

        '        Try

        Dim lstSplitByTag_lcl As New List(Of Object)

        With lstSplitByTag_lcl

            If (IsArray(SplitByTagName_par)) Then
                .AddRange(SplitByTagName_par)
            Else
                .Add(SplitByTagName_par)
            End If

            If (.Count = 0) Then
                GoTo end_sub
            ElseIf (.Count = 1) Then
                .Add(.Item(0))
            End If

            .Item(0) = "<" & .Item(0)
            .Item(1) = "</" & .Item(1) & ">"

        End With


        Dim aLine, aParagraph As String
        Dim strReader As New StringReader(XmlText_par)

        aLine = ""
        aParagraph = ""

        While True

            aLine = strReader.ReadLine()

            If aLine Is Nothing Then Exit While

            If (Len(aLine) > 0) Then aLine &= vbCrLf
            aParagraph = aParagraph & aLine

            If (InStr(aLine, lstSplitByTag_lcl.Item(0), CompareMethod.Text) > 0) Then
                aParagraph = aLine
            ElseIf (InStr(aLine, lstSplitByTag_lcl.Item(1), CompareMethod.Text) > 0) Then
                lstRet_lcl.Add(aParagraph)
                aParagraph = ""
            End If

        End While


        If (Len(aParagraph) > 0 And
            InStr(aLine, lstSplitByTag_lcl.Item(0), CompareMethod.Text) > 0 And
            InStr(aLine, lstSplitByTag_lcl.Item(1), CompareMethod.Text) > 0) Then

            lstRet_lcl.Add(aParagraph)

        End If

        '        Catch

        '        End Try

end_sub:

        SplitXmlText = lstRet_lcl

    End Function

    Public Function PostXml(ByVal XmlToPostXml_par As Object,
                            ByVal Address_par As String, ByVal Port_par As String) As Object

        Dim ret_val As Object = ""

        '        Try

        Dim xdoc As New System.Xml.XmlDocument

        xdoc.LoadXml(XmlToPostXml_par)

        frm_master_det.TextBox1.Text = XmlToPostXml_par
h:
        Dim bdata() As Byte = System.Text.Encoding.ASCII.GetBytes(xdoc.OuterXml)

        Dim bresp() As Byte

        Dim wc As New System.Net.WebClient
        wc.Headers.Add("Content-Type", "text/xml")

        Address_par = Trim(Address_par)
        If (Len(Address_par) = 0) Then
            Address_par = "http://localhost:"
        End If

        Port_par = Trim(Port_par)

        bresp = wc.UploadData(Address_par & Port_par, bdata)

        ret_val = System.Text.Encoding.ASCII.GetString(bresp)

        '        Catch

        '        End Try


        PostXml = ret_val

    End Function

    Public Function ReadXmlFile(ByVal XmlFileName_par As Object,
                                ByVal ReadByTagName_par As Object) As List(Of Object)

        Dim lstRet_lcl As New List(Of Object)

        '        Try

        Dim StrmRdr As StreamReader = New StreamReader(XmlFileName_par.ToString)


        Dim lstParseByTag_lcl As New List(Of Object)

        With lstParseByTag_lcl

            If (IsArray(ReadByTagName_par)) Then
                .AddRange(ReadByTagName_par)
            Else
                .Add(ReadByTagName_par)
            End If

            If (.Count = 0) Then
                GoTo end_sub
            ElseIf (.Count = 1) Then
                .Add(.Item(0))
            End If

            .Item(0) = "<" & .Item(0)
            .Item(1) = "</" & .Item(1) & ">"

        End With


        Dim aLine, aParagraph As Object

        aLine = ""
        aParagraph = ""

        Do While StrmRdr.Peek() >= 0

            aLine = StrmRdr.ReadLine()

            If (Len(aLine) > 0) Then aLine &= vbCrLf
            aParagraph = aParagraph & aLine

            If (InStr(aLine, lstParseByTag_lcl.Item(0), CompareMethod.Text) > 0) Then

                aParagraph = aLine

            ElseIf (InStr(aLine, lstParseByTag_lcl.Item(1), CompareMethod.Text) > 0) Then

                With lstRet_lcl
                    .Add(aParagraph)
                End With

                aParagraph = ""

            End If

        Loop


        If (Len(aParagraph) > 0 And
            InStr(aLine, lstParseByTag_lcl.Item(0), CompareMethod.Text) > 0 And
            InStr(aLine, lstParseByTag_lcl.Item(1), CompareMethod.Text) > 0) Then

            With lstRet_lcl
                .Add(aParagraph)
            End With

        End If


        StrmRdr.Close()

        '        Catch

        '        End Try

end_sub:

        ReadXmlFile = lstRet_lcl

    End Function

    Public Function ParseXmlFile(ByVal XmlFileName_par As Object, ByVal ParseByTagName_par As Object,
                                 Optional ByVal LstDataToExtract As List(Of clsXmlTag) = Nothing) As List(Of List(Of clsXmlTag))

        Dim lstRet_lcl As New List(Of List(Of clsXmlTag))

        '        Try

        Dim StrmRdr As StreamReader = New StreamReader(XmlFileName_par.ToString)


        Dim lstParseByTag_lcl As New List(Of Object)

        With lstParseByTag_lcl

            If (IsArray(ParseByTagName_par)) Then
                .AddRange(ParseByTagName_par)
            Else
                .Add(ParseByTagName_par)
            End If

            If (.Count = 0) Then
                GoTo end_sub
            ElseIf (.Count = 1) Then
                .Add(.Item(0))
            End If

            .Item(0) = "<" & .Item(0)
            .Item(1) = "</" & .Item(1) & ">"

        End With


        Dim aLine, aParagraph As Object

        aLine = ""
        aParagraph = ""

        Do While StrmRdr.Peek() >= 0

            aLine = StrmRdr.ReadLine()

            ' Ignore undeclared prefixex like 'udf:'
            '            If (InStr(aLine, "UDF", CompareMethod.Text) > 0) Then GoTo next_loop

            If (Len(aLine) > 0) Then aLine &= vbCrLf
            aParagraph = aParagraph & aLine

            If (InStr(aLine, lstParseByTag_lcl.Item(0), CompareMethod.Text) > 0) Then

                aParagraph = aLine

            ElseIf (InStr(aLine, lstParseByTag_lcl.Item(1), CompareMethod.Text) > 0) Then

                With lstRet_lcl
                    .Add(New List(Of clsXmlTag))
                    EnumXmlChildNode(aParagraph, .Item(.Count - 1), LstDataToExtract)
                    If (.Item(.Count - 1).Count = 0) Then .RemoveAt(.Count - 1)
                End With

                aParagraph = ""

            End If

next_loop:

        Loop


        If (Len(aParagraph) > 0 And
            InStr(aLine, lstParseByTag_lcl.Item(0), CompareMethod.Text) > 0 And
            InStr(aLine, lstParseByTag_lcl.Item(1), CompareMethod.Text) > 0) Then

            With lstRet_lcl
                .Add(New List(Of clsXmlTag))
                EnumXmlChildNode(aParagraph, .Item(.Count - 1), LstDataToExtract)
                If (.Item(.Count - 1).Count = 0) Then .RemoveAt(.Count - 1)
            End With

        End If


        StrmRdr.Close()

        '        Catch

        '        End Try

end_sub:

        ParseXmlFile = lstRet_lcl

    End Function

    Public Function ParseXmlText(ByVal XmlText_par As Object, ByVal ParseByTagName_par As Object,
                                 Optional ByVal LstDataToExtract As List(Of clsXmlTag) = Nothing) As List(Of List(Of clsXmlTag))

        Dim lstRet_lcl As New List(Of List(Of clsXmlTag))

        '        Try

        '        XmlText_par = UCase(XmlText_par.ToString)

        Dim XmlDoc As New System.Xml.XmlDocument
        XmlDoc.LoadXml(XmlText_par)

        Dim XmlNode As System.Xml.XmlNodeList
        XmlNode = XmlDoc.GetElementsByTagName(UCase(ParseByTagName_par.ToString))


        For node_no As Integer = 0 To XmlNode.Count - 1

            With lstRet_lcl

                .Add(New List(Of clsXmlTag))

                EnumXmlChildNode(XmlNode(node_no), .Item(.Count - 1), LstDataToExtract)

                If (.Item(.Count - 1).Count = 0) Then
                    .RemoveAt(.Count - 1)
                End If

            End With

        Next node_no

        '        Catch

        '        End Try


        ParseXmlText = lstRet_lcl

    End Function

    Function EnumXmlChildNode(ByVal XmlNode_par As System.Xml.XmlNode,
                              ByRef LstData_par As List(Of clsXmlTag),
                              ByVal LstDataToExtract As List(Of clsXmlTag)) As Object

        Dim ret_val As Object = ""
        Dim proc_stat As Object = False

        '        Try

        Select Case XmlNode_par.NodeType

            Case Xml.XmlNodeType.Element

                Dim _xml As System.Xml.XmlElement = XmlNode_par

                If _xml.InnerText = _xml.InnerXml Then

                    proc_stat = False

                    If (IsNothing(LstDataToExtract)) Then
                        proc_stat = True
                    Else
                        proc_stat = FindXmlTag(LstDataToExtract, _xml.Name, FindById)
                    End If


                    If ((IsNothing(LstDataToExtract) And proc_stat) Or
                        (Not IsNothing(LstDataToExtract) And proc_stat >= 0)) Then

                        With LstData_par

                            .Add(New clsXmlTag)

                            If (Not IsNothing(LstDataToExtract) And proc_stat >= 0) Then
                                .Item(.Count - 1) = LstDataToExtract.Item(proc_stat).Clone
                            End If

                            With .Item(.Count - 1)
                                .Id = _xml.Name
                                .Data = Trim(_xml.InnerText)
                            End With

                            ret_val &= _xml.Name & ": " & _xml.InnerText & vbCrLf

                        End With

                    End If

                Else

                    '                    ret_val &= _xml.Name & vbCrLf

                End If

        End Select


        For Each _xml As System.Xml.XmlNode In XmlNode_par.ChildNodes
            ret_val &= EnumXmlChildNode(_xml, LstData_par, LstDataToExtract)
        Next

        '        Catch

        '        End Try


        EnumXmlChildNode = ret_val

    End Function

    Function EnumXmlChildNode(ByVal XmlText_par As Object,
                              ByRef LstData_par As List(Of clsXmlTag),
                              ByVal LstDataToExtract As List(Of clsXmlTag),
                              Optional ByVal NameSpacesStat_par As Boolean = False) As Object

        Dim ret_val As Object = ""
        Dim proc_stat As Object = False

        Dim oXmlEleTag As New clsXmlTag

        '        Try

        '        Dim XmlTxtRdr As XmlTextReader = XmlTextReader.Create(New StringReader(XmlText_par))
        Dim XmlTxtRdr As XmlTextReader = New XmlTextReader(New StringReader(XmlText_par))

        ' Namespaces = False (Ignores element name prefix)
        XmlTxtRdr.Namespaces = NameSpacesStat_par


        Do While (XmlTxtRdr.Read())

            Select Case XmlTxtRdr.NodeType

                Case XmlNodeType.Element 'Beginning of element.

                    With oXmlEleTag
                        .Id = XmlTxtRdr.Name
                        .Data = Nothing
                        .ParentId = Nothing
                        .ParentData = Nothing
                    End With


                    If XmlTxtRdr.HasAttributes Then 'If attributes exist

                        With oXmlEleTag
                            .ParentId = XmlTxtRdr.Name
                            .ParentData = XmlTxtRdr.Value
                        End With


                        While XmlTxtRdr.MoveToNextAttribute()

                            With oXmlEleTag
                                .Id = XmlTxtRdr.Name
                                .Data = XmlTxtRdr.Value
                            End With


                            proc_stat = False

                            If (IsNothing(LstDataToExtract)) Then
                                proc_stat = True
                            Else
                                proc_stat = FindXmlTag(LstDataToExtract, oXmlEleTag.Id, FindById)
                            End If


                            If ((IsNothing(LstDataToExtract) And proc_stat) Or
                                (Not IsNothing(LstDataToExtract) And proc_stat >= 0)) Then

                                With LstData_par

                                    .Add(New clsXmlTag)

                                    If (Not IsNothing(LstDataToExtract) And proc_stat >= 0) Then
                                        .Item(.Count - 1) = LstDataToExtract.Item(proc_stat).Clone
                                    End If

                                    With .Item(.Count - 1)
                                        .Id = oXmlEleTag.Id
                                        .Data = oXmlEleTag.Data
                                        .ParentId = oXmlEleTag.ParentId
                                        .ParentData = oXmlEleTag.ParentData
                                    End With

                                    ret_val &= oXmlEleTag.Id & ": " & oXmlEleTag.Data & vbCrLf

                                End With

                            End If


                            With oXmlEleTag
                                .Id = Nothing
                                .Data = Nothing
                                .ParentId = Nothing
                                .ParentData = Nothing
                            End With

                        End While

                    End If


                Case XmlNodeType.Text 'Text in each element.

                    oXmlEleTag.Data = XmlTxtRdr.Value


                Case XmlNodeType.EndElement 'End of element.

                    If (Not IsNothing(oXmlEleTag.Id)) Then


                        proc_stat = False

                        If (IsNothing(LstDataToExtract)) Then
                            proc_stat = True
                        Else
                            proc_stat = FindXmlTag(LstDataToExtract, oXmlEleTag.Id, FindById)
                        End If


                        If ((IsNothing(LstDataToExtract) And proc_stat) Or
                            (Not IsNothing(LstDataToExtract) And proc_stat >= 0)) Then

                            With LstData_par

                                .Add(New clsXmlTag)

                                If (Not IsNothing(LstDataToExtract) And proc_stat >= 0) Then
                                    .Item(.Count - 1) = LstDataToExtract.Item(proc_stat).Clone
                                End If

                                With .Item(.Count - 1)
                                    .Id = oXmlEleTag.Id
                                    .Data = oXmlEleTag.Data
                                End With

                                ret_val &= oXmlEleTag.Id & ": " & oXmlEleTag.Data & vbCrLf

                            End With

                        End If

                    End If


                    oXmlEleTag.Id = Nothing

            End Select

        Loop

        '        Catch

        '        End Try


        EnumXmlChildNode = ret_val

    End Function

End Module

Public Module List_Related

    Public Function CreateList(ByVal Data_par() As Object) As List(Of Object)

        Dim lst_ret_val As New List(Of Object)

        '        Try

        lst_ret_val.AddRange(Data_par)

        '        Catch

        '        End Try

end_sub:

        CreateList = lst_ret_val

    End Function

    Public Function CreateList(ByVal Data_par As Object, ByVal DelimiterChar_par As Object) As List(Of Object)

        Dim lst_ret_val As New List(Of Object)

        '        Try

        If (Len(DelimiterChar_par) = 0) Then DelimiterChar_par = " "

        lst_ret_val.AddRange(Split(Data_par, DelimiterChar_par))

        '        Catch

        '        End Try

end_sub:

        CreateList = lst_ret_val

    End Function

    Public Function CreateList(ByVal DbConn_par As Object,
                               ByVal SqlDet_par As Object,
                               ByVal TextFieldName_par As Object,
                               ByVal CodeFieldName_par As Object) As List(Of clsGenData)

        Dim lst_ret_val As New List(Of clsGenData)

        '        Try

        SqlDet_par = Trim(SqlDet_par)
        CodeFieldName_par = Trim(CodeFieldName_par)
        TextFieldName_par = Trim(TextFieldName_par)

        If (DbConn_par.State = 0 Or Len(SqlDet_par) = 0 Or Len(CodeFieldName_par) = 0 Or
            Len(TextFieldName_par) = 0) Then GoTo end_sub

        Dim Ds As System.Data.DataSet = CreateDataSet(DbConn_par, SqlDet_par)

        With Ds

            Dim Dr As System.Data.DataRow

            For Each Dr In .Tables(0).Rows

                With lst_ret_val

                    .Add(New clsGenData)

                    '                    With .Item(.Count - 1)
                    '                    .Code = Dr(CodeFieldName_par)
                    '                    .Text = Dr(TextFieldName_par)
                    '                End With

                    CallByName(.Item(.Count - 1), "code", Microsoft.VisualBasic.vbSet, Dr(CodeFieldName_par))
                    CallByName(.Item(.Count - 1), "text", Microsoft.VisualBasic.vbSet, Dr(TextFieldName_par))

                End With

            Next Dr

        End With

        Ds.Dispose()

        '        Catch

        '        End Try

end_sub:

        CreateList = lst_ret_val

    End Function

    Public Function AddListItem(ByRef List_par As List(Of clsDbFieldDet),
                                ByVal Id_par As Object) As Object

        Dim ret_val As Object = False

        '        Try

        Dim mpointer_lcl = Cursor.Current
        Cursor.Current = Cursors.WaitCursor

        List_par.Add(New clsDbFieldDet)

        With List_par(List_par.Count - 1)
            .Name = Id_par
        End With

        ret_val = True

        '        Catch

        '        End Try

end_sub:

        Cursor.Current = mpointer_lcl

        AddListItem = ret_val

    End Function

End Module

Public Module General

    Public Const EncryptCharCode = 16

    Public Const DelimiterChar = " "      ' chr(255)
    Public Const DelimiterChar_2 = "²"    ' chr(253)
    Public Const DelimiterChar_3 = "û"    ' chr(150)
    Public Const DelimiterChar_4 = "ù"    ' chr(151)

    Public Const PropAutoSizeMode = "AutoSizeMode"
    Public Const PropWidth = "Width"
    Public Const PropHeight = "Height"
    Public Const PropAlignment = "Alignment"
    Public Const PropFont = "Font"
    Public Const PropFontName = "FontName"
    Public Const PropFontSize = "FontSize"
    Public Const PropFontStyle = "FontStyle"
    Public Const PropColor = "Color"
    Public Const PropBackColor = "BackColor"
    Public Const PropForeColor = "ForeColor"
    Public Const PropSelColor = "SelColor"
    Public Const PropSelBackColor = "SelBackColor"
    Public Const PropSelForeColor = "SelForeColor"
    Public Const PropDataType = "DataType"
    Public Const PropInputMask = "InputMask"
    Public Const PropDispFormat = "DispFormat"
    Public Const PropCode = "Code"
    Public Const PropText = "Text"
    Public Const PropOthText = "OthText"
    Public Const PropValue = "Value"
    Public Const PropValueForRelCondn = "ValueForRelCondn"
    Public Const PropLocked = "Locked"
    Public Const PropFrozenStat = "FrozenStat"
    Public Const PropDataAddress = "DataAddress"
    Public Const PropImpStat = "ImpStat"
    Public Const PropDispControlType = "DispControlType"
    Public Const PropDispControl = "DispControl"

    Public Const InputMask_hh_mm = "##:##"
    Public Const inputMask_dd_mm_yyyy = "##/##/####"

    Public Const DispFormat_hh_mm = "hh:mm"
    Public Const DispFormat_hh_mm_ampm = "hh:mm AM/PM"
    Public Const DispFormat_hh_mm_ss = "hh:mm:ss"
    Public Const DispFormat_dd_mm_yyyy = "dd/mm/yyyy"
    Public Const DispFormat_dd_mm_yy = "dd/mm/yy"
    Public Const DispFormat_dd_mmm_yyyy = "dd/mmm/yyyy"
    Public Const DispFormat_dd_mmm_yy = "dd/mmm/yy"
    Public Const DispFormat_mm_dd_yyyy = "mm/dd/yyyy"
    Public Const DispFormat_2deci = ".#0"
    Public Const DispFormat_3deci = ".##0"

    Public Const FontRegular = System.Drawing.FontStyle.Regular
    Public Const FontBold = System.Drawing.FontStyle.Bold
    Public Const FontItalic = System.Drawing.FontStyle.Italic
    Public Const FontUnderline = System.Drawing.FontStyle.Underline
    Public Const FontStrikeOut = System.Drawing.FontStyle.Strikeout

    Public Const FindByKey = "FindByKey"
    Public Const FindByItem = "FindByItem"
    Public Const FindById = "FindById"
    Public Const FindByData = "FindByData"
    Public Const FindByName = "FindByName"
    Public Const FindByParentId = "FindByParentId"
    Public Const FindByParentData = "FindByParentData"
    Public Const FindByParentName = "FindByParentName"
    Public Const FindByNo = "FindByNo"
    Public Const FindByRefNo = "FindByRefNo"
    Public Const FindByCode = "FindByCode"
    Public Const FindByOthCode = "FindByOthCode"
    Public Const FindByMainDet = "FindByMainDet"
    Public Const FindByLinkDet = "FindByLinkDet"
    Public Const FindBySql = "FindBySql"
    Public Const FindByNone = "FindByNone"

    Public Enum EnumDataType
        Text
        Amount
        Int
        Real
        Dt
        Bool
        Weight
        Volume
        None
    End Enum

    Public Enum EnumAlignment
        Left
        Right
        Center
    End Enum

    Enum EnumColType
        Header
        Data
        Both
    End Enum

    Public Enum EnumDataOpn
        NewItem
        None
    End Enum

    Enum EnumApplyTo
        CurrentItem
        AllItems
    End Enum

    Public Enum EnumDataPointer
        First
        Last
        None
    End Enum

    Public Enum EnumResult
        Found
        None
    End Enum

    Public Enum EnumDataStat
        Inactive = -1
        Active = 0
        Deleted = 1
        ToDelete = 2
    End Enum

    Enum EnumDispControlType
        ComboBox
        None
    End Enum

    <Serializable()>
    Public Class clsGenData

        Implements ICloneable
        Implements INotifyPropertyChanged

        Dim _Name As Object

        Dim _ImpStat As Boolean
        Dim _FrozenStat As Boolean

        Dim _Code As Object
        Dim _Text As Object
        Dim _OthText As Object

        Dim _OthDet_1 As Object
        Dim _OthDet_2 As Object
        Dim _OthDet_3 As Object

        Dim _DataType As EnumDataType
        Dim _DispFormat As Object

        Dim _DispControlType As Object
        Dim _DispControlData As Object

        Dim _AutoSizeMode As Object
        Dim _Alignment As EnumAlignment

        Dim _FontName As Object
        Dim _FontSize As Object
        Dim _FontStyle As Object

        Dim _BackColor As Object
        Dim _ForeColor As Object

        Dim _SelBackColor As Object
        Dim _SelForeColor As Object

        Dim _WidthCharCnt As Object
        Dim _HeightLineCnt As Object

        Dim _Locked As Boolean

        ' Useful with Excel Workbook, Etc.,

        Dim _DataAddress As Object

        ' Control/Object this class instance is bound to

        Dim _oBindingObj As Object

        Public Event PropertyChanged(ByVal sender As Object,
                                     ByVal e As System.ComponentModel.PropertyChangedEventArgs) _
                                     Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

        Sub New()

            _ImpStat = False

            _DataType = EnumDataType.Text
            _DispFormat = ""
            _Alignment = EnumAlignment.Left

            _FontName = "courier new"
            _FontSize = 10
            _FontStyle = FontRegular

            _BackColor = Color.White
            _ForeColor = Color.Black

            _WidthCharCnt = 10
            _HeightLineCnt = 1

            _Locked = False

            _DataAddress = CreateList({""})

        End Sub

        Public Property Name() As Object

            Get
                Name = _Name
            End Get

            Set(ByVal Value As Object)
                _Name = Value
            End Set

        End Property

        Public Property ImpStat() As Boolean

            Get
                ImpStat = _ImpStat
            End Get

            Set(ByVal Value As Boolean)
                _ImpStat = Value
            End Set

        End Property

        Public Property FrozenStat() As Boolean

            Get
                FrozenStat = _FrozenStat
            End Get

            Set(ByVal Value As Boolean)
                _FrozenStat = Value
            End Set

        End Property

        Public Property Code() As Object

            Get
                Code = _Code
            End Get

            Set(ByVal Value As Object)
                _Code = Value
            End Set

        End Property

        Public Property Text() As Object

            Get
                Text = _Text

                Try
                    If (Len(Trim(_Text)) > 0 And Len(_DispFormat) > 0) Then

                        If (IfNumber(_Text)) Then
                            Text = Format(Val(_Text), _DispFormat)
                        ElseIf (IfDate(_Text)) Then
                            Text = Format(CDate(_Text), _DispFormat)
                        End If

                    End If

                Catch

                End Try

            End Get


            Set(ByVal Value As Object)
                _Text = Value
            End Set

        End Property

        Public Property OthText() As Object

            Get
                OthText = _OthText
            End Get

            Set(ByVal Value As Object)
                _OthText = Value
            End Set

        End Property

        Public Property OthDet_1() As Object

            Get
                OthDet_1 = _OthDet_1
            End Get

            Set(ByVal Value As Object)
                _OthDet_1 = Value
            End Set

        End Property

        Public Property OthDet_2() As Object

            Get
                OthDet_2 = _OthDet_2
            End Get

            Set(ByVal Value As Object)
                _OthDet_2 = Value
            End Set

        End Property

        Public Property OthDet_3() As Object

            Get
                OthDet_3 = _OthDet_3
            End Get

            Set(ByVal Value As Object)
                _OthDet_3 = Value
            End Set

        End Property

        Public Property DataType() As EnumDataType

            Get
                DataType = _DataType
            End Get


            Set(ByVal Value As EnumDataType)

                If (_DataType <> Value) Then

                    If (Value = EnumDataType.Dt) Then
                        DispFormat = DispFormat_dd_mm_yy
                        Alignment = EnumAlignment.Center
                    ElseIf (Value = EnumDataType.Amount) Then
                        DispFormat = DispFormat_2deci
                        Alignment = EnumAlignment.Right
                    ElseIf (Value = EnumDataType.Weight) Then
                        DispFormat = DispFormat_3deci
                        Alignment = EnumAlignment.Right
                    ElseIf (Value = EnumDataType.Volume) Then
                        DispFormat = DispFormat_3deci
                        Alignment = EnumAlignment.Right
                    ElseIf (Value = EnumDataType.Bool) Then
                        Alignment = EnumAlignment.Center
                    Else
                        DispFormat = DispFormat_3deci
                        Alignment = EnumAlignment.Right
                    End If

                End If

                _DataType = Value

            End Set

        End Property

        Public Property DispFormat() As Object

            Get
                DispFormat = _DispFormat
            End Get

            Set(ByVal Value As Object)
                _DispFormat = Value
            End Set

        End Property

        Public Property DispControlType() As Object

            Get
                DispControlType = _DispControlType
            End Get

            Set(ByVal Value As Object)
                _DispControlType = Value
            End Set

        End Property

        Public Property DispControlData() As Object

            Get
                DispControlData = _DispControlData
            End Get

            Set(ByVal Value As Object)
                _DispControlData = Value
            End Set

        End Property

        Public Property AutoSizeMode() As Object

            Get
                AutoSizeMode = _AutoSizeMode
            End Get

            Set(ByVal Value As Object)
                _AutoSizeMode = Value
            End Set

        End Property

        Public Property Alignment() As EnumAlignment

            Get
                Alignment = _Alignment
            End Get

            Set(ByVal Value As EnumAlignment)
                _Alignment = Value
            End Set

        End Property

        Public Property FontName() As Object

            Get
                FontName = _FontName
            End Get


            Set(ByVal Value As Object)

                Dim PrevVal = _FontName

                _FontName = Value

                If (Len(Trim(_FontName)) > 0 And UCase(PrevVal) <> UCase(_FontName)) Then
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FontName"))
                End If

            End Set

        End Property

        Public Property FontSize() As Object

            Get
                FontSize = _FontSize
            End Get


            Set(ByVal Value As Object)

                Dim PrevVal = _FontSize

                _FontSize = Value

                If (_FontSize > 0 And PrevVal <> _FontSize) Then
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FontSize"))
                End If

            End Set

        End Property

        Public Property FontStyle() As Object

            Get
                FontStyle = _FontStyle
            End Get


            Set(ByVal Value As Object)

                Dim PrevVal = _FontStyle

                _FontStyle = Value

                If (PrevVal <> _FontStyle) Then
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("FontStyle"))
                End If

            End Set

        End Property

        Public Property BackColor() As Object

            Get
                BackColor = _BackColor
            End Get


            Set(ByVal Value As Object)

                Dim PrevVal = _BackColor

                _BackColor = Value

                If (PrevVal <> _BackColor) Then
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("BackColor"))
                End If

            End Set

        End Property

        Public Property ForeColor() As Object

            Get
                ForeColor = _ForeColor
            End Get


            Set(ByVal Value As Object)

                Dim PrevVal = _ForeColor

                _ForeColor = Value

                If (PrevVal <> _ForeColor) Then
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("ForeColor"))
                End If

            End Set

        End Property

        Public Property SelBackColor() As Object

            Get
                SelBackColor = _SelBackColor
            End Get

            Set(ByVal Value As Object)
                _SelBackColor = Value
            End Set

        End Property

        Public Property SelForeColor() As Object

            Get
                SelForeColor = _SelForeColor
            End Get

            Set(ByVal Value As Object)
                _SelForeColor = Value
            End Set

        End Property

        Public Property WidthCharsCnt() As Object

            Get
                WidthCharsCnt = _WidthCharCnt
            End Get

            Set(ByVal Value As Object)
                _WidthCharCnt = Value
            End Set

        End Property

        Public Property HeightLinesCnt() As Object

            Get
                HeightLinesCnt = _HeightLineCnt
            End Get

            Set(ByVal Value As Object)
                _HeightLineCnt = Value
            End Set

        End Property

        Public Property Locked() As Boolean

            Get
                Locked = _Locked
            End Get


            Set(ByVal Value As Boolean)

                Dim PrevVal = _Locked

                _Locked = Value

                If (PrevVal <> _Locked) Then
                    RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs("Locked"))
                End If

            End Set

        End Property

        Public Property DataAddress() As Object

            Get
                DataAddress = _DataAddress
            End Get

            Set(ByVal Value As Object)
                _DataAddress = Value
            End Set

        End Property

        Public Property BindingObj() As Object

            Get
                BindingObj = _oBindingObj
            End Get

            Set(ByVal Value As Object)
                _oBindingObj = Value
            End Set

        End Property

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

        Private Sub clsGenData_PropertyChanged(ByVal sender As Object,
          ByVal e As System.ComponentModel.PropertyChangedEventArgs) Handles Me.PropertyChanged

            With DirectCast(sender, clsGenData)

                If (Not IsNothing(.BindingObj)) Then

                End If

            End With

        End Sub

        Protected Overrides Sub Finalize()

            MyBase.Finalize()

        End Sub

    End Class

    Function IfValid(ByVal Obj_par As Object,
                     Optional ByVal Oth_par As Object = Nothing) As Boolean

        Dim ret_val : ret_val = False

        '        Try

        If (Not IfNull(Obj_par)) Then

            Dim DataType = UCase(Obj_par.GetType.ToString)

            If (InStr(DataType, UCase("System.Collections.Generic.List`1[System.Collections.Generic.List`1[")) > 0) Then

                If (Not IsNothing(Oth_par)) Then

                    If (GetDataType(Oth_par) = EnumDataType.Int) Then

                        With Obj_par
                            If (Oth_par <= .Count - 1) Then
                                ret_val = .Item(Oth_par).Count > 0
                            End If
                        End With

                    End If

                End If


            ElseIf (InStr(DataType, UCase("System.Collections.Generic.List`1[")) > 0) Then

                ret_val = Obj_par.Count > 0

            End If

        End If

        '        Catch

        '        End Try

end_sub:

        IfValid = ret_val

    End Function

    Function IfNull(ByVal Obj_par As Object) As Boolean

        Dim ret_val

        If (IsNothing(Obj_par)) Then
            ret_val = True
        Else
            ret_val = Obj_par.GetType() Is GetType(DBNull)
        End If

        IfNull = ret_val

    End Function

    Function IfString(ByVal Obj_par As Object) As Boolean

        IfString = Obj_par.GetType() Is GetType(String)

    End Function

    Function IfDate(ByVal Obj_par As Object) As Boolean

        IfDate = Obj_par.GetType() Is GetType(Date)

    End Function

    Function IfBoolean(ByVal Obj_par As Object) As Boolean

        IfBoolean = Obj_par.GetType() Is GetType(Boolean)

    End Function

    Function IfInteger(ByVal Obj_par As Object) As Boolean

        With Obj_par
            IfInteger = .GetType() Is GetType(Integer)
        End With

    End Function

    Function IfSingle(ByVal Obj_par As Object) As Boolean

        With Obj_par
            IfSingle = .GetType() Is GetType(Single)
        End With

    End Function

    Function IfDouble(ByVal Obj_par As Object) As Boolean

        With Obj_par
            IfDouble = .GetType() Is GetType(Double)
        End With

    End Function

    Function IfDecimal(ByVal Obj_par As Object) As Boolean

        With Obj_par
            IfDecimal = .GetType() Is GetType(Decimal)
        End With

    End Function

    Function IfNumber(ByVal Obj_par As Object) As Boolean

        With Obj_par
            IfNumber = IsNumeric(Obj_par)
        End With

    End Function

    Function IfReal(ByVal Obj_par As Object) As Boolean

        With Obj_par
            IfReal = IfSingle(Obj_par) Or IfDouble(Obj_par) Or IfDecimal(Obj_par)
        End With

    End Function

    Function IfFileExists(ByVal FileName_par As Object,
                          Optional ByVal FilePath_par As Object = "") As Boolean

        Dim ret_val : ret_val = False

        Try
            ret_val = File.Exists(CombinePaths(FilePath_par, FileName_par))
        Catch

        End Try

        IfFileExists = ret_val

    End Function

    Function CombinePaths(ByVal P1 As Object, ByVal P2 As Object) As String

        Dim ret_val As String = ""

        Try
            ret_val = Path.Combine(P1, P2)
        Catch

        End Try

end_sub:

        CombinePaths = ret_val

    End Function

    Public Function TextWidth(ByVal Text_par As String, ByVal Font_par As Font) As Integer

        Dim ret_val As Integer

        '        Try

        ret_val = TextRenderer.MeasureText(Text_par, Font_par).Width

        '        Catch

        '        End Try

        TextWidth = ret_val

    End Function

    Public Function GetDataType(ByVal Data_par As Object) As Object

        Dim ret_val = EnumDataType.None

        If (IfString(Data_par)) Then
            ret_val = EnumDataType.Text
        ElseIf (IfInteger(Data_par)) Then
            ret_val = EnumDataType.Int
        ElseIf (IfReal(Data_par)) Then
            ret_val = EnumDataType.Real
        ElseIf (IfDate(Data_par)) Then
            ret_val = EnumDataType.Dt
        ElseIf (IfBoolean(Data_par)) Then
            ret_val = EnumDataType.Bool
        End If

        GetDataType = ret_val

    End Function

    Public Function GetDataType(ByVal Data_par As System.Data.DataColumn) As Object

        Dim ret_val = EnumDataType.None

        '        Try

        With Data_par

            Dim data_type = .DataType

            If (data_type = System.Type.GetType("System.String")) Then

                ret_val = EnumDataType.Text

            ElseIf (data_type = System.Type.GetType("System.Int16") Or
                    data_type = System.Type.GetType("System.Int32") Or
                    data_type = System.Type.GetType("System.Int64")) Then

                ret_val = EnumDataType.Int

            ElseIf (data_type = System.Type.GetType("System.Single") Or
                    data_type = System.Type.GetType("System.Decimal") Or
                    data_type = System.Type.GetType("System.Double")) Then

                ret_val = EnumDataType.Real

            ElseIf (data_type = System.Type.GetType("System.DateTime")) Then

                ret_val = EnumDataType.Dt

            ElseIf (data_type = System.Type.GetType("System.Boolean")) Then

                ret_val = EnumDataType.Bool

            End If

        End With

        '        Catch

        '        End Try

end_func:

        GetDataType = ret_val

    End Function

    Public Function ValidateData(ByVal Data_par As Object,
                                 Optional ByVal DataType_par As Object = "",
                                 Optional ByVal ReqdData_par As Object = PropValue,
                                 Optional ByVal OmitDataStat_par As Object = True,
                                 Optional ByVal OmitData_par As Object = "") As Object

        Dim ret_val = Data_par

        If (IfNull(ret_val)) Then ret_val = ""
        If (Len(DataType_par) = 0) Then DataType_par = GetDataType(Data_par)

        Dim OmitData_lcl = OmitData_par &
                           "nil" & DelimiterChar &
                           "null" & DelimiterChar

        Dim lstOmitData_lcl As New List(Of Object)
        lstOmitData_lcl.AddRange(Split(OmitData_lcl, DelimiterChar))

        If (DataType_par = EnumDataType.Text And OmitDataStat_par And lstOmitData_lcl.Count > 0) Then

            With lstOmitData_lcl

                For ele_no = 0 To .Count - 1
                    ret_val = Replace(UCase(ret_val), UCase(.Item(ele_no)), "", 1, -1, CompareMethod.Text)
                Next ele_no

            End With

        End If


        Dim rect_data = ""

        If (DataType_par = EnumDataType.Int Or DataType_par = EnumDataType.Real Or
            DataType_par = EnumDataType.Amount) Then

            rect_data = ""

            For i = 1 To Len(ret_val)

                Dim check_char = Mid(ret_val, i, 1)
                Dim proc_stat = False

                Select Case Asc(check_char)
                    Case Asc("0") To Asc("9")
                        proc_stat = True
                    Case Asc(".")
                        proc_stat = True
                End Select

                If (proc_stat) Then
                    rect_data = Val(rect_data & check_char)
                End If

            Next i

            If (Len(rect_data) = 0) Then
                If (DataType_par = EnumDataType.Int) Then
                    rect_data = 0
                ElseIf (DataType_par = EnumDataType.Real) Then
                    rect_data = 0.0
                ElseIf (DataType_par = EnumDataType.Amount) Then
                    rect_data = 0.0
                End If
            End If


        ElseIf (DataType_par = EnumDataType.Dt) Then

            rect_data = ret_val

            If (Not IsDate(rect_data)) Then
                rect_data = "01/01/1900"
            End If

            rect_data = DateValue(rect_data)


        Else

            rect_data = ret_val

        End If


        ret_val = rect_data

        If (ReqdData_par = PropValueForRelCondn) Then
            If (DataType_par = EnumDataType.Text Or DataType_par = EnumDataType.Dt) Then
                ret_val = "'" & ret_val & "'"
            End If
        End If


        ValidateData = ret_val

    End Function

    Public Function GetUniqueVal(Optional ByVal Pfx_par As Object = "") As Object

        Dim ret_val

        Pfx_par = UCase(Trim(Pfx_par))

        ret_val = DateTimeWithMillisec()
        ret_val = RandomNumber(CDbl("1.0"), CDbl("1000.0")) & "/" & ret_val

        If (Len(Trim(Pfx_par)) > 0) Then
            ret_val = Trim(Pfx_par) & "/" & ret_val
        End If

        GetUniqueVal = ret_val

    End Function

    Public Function RandomNumber(ByVal LowerLimit_par As Double, ByVal UpperLimit_par As Double) As Double

        Rnd()
        Randomize()

        RandomNumber = Rnd() * (UpperLimit_par - LowerLimit_par) + LowerLimit_par

    End Function

    Public Function MinDateInMonth(ByVal Month_par As Integer, ByVal Year_par As Integer) As Date

        MinDateInMonth = Format("01/" & Format(Month_par, "0#") & "/" & Year_par, DispFormat_dd_mm_yyyy)

    End Function

    Public Function MaxDateInMonth(ByVal Month_par As Integer, ByVal Year_par As Integer) As Date

        Dim min_date = Format("01/" & Format(Month_par, "0#") & "/" & Year_par, DispFormat_dd_mm_yyyy)

        MaxDateInMonth = DateAdd("d", -1, DateAdd("m", 1, min_date))

    End Function

    Public Function DateTimeWithMillisec() As Object

        DateTimeWithMillisec = Format(Now, "dd-MMM-yyyy HH:mm:ss") & "." &
                               Right(Format(Timer, "#0.00"), 2)
    End Function

    Public Function Timer() As Object

        Timer = Microsoft.VisualBasic.Timer()

    End Function

    Public Sub AddArrItem(Of T)(ByRef Arr_par As T(), ByVal Item_par As T)

        If Arr_par IsNot Nothing Then
            Array.Resize(Arr_par, Arr_par.Length + 1)
            Arr_par(Arr_par.Length - 1) = Item_par
        Else
            ReDim Arr_par(0)
            Arr_par(0) = Item_par
        End If

    End Sub

    Public Sub DelArrItem(Of T)(ByRef Arr_par As T(), ByVal Index_par As Object)

        ' Move elements after "index" down 1 position.

        Array.Copy(Arr_par, Index_par + 1, Arr_par, Index_par, UBound(Arr_par) - Index_par)

        ' Shorten by 1 element.

        ReDim Preserve Arr_par(UBound(Arr_par) - 1)

    End Sub

End Module

Public Module Progressbar_Related

    <Serializable()>
    Public Class clsProgressbar

        Implements ICloneable

        Dim _IsInitialized As Boolean

        Private WithEvents _Ctrl As ProgressBar

        Dim WithEvents _Timer As New Timers.Timer
        Public WithEvents _Bwrkr As New BackgroundWorker

        Private Delegate Sub ShowProgressDelegateType()
        Private _ShowProgressDelegate As ShowProgressDelegateType

        Sub New(ByRef Pbr_par As ProgressBar)

            _Ctrl = Pbr_par

            With _Ctrl
                .Visible = False
                .Minimum = 0
                .Maximum = 50
                .Step = 1
                .Style = ProgressBarStyle.Blocks
            End With

            With _Timer
                .Interval = 1000
            End With

            _Bwrkr.WorkerReportsProgress = True
            _Bwrkr.WorkerSupportsCancellation = True

            _ShowProgressDelegate = AddressOf ShowProgress

            IsInitialized = True

        End Sub

        Protected Overrides Sub Finalize()

            If (IsInitialized) Then

                _Timer.Stop()

            End If

        End Sub

        Public ReadOnly Property Ctrl() As ProgressBar

            Get
                Ctrl = _Ctrl
            End Get

        End Property

        Public ReadOnly Property Timer() As Timers.Timer

            Get
                Timer = _Timer
            End Get

        End Property

        Public Property IsInitialized() As Boolean

            Get
                IsInitialized = _IsInitialized
            End Get

            Set(ByVal Value As Boolean)
                _IsInitialized = Value
            End Set

        End Property

        Public Sub Reset()

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            _Ctrl.Value = _Ctrl.Minimum

            _Timer.AutoReset = True

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Sub ShowProgress()

            With _Ctrl

                If (.InvokeRequired) Then

                    .Invoke(_ShowProgressDelegate)

                Else

                    If (.Value = .Maximum) Then
                        .Value = .Minimum
                    Else
                        .Increment(.Step)
                    End If

                End If

            End With

        End Sub

        Public Sub _Start(ByRef oSrc_par As Object, ByVal ProcName_par As Object, ByVal ParamArr_par() As Object)

            '            Try

            If (_Bwrkr.IsBusy) Then GoTo end_sub

            Dim CallType_lcl As Object = CallType.Method
            Dim lstParam_lcl As New List(Of Object)

            lstParam_lcl.AddRange({oSrc_par, ProcName_par, CallType_lcl, ParamArr_par})

            With _Bwrkr
                .RunWorkerAsync(lstParam_lcl)
            End With

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub _Stop()

            '            Try

            _Bwrkr.CancelAsync()

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

        Private Sub _Timer_Elapsed(ByVal sender As Object,
                                   ByVal e As System.Timers.ElapsedEventArgs) Handles _Timer.Elapsed

            ShowProgress()

        End Sub

        Private Sub _Bwrkr_DoWork(ByVal sender As Object,
                                  ByVal e As System.ComponentModel.DoWorkEventArgs) Handles _Bwrkr.DoWork

            '            Try

            With _Bwrkr

                If .CancellationPending Then
                    e.Cancel = True
                    Exit Sub
                End If

                With e
                    Dim arrArg As Object() = .Argument(3)
                    CallByName(.Argument(0), .Argument(1), .Argument(2), arrArg)
                End With

                .ReportProgress(1)

            End With

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Sub _Bwrkr_ProgressChanged(ByVal sender As Object,
                                           ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles _Bwrkr.ProgressChanged

            If (Not IsInitialized Or _Ctrl.Visible) Then GoTo end_sub

            With _Ctrl

                Reset()

                .Visible = True

                _Timer.Start()

            End With

end_sub:

        End Sub

        Private Sub _Bwrkr_RunWorkerCompleted(ByVal sender As Object,
                                              ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles _Bwrkr.RunWorkerCompleted

            Exit Sub

            If (Not IsInitialized) Then GoTo end_sub

            With _Bwrkr

                If (.IsBusy) Then
                    .CancelAsync()
                End If

                _Timer.Stop()

                _Ctrl.Visible = False

            End With

end_sub:

        End Sub

    End Class

End Module

Public Module DataGridView_Related

    Public Const DgridColAutoSizeModeNone = DataGridViewAutoSizeColumnMode.None
    Public Const DgridColAutoSizeModeAllCells = DataGridViewAutoSizeColumnMode.AllCells
    Public Const DgridColAutoSizeModeAllCellsExceptHeader = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader
    Public Const DgridColAutoSizeModeFill = DataGridViewAutoSizeColumnMode.Fill

    Public Const DgridAlignTopLeft = DataGridViewContentAlignment.TopLeft
    Public Const DgridAlignTopRight = DataGridViewContentAlignment.TopRight
    Public Const DgridAlignTopCenter = DataGridViewContentAlignment.TopCenter
    Public Const DgridAlignMiddleLeft = DataGridViewContentAlignment.MiddleLeft
    Public Const DgridAlignMiddleRight = DataGridViewContentAlignment.MiddleRight
    Public Const DgridAlignMiddleCenter = DataGridViewContentAlignment.MiddleCenter
    Public Const DgridAlignBottomLeft = DataGridViewContentAlignment.BottomLeft
    Public Const DgridAlignBottomRight = DataGridViewContentAlignment.BottomRight
    Public Const DgridAlignBottomCenter = DataGridViewContentAlignment.BottomCenter

    Public Const DgridSortModeAutomatic = DataGridViewColumnSortMode.Automatic
    Public Const DgridSortModeNotSortable = DataGridViewColumnSortMode.NotSortable
    Public Const DgridSortModeProgrammatic = DataGridViewColumnSortMode.Programmatic

    <Serializable()>
    Public Class clsDataGridView

        Implements ICloneable

        Private WithEvents _Ctrl As New DataGridView

        Dim _IsInitialized As Boolean

        Dim _IfAllowUserToAddRows As Boolean
        Dim _IfReadOnly As Boolean

        Dim _DefaHeader As New clsGenData

        Dim _DefaData As New clsGenData
        Dim _LstDefaData As New List(Of clsGenData)

        Dim _LstHeader As New List(Of clsGenData)
        Dim _LstData As New List(Of List(Of clsGenData))

        Dim _LstHeaderRst As New List(Of clsRecordset)

        Dim _TqPrepareRow As clsThreadQueue

        Public Event CurrentCellDirtyStateChanged(ByVal sender As Object,
                                                  ByVal e As System.EventArgs)

        Public Event CellValueChanged(ByVal sender As Object,
                                      ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

        Public Event EditingControlShowing(ByVal sender As Object,
                                           ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs)

        Sub New(ByRef Dgrid_par As DataGridView,
                Optional ByVal AsyncRunStat_par As Object = True)

            _Ctrl = Dgrid_par

            New_Dgrid(AsyncRunStat_par)

        End Sub

        Sub New_Dgrid(Optional ByVal AsyncRunStat_par As Object = True)

            If (AsyncRunStat_par) Then
                If (RunDelegate(_Ctrl, New SubWithoutParDelegateType(AddressOf Me.New_Dgrid),
                                AsyncRunStat_par)) Then GoTo end_sub
            End If

            With _Ctrl
                .Visible = Not .Visible
                .AutoGenerateColumns = False
                .EnableHeadersVisualStyles = False
            End With

            IfAllowUserToAddRows = True

            With _DefaHeader
                .AutoSizeMode = DgridColAutoSizeModeAllCells
            End With

            With _DefaData
            End With

            _TqPrepareRow = New clsThreadQueue

            IsInitialized = True

            UpdateHeaderList(-1)
            UpdateDataList(-1, -1)

            SortingMode(DgridSortModeNotSortable)

            With _Ctrl
                .Visible = Not .Visible
            End With
end_sub:

        End Sub

        Protected Overrides Sub Finalize()

            MyBase.Finalize()

        End Sub

        Public Property Ctrl() As DataGridView

            Get
                Ctrl = _Ctrl
            End Get

            Set(ByVal Value As DataGridView)
                _Ctrl = Value
            End Set

        End Property

        Public Property IsInitialized() As Boolean

            Get
                IsInitialized = _IsInitialized
            End Get

            Set(ByVal Value As Boolean)
                _IsInitialized = Value
            End Set

        End Property

        Public Property IfAllowUserToAddRows() As Boolean

            Get
                IfAllowUserToAddRows = _IfAllowUserToAddRows
            End Get

            Set(ByVal Value As Boolean)
                _IfAllowUserToAddRows = Value
                _Ctrl.AllowUserToAddRows = _IfAllowUserToAddRows
            End Set

        End Property

        Public Property IfReadOnly() As Boolean

            Get
                IfReadOnly = _IfReadOnly
            End Get

            Set(ByVal Value As Boolean)
                _IfReadOnly = Value
                _Ctrl.ReadOnly = _IfReadOnly
            End Set

        End Property

        Private ReadOnly Property IfDefaultCellStyleSet() As Boolean

            Get
                IfDefaultCellStyleSet = Rows() > 0
            End Get

        End Property

        Public Property DefaHeader() As clsGenData

            Get
                DefaHeader = _DefaHeader
            End Get

            Set(ByVal Value As clsGenData)
                _DefaHeader = Value
            End Set

        End Property

        Public Property DefaData() As clsGenData

            Get
                DefaData = _DefaData
            End Get

            Set(ByVal Value As clsGenData)
                _DefaData = Value
            End Set

        End Property

        Public Property ListDefaData() As List(Of clsGenData)

            Get
                ListDefaData = _LstDefaData
            End Get

            Set(ByVal Value As List(Of clsGenData))
                _LstDefaData = Value
            End Set

        End Property

        Public ReadOnly Property IfHeaderInitialized() As Boolean

            Get
                IfHeaderInitialized = IsInitialized And Cols() > 0 And _LstHeader.Count = Cols()
            End Get

        End Property

        Public ReadOnly Property IfDataInitialized() As Boolean

            Get
                IfDataInitialized = IsInitialized And Rows() > 0 And Cols() > 0 And _LstData.Count = Rows()
            End Get

        End Property

        Private Sub SortingMode(ByVal SortMode_par As DataGridViewColumnSortMode,
                                Optional ByVal AsyncRunStat_par As Object = False)

            With _Ctrl

                If (AsyncRunStat_par) Then
                    If (RunDelegate(_Ctrl, New SubWithParDelegateType(AddressOf Me.SortingMode),
                                    SortMode_par, AsyncRunStat_par)) Then GoTo end_sub
                End If

                For Each column In .Columns
                    column.sortMode = SortMode_par
                Next column

            End With
end_sub:

        End Sub

        Public Sub Hide(Optional ByVal AsyncRunStat_par As Object = True)

            With _Ctrl

                If (AsyncRunStat_par) Then
                    If (RunDelegate(_Ctrl, New SubWithoutParDelegateType(AddressOf Me.Hide),
                                    AsyncRunStat_par)) Then GoTo end_sub
                End If

                .Visible = False

            End With
end_sub:

        End Sub

        Public Sub Show(Optional ByVal AsyncRunStat_par As Object = True)

            With _Ctrl

                If (AsyncRunStat_par) Then
                    If (RunDelegate(_Ctrl, New SubWithoutParDelegateType(AddressOf Me.Show),
                                    AsyncRunStat_par)) Then GoTo end_sub
                End If

                .Visible = True

            End With
end_sub:

        End Sub

        Public Sub Clear(Optional ByVal AsyncRunStat_par As Object = False)

            With _Ctrl

                If (AsyncRunStat_par) Then
                    If (RunDelegate(_Ctrl, New SubWithoutParDelegateType(AddressOf Me.Clear),
                                    AsyncRunStat_par)) Then GoTo end_sub
                End If

                If (Not IfHeaderInitialized) Then GoTo end_sub

                _LstData.Clear()

                .Rows.Clear()

            End With
end_sub:

        End Sub

        Public Sub GotoCell(ByVal Row_par As Object, ByVal Col_par As Object)

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidRow(Row_par) Or Not IfValidCol(Col_par)) Then GoTo end_sub

            Ctrl.CurrentCell = Ctrl.Rows(Row_par).Cells(Col_par)

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Function CurrentCell() As DataGridViewCell

            '            Try

            '            Catch

            '            End Try

end_sub:
            CurrentCell = Ctrl.CurrentCell

        End Function

        Public Sub SelectItem(ByVal Row_par As Object, ByVal Col_par As Object,
                              ByVal ItemVal_par As Object)

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidRow(Row_par) Or Not IfValidCol(Col_par) Or
                IsNothing(ItemVal_par)) Then GoTo end_sub

            GotoCell(Row_par, Col_par)

            SelectItem(ItemVal_par)

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub SelectItem(ByVal ItemVal_par As Object)

            '            Try

            If (Not IsInitialized Or IsNothing(CurrentCell) Or IsNothing(ItemVal_par)) Then GoTo end_sub

            SendKeys.Send("%{down}") : SendKeys.SendWait("{esc}")

            CurrentCell.Value = ItemVal_par

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Function GetColProp(ByVal Arg_par As Object) As Object

            Dim ret_val As Object = Nothing

            '        Try

            Dim Col_lcl As Object = Arg_par(0)
            Dim Prop_lcl As Object = Arg_par(1)
            Dim ColType_lcl As EnumColType = Arg_par(2)

            Dim oCol As Object

            oCol = _Ctrl.Columns(Col_lcl)

            If (IsNothing(oCol) Or IsNothing(Prop_lcl)) Then GoTo end_sub


            If (Prop_lcl = PropFrozenStat) Then

                ret_val = oCol.Frozen


            ElseIf (Prop_lcl = PropAutoSizeMode) Then

                ret_val = oCol.AutoSizeMode


            ElseIf (Prop_lcl = PropAlignment) Then

                If (ColType_lcl = EnumColType.Header Or (ColType_lcl = EnumColType.Data And Rows() = 0)) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                ret_val = oCol.Alignment


            ElseIf (Prop_lcl = PropFont) Then

                If (ColType_lcl = EnumColType.Header Or (ColType_lcl = EnumColType.Data And Rows() = 0)) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                ret_val = oCol.Font


            ElseIf (Prop_lcl = PropFontName) Then

                If (ColType_lcl = EnumColType.Header Or (ColType_lcl = EnumColType.Data And Rows() = 0)) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                If (Not IsNothing(oCol.font)) Then ret_val = oCol.Font.Name


            ElseIf (Prop_lcl = PropFontSize) Then

                If (ColType_lcl = EnumColType.Header Or (ColType_lcl = EnumColType.Data And Rows() = 0)) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                If (Not IsNothing(oCol.font)) Then ret_val = oCol.Font.Size


            ElseIf (Prop_lcl = PropFontStyle) Then

                If (ColType_lcl = EnumColType.Header Or (ColType_lcl = EnumColType.Data And Rows() = 0)) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                If (Not IsNothing(oCol.font)) Then ret_val = oCol.Font.style


            ElseIf (Prop_lcl = PropBackColor) Then

                If (ColType_lcl = EnumColType.Header Or (ColType_lcl = EnumColType.Data And Rows() = 0)) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                ret_val = oCol.BackColor


            ElseIf (Prop_lcl = PropForeColor) Then

                If (ColType_lcl = EnumColType.Header Or (ColType_lcl = EnumColType.Data And Rows() = 0)) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                ret_val = oCol.ForeColor


            ElseIf (Prop_lcl = PropText) Then

                oCol = oCol.HeaderCell

                ret_val = oCol.Value

            ElseIf (Prop_lcl = PropLocked) Then

                ret_val = oCol.ReadOnly

            End If

            '        Catch

            '        End Try

end_sub:
            GetColProp = ret_val

        End Function

        Private Sub SetColProp(ByVal Arg_par As Object,
                               Optional ByVal AsyncRunStat_par As Object = False)

            '        Try

            If (RunDelegate(_Ctrl, New SubWithParDelegateType(AddressOf Me.SetColProp), Arg_par,
                                AsyncRunStat_par)) Then GoTo end_sub

            Dim Col_lcl As Object = Arg_par.Item(0)
            Dim Prop_lcl As Object = Arg_par.Item(1)
            Dim PropValue_lcl As Object = Arg_par.Item(2)
            Dim ColType_lcl As EnumColType = Arg_par.Item(3)

            Dim oCol As Object

            oCol = _Ctrl.Columns(Col_lcl)

            If (IsNothing(oCol) Or IsNothing(Prop_lcl) Or IsNothing(PropValue_lcl)) Then GoTo end_sub


            If (Prop_lcl = PropFrozenStat) Then

                oCol.Frozen = PropValue_lcl


            ElseIf (Prop_lcl = PropAutoSizeMode) Then

                oCol.AutoSizeMode = PropValue_lcl


            ElseIf (Prop_lcl = PropAlignment) Then

                If (ColType_lcl = EnumColType.Header) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                oCol.Alignment = PropValue_lcl


            ElseIf (Prop_lcl = PropFont) Then

                If (ColType_lcl = EnumColType.Header) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                With oCol
                    .Font = New System.Drawing.Font(CStr(PropValue_lcl.name), CSng(PropValue_lcl.size))
                End With


            ElseIf (Prop_lcl = PropFontName) Then

                Dim FontName As String
                Dim FontSize As Single

                If (ColType_lcl = EnumColType.Header) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                With oCol

                    If (IsNothing(.font)) Then
                        FontSize = 10
                    Else
                        FontSize = .font.size
                    End If

                    FontName = CStr(PropValue_lcl)

                    .Font = New System.Drawing.Font(FontName, FontSize)

                End With


            ElseIf (Prop_lcl = PropFontSize) Then

                Dim FontName As String
                Dim FontSize As Single

                If (ColType_lcl = EnumColType.Header) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                With oCol

                    If (IsNothing(.font)) Then
                        FontName = "courier new"
                    Else
                        FontName = .font.name
                    End If

                    FontSize = CSng(PropValue_lcl)

                    .Font = New System.Drawing.Font(FontName, FontSize)

                End With


            ElseIf (Prop_lcl = PropFontStyle) Then

                Dim FontName As String
                Dim FontSize As Single
                Dim FontStyle As System.Drawing.FontStyle

                If (ColType_lcl = EnumColType.Header) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                With oCol

                    If (IsNothing(.font)) Then
                        FontName = "courier new"
                        FontSize = 10
                    Else
                        FontName = .font.name
                        FontSize = .font.size
                    End If

                    FontStyle = PropValue_lcl

                    .Font = New System.Drawing.Font(FontName, FontSize, FontStyle)

                End With


            ElseIf (Prop_lcl = PropBackColor) Then

                If (ColType_lcl = EnumColType.Header) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                With oCol
                    .BackColor = PropValue_lcl
                End With


            ElseIf (Prop_lcl = PropForeColor) Then

                If (ColType_lcl = EnumColType.Header) Then
                    oCol = oCol.HeaderCell.Style
                Else
                    oCol = oCol.DefaultCellStyle
                End If

                With oCol
                    .ForeColor = PropValue_lcl
                End With


            ElseIf (Prop_lcl = PropText) Then

                oCol = oCol.HeaderCell

                oCol.Value = PropValue_lcl


            ElseIf (Prop_lcl = PropLocked) Then

                oCol.ReadOnly = PropValue_lcl

            End If

            '        Catch

            '        End Try

end_sub:

        End Sub

        Private Function GetCellProp(ByVal Arg_par As Object) As Object

            Dim ret_val As Object = Nothing

            '        Try

            Dim Row_lcl As Object = Arg_par(0)
            Dim Col_lcl As Object = Arg_par(1)
            Dim Prop_lcl As Object = Arg_par(2)

            Dim oCol As Object

            oCol = _Ctrl.Rows(Row_lcl).Cells(Col_lcl)

            If (IsNothing(oCol) Or IsNothing(Prop_lcl)) Then GoTo end_sub


            If (Prop_lcl = PropAlignment) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                ret_val = oCol.Alignment


            ElseIf (Prop_lcl = PropFont) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                ret_val = oCol.Font


            ElseIf (Prop_lcl = PropFontName) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                ret_val = oCol.Font.Name


            ElseIf (Prop_lcl = PropFontSize) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                ret_val = oCol.Font.Size


            ElseIf (Prop_lcl = PropFontStyle) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                ret_val = oCol.Font.style


            ElseIf (Prop_lcl = PropBackColor) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                ret_val = oCol.BackColor


            ElseIf (Prop_lcl = PropForeColor) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                ret_val = oCol.ForeColor


            ElseIf (Prop_lcl = PropText) Then

                ret_val = oCol.Value


            ElseIf (Prop_lcl = PropLocked) Then

                ret_val = oCol.ReadOnly

            End If

            '        Catch

            '        End Try

end_sub:
            GetCellProp = ret_val

        End Function

        Private Sub SetCellProp(ByVal Arg_par As Object,
                                Optional ByVal AsyncRunStat_par As Object = False)

            '        Try

            If (AsyncRunStat_par) Then
                If (RunDelegate(_Ctrl, New SubWithParDelegateType(AddressOf Me.SetCellProp), Arg_par,
                                AsyncRunStat_par)) Then GoTo end_sub
            End If

            Dim Row_lcl As Object = Arg_par.Item(0)
            Dim Col_lcl As Object = Arg_par.Item(1)
            Dim Prop_lcl As Object = Arg_par.Item(2)
            Dim PropValue_lcl As Object = Arg_par.Item(3)

            Dim oCol As Object

            oCol = _Ctrl.Rows(Row_lcl).Cells(Col_lcl)

            If (IsNothing(oCol) Or IsNothing(Prop_lcl) Or IsNothing(PropValue_lcl)) Then GoTo end_sub


            If (Prop_lcl = PropAlignment) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                oCol.Alignment = Alignment(PropValue_lcl)


            ElseIf (Prop_lcl = PropFont) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                With oCol
                    .Font = New System.Drawing.Font(CStr(PropValue_lcl.name), CSng(PropValue_lcl.size))
                End With


            ElseIf (Prop_lcl = PropFontName) Then

                Dim FontName As String
                Dim FontSize As Single

                oCol = oCol.OwningColumn.DefaultCellStyle

                With oCol

                    If (IsNothing(.font)) Then
                        FontSize = 10
                    Else
                        FontSize = .font.size
                    End If

                    FontName = CStr(PropValue_lcl)

                    .Font = New System.Drawing.Font(FontName, FontSize)

                End With


            ElseIf (Prop_lcl = PropFontSize) Then

                Dim FontName As String
                Dim FontSize As Single

                oCol = oCol.OwningColumn.DefaultCellStyle

                With oCol

                    If (IsNothing(.font)) Then
                        FontName = "courier new"
                    Else
                        FontName = .font.name
                    End If

                    FontSize = CSng(PropValue_lcl)

                    .Font = New System.Drawing.Font(FontName, FontSize)

                End With


            ElseIf (Prop_lcl = PropFontStyle) Then

                Dim FontName As String
                Dim FontSize As Single
                Dim FontStyle As System.Drawing.FontStyle

                oCol = oCol.OwningColumn.DefaultCellStyle

                With oCol

                    If (IsNothing(.font)) Then
                        FontName = "courier new"
                        FontSize = 10
                    Else
                        FontName = .font.name
                        FontSize = .font.size
                    End If

                    FontStyle = PropValue_lcl

                    .Font = New System.Drawing.Font(FontName, FontSize, FontStyle)

                End With


            ElseIf (Prop_lcl = PropBackColor) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                With oCol
                    .BackColor = PropValue_lcl
                End With


            ElseIf (Prop_lcl = PropForeColor) Then

                oCol = oCol.OwningColumn.DefaultCellStyle

                With oCol
                    .ForeColor = PropValue_lcl
                End With


            ElseIf (Prop_lcl = PropText) Then

                oCol.Value = PropValue_lcl


            ElseIf (Prop_lcl = PropLocked) Then

                oCol.ReadOnly = PropValue_lcl

            End If

            '        Catch

            '        End Try

end_sub:

        End Sub

        Private Sub UpdateColFromList(ByVal Col_par As Object, ByVal ColType_par As EnumColType)

            '            Try

            If (Not IfHeaderInitialized Or ColType_par = EnumColType.Both Or
                (ColType_par = EnumColType.Data And Not IfDefaultCellStyleSet)) Then GoTo end_sub


            Dim FromCol, ToCol As Integer

            Col_par = ColIndex(Col_par)

            If (Col_par >= 0) Then

                If (Not IfValidCol(Col_par)) Then GoTo end_sub

                FromCol = Col_par + 1
                ToCol = FromCol

            Else

                FromCol = 1
                ToCol = Cols()

            End If


            Dim lst_det As New List(Of clsGenData)

            If (ColType_par = EnumColType.Header) Then
                lst_det = _LstHeader
            ElseIf (ColType_par = EnumColType.Data) Then
                lst_det = _LstDefaData
            End If


            For col_no = FromCol To ToCol

                Col_par = col_no - 1

                With lst_det.Item(Col_par)
                    SetCol({Col_par, PropText, .Text, ColType_par, False})
                    SetCol({Col_par, PropFrozenStat, .FrozenStat, ColType_par, False})
                    SetCol({Col_par, PropAutoSizeMode, .AutoSizeMode, ColType_par, False})
                    SetCol({Col_par, PropAlignment, .Alignment, ColType_par, False})
                    SetCol({Col_par, PropFontName, .FontName, ColType_par, False})
                    SetCol({Col_par, PropFontSize, .FontSize, ColType_par, False})
                    SetCol({Col_par, PropFontStyle, .FontStyle, ColType_par, False})
                    SetCol({Col_par, PropBackColor, .BackColor, ColType_par, False})
                    SetCol({Col_par, PropForeColor, .ForeColor, ColType_par, False})
                    SetCol({Col_par, PropLocked, .Locked, ColType_par, False})
                End With

            Next col_no

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Sub UpdateCellFromList(ByVal Row_par As Object, ByVal Col_par As Object)

            '            Try

            If (Not IfDataInitialized) Then GoTo end_sub


            Dim FromRow, ToRow As Integer

            If (Row_par >= 0) Then

                If (Not IfValidRow(Row_par)) Then GoTo end_sub

                FromRow = Row_par + 1
                ToRow = FromRow

            Else

                FromRow = 1
                ToRow = Rows()

            End If


            Dim FromCol, ToCol As Integer

            Col_par = ColIndex(Col_par)

            If (Col_par >= 0) Then

                If (Not IfValidCol(Col_par)) Then GoTo end_sub

                FromCol = Col_par + 1
                ToCol = FromCol

            Else

                FromCol = 1
                ToCol = Cols()

            End If


            For row_no = FromRow To ToRow

                For col_no = FromCol To ToCol

                    Row_par = row_no - 1
                    Col_par = col_no - 1

                    With _LstData.Item(Row_par).Item(Col_par)
                        SetCell(Row_par, Col_par, PropText, .Text, False)
                        SetCell(Row_par, Col_par, PropAlignment, .Alignment, False)
                        SetCell(Row_par, Col_par, PropFontName, .FontName, False)
                        SetCell(Row_par, Col_par, PropFontSize, .FontSize, False)
                        SetCell(Row_par, Col_par, PropFontStyle, .FontStyle, False)
                        SetCell(Row_par, Col_par, PropBackColor, .BackColor, False)
                        SetCell(Row_par, Col_par, PropForeColor, .ForeColor, False)
                        SetCell(Row_par, Col_par, PropLocked, .Locked, False)
                    End With

                Next col_no

            Next row_no

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Sub UpdateHeaderList(ByVal Col_par As Object,
                                     Optional ByVal DispCtrlType_par As Object = Nothing)

            '            Try

            If (Not IsInitialized Or Cols() = 0) Then GoTo end_sub


            Dim FromCol, ToCol As Integer

            Col_par = ColIndex(Col_par)

            If (Col_par >= 0) Then

                If (Not IfValidCol(Col_par)) Then GoTo end_sub

                FromCol = Col_par + 1
                ToCol = FromCol

            Else

                FromCol = 1
                ToCol = Cols()

            End If


            With _LstHeader

                For col_no = FromCol To ToCol

                    If (.Count < col_no) Then

                        .Add(New clsGenData)

                        .Item(col_no - 1) = DefaHeader.Clone

                        With .Item(col_no - 1)
                            .DispControlType = DispCtrlType_par
                            .BindingObj = _Ctrl.Columns(col_no - 1)
                        End With

                        With _LstHeaderRst
                            If (IsNothing(DispCtrlType_par)) Then
                                .Add(Nothing)
                            Else
                                .Add(New clsRecordset)
                            End If
                        End With

                        UpdateColFromList(col_no - 1, EnumColType.Header)

                        With ListDefaData
                            .Add(New clsGenData)
                            .Item(col_no - 1) = DefaData.Clone
                            UpdateCellFromList(-1, col_no - 1)
                        End With

                    End If

                Next

            End With


            If (Not IfValidCol(Col_par)) Then GoTo end_sub

            With _LstHeader.Item(Col_par)
                .Text = GetHeadCol(Col_par, PropText)
                .FrozenStat = GetHeadCol(Col_par, PropFrozenStat)
                .AutoSizeMode = GetHeadCol(Col_par, PropAutoSizeMode)
                .Alignment = GetHeadCol(Col_par, PropAlignment)
                .FontName = GetHeadCol(Col_par, PropFontName)
                .FontSize = GetHeadCol(Col_par, PropFontSize)
                .FontStyle = GetHeadCol(Col_par, PropFontStyle)
                .BackColor = GetHeadCol(Col_par, PropBackColor)
                .ForeColor = GetHeadCol(Col_par, PropForeColor)
                .Locked = GetHeadCol(Col_par, PropLocked)
            End With


            If (IfDefaultCellStyleSet) Then

                With _LstDefaData.Item(Col_par)
                    .Text = GetDataCol(Col_par, PropText)
                    .Alignment = GetDataCol(Col_par, PropAlignment)
                    .FontName = GetDataCol(Col_par, PropFontName)
                    .FontSize = GetDataCol(Col_par, PropFontSize)
                    .FontStyle = GetDataCol(Col_par, PropFontStyle)
                    .BackColor = GetDataCol(Col_par, PropBackColor)
                    .ForeColor = GetDataCol(Col_par, PropForeColor)
                    .Locked = GetDataCol(Col_par, PropLocked)
                End With

            End If

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Sub UpdateDataList(ByVal Row_par As Integer, ByVal Col_par As Object)

            '            Try

            If (Not IsInitialized Or Rows() = 0 Or Cols() = 0) Then GoTo end_sub


            Dim FromRow, ToRow As Integer

            If (Row_par >= 0) Then

                If (Not IfValidRow(Row_par)) Then GoTo end_sub

                FromRow = Row_par + 1
                ToRow = FromRow

            Else

                FromRow = 1
                ToRow = Rows()

            End If


            Dim FromCol, ToCol As Integer

            Col_par = ColIndex(Col_par)

            If (Col_par >= 0) Then

                If (Not IfValidCol(Col_par)) Then GoTo end_sub

                FromCol = Col_par + 1
                ToCol = FromCol

            Else

                FromCol = 1
                ToCol = Cols()

            End If


            For row_no = FromRow To ToRow

                With _LstData

                    If (.Count < row_no) Then .Add(New List(Of clsGenData))

                    For col_no = FromCol To ToCol

                        With .Item(row_no - 1)

                            If (.Count < col_no) Then

                                .Add(New clsGenData)

                                .Item(col_no - 1) = ListDefaData(col_no - 1).Clone

                                With .Item(col_no - 1)
                                    .BindingObj = _Ctrl.Rows(row_no - 1).Cells(col_no - 1)
                                End With

                                UpdateCellFromList(row_no - 1, col_no - 1)

                            End If

                        End With

                    Next

                End With

            Next row_no


            If (Not IfValidRow(Row_par) Or Not IfValidCol(Col_par)) Then GoTo end_sub

            With _LstData.Item(Row_par).Item(Col_par)
                .Text = GetCell(Row_par, Col_par, PropText)
                .Alignment = GetCell(Row_par, Col_par, PropAlignment)
                .FontName = GetCell(Row_par, Col_par, PropFontName)
                .FontSize = GetCell(Row_par, Col_par, PropFontSize)
                .FontStyle = GetCell(Row_par, Col_par, PropFontStyle)
                .BackColor = GetCell(Row_par, Col_par, PropBackColor)
                .ForeColor = GetCell(Row_par, Col_par, PropForeColor)
                .Locked = GetCell(Row_par, Col_par, PropLocked)
            End With

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub AddRow(Optional ByVal AsyncRunStat_par As Object = True)

            If (AsyncRunStat_par) Then
                If (RunDelegate(_Ctrl, New SubWithParDelegateType(AddressOf Me.AddRow), {AsyncRunStat_par},
                                AsyncRunStat_par)) Then GoTo end_sub
            End If

            If (Not IfHeaderInitialized) Then GoTo end_sub

            With _Ctrl

                .Rows.Add()

                If (Not IfDefaultCellStyleSet) Then
                    UpdateColFromList(-1, EnumColType.Data)
                End If

                UpdateDataList(Rows() - 1, -1)

            End With
end_sub:

        End Sub

        Public Sub AddCol(ByVal Name_par As String, ByVal Header_par As String,
                          Optional ByVal DispCtrlType_par As Object = Nothing,
                          Optional ByVal AsyncRunStat_par As Object = True)

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            AddCol(CreateList({Name_par, Header_par, DispCtrlType_par}), AsyncRunStat_par)

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Sub AddCol(ByVal Arg_par As Object,
                           Optional ByVal AsyncRunStat_par As Object = False)

            If (AsyncRunStat_par) Then
                If (RunDelegate(_Ctrl, New SubWithParDelegateType(AddressOf Me.AddCol), Arg_par,
                                AsyncRunStat_par)) Then GoTo end_sub
            End If

            With _Ctrl.Columns

                Dim Name_lcl As String = Arg_par.Item(0)
                Dim Header_lcl As String = Arg_par.Item(1)
                Dim DispCtrlType_lcl As Object = Arg_par.Item(2)

                If (Not IsNothing(DispCtrlType_lcl)) Then

                    If (Not SetColCtrl(Name_lcl, Header_lcl, DispCtrlType_lcl)) Then
                        .Add(Name_lcl, Header_lcl)
                    End If

                Else

                    .Add(Name_lcl, Header_lcl)

                End If


                UpdateHeaderList(Cols() - 1, DispCtrlType_lcl)
                UpdateDataList(-1, -1)

            End With

            SortingMode(DataGridViewColumnSortMode.NotSortable)

end_sub:

        End Sub

        Public Function ColIndex(ByVal Col_par As Object) As Integer

            Dim ret_val As Integer = -1

            Try

                With _Ctrl.Columns

                    If (IfString(Col_par)) Then
                        ret_val = .Item(CStr(Col_par)).Index
                    Else
                        ret_val = Col_par
                    End If

                End With

            Catch

            End Try

end_sub:
            ColIndex = ret_val

        End Function

        Public Function IfValidRow(ByVal Row_par As Object) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            If (Not IsInitialized Or Rows() = 0) Then GoTo end_sub

            ret_val = (Row_par >= 0 And Row_par <= Rows() - 1)

            '            Catch

            '            End Try

end_sub:

            IfValidRow = ret_val

        End Function

        Public Function IfValidCol(ByVal Col_par As Object) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            If (Not IsInitialized Or Cols() = 0) Then GoTo end_sub

            Col_par = ColIndex(Col_par)

            ret_val = (Col_par >= 0 And Col_par <= Cols() - 1)

            '            Catch

            '            End Try

end_sub:

            IfValidCol = ret_val

        End Function

        Public Function Rows() As Integer

            Dim ret_val As Integer = 0

            If (Not IsInitialized) Then GoTo end_sub

            ret_val = _Ctrl.Rows.Count

end_sub:
            Rows = ret_val

        End Function

        Public Function Cols() As Integer

            Dim ret_val As Integer = 0

            If (Not IsInitialized) Then GoTo end_sub

            ret_val = _Ctrl.Columns.Count

end_sub:
            Cols = ret_val

        End Function

        Private Function Alignment(ByVal Value_par As Object) As Object

            Dim ret_val As Object = Value_par

            If (Value_par = EnumAlignment.Left) Then
                ret_val = DgridAlignMiddleLeft
            ElseIf (Value_par = EnumAlignment.Right) Then
                ret_val = DgridAlignMiddleRight
            ElseIf (Value_par = EnumAlignment.Center) Then
                ret_val = DgridAlignMiddleCenter
            End If

            Alignment = ret_val

        End Function

        Private Function SetColCtrl(ByVal Name_par As String,
                                    ByVal Header_par As Object,
                                    ByVal DispCtrlType_par As EnumDispControlType,
                                    Optional ByVal AsyncRunStat_par As Object = True) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            If (Not IsInitialized) Then GoTo end_sub


            If (DispCtrlType_par = EnumDispControlType.ComboBox) Then

                Dim DgvCmbCol As New DataGridViewComboBoxColumn()

                With DgvCmbCol
                    .HeaderText = Header_par
                    .Name = Name_par
                End With

                _Ctrl.Columns.Add(DgvCmbCol)

                ret_val = True

            End If

            '            Catch 

            '            End Try

end_sub:
            SetColCtrl = ret_val

        End Function

        Public Function SetColCtrlData(ByVal Col_par As Object,
                                       ByVal LstData_par As List(Of Object),
                                       Optional ByVal AsyncRunStat_par As Object = True) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidCol(Col_par)) Then GoTo end_sub


            If (_LstHeader.Item(Col_par).DispControlType = EnumDispControlType.ComboBox) Then

                Dim DgvCmbCol As DataGridViewComboBoxColumn = _Ctrl.Columns(Col_par)

                With DgvCmbCol
                    .Items.Clear()
                    .DataSource = LstData_par
                End With

                ret_val = True

            End If

            '            Catch 

            '            End Try

end_sub:

            SetColCtrlData = ret_val

        End Function

        Public Function SetColCtrlData(ByVal Col_par As Object,
                                       ByVal LstData_par As List(Of clsGenData),
                                       Optional ByVal AsyncRunStat_par As Object = True) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidCol(Col_par)) Then GoTo end_sub


            If (_LstHeader.Item(Col_par).DispControlType = EnumDispControlType.ComboBox) Then

                Dim DgvCmbCol As DataGridViewComboBoxColumn = _Ctrl.Columns(Col_par)

                With DgvCmbCol
                    .Items.Clear()
                    .DataSource = LstData_par
                    .DisplayMember = "text"
                    .ValueMember = "code"
                End With

                ret_val = True

            End If

            '            Catch 

            '            End Try

end_sub:
            SetColCtrlData = ret_val

        End Function

        Public Function SetColCtrlData(ByVal Col_par As Object,
                                       ByVal DbConn_par As Object,
                                       ByVal SqlDet_par As Object,
                                       ByVal FieldPfx_par As Object,
                                       ByVal DisplayFieldName_par As Object,
                                       ByVal ValueFieldName_par As Object,
                                       Optional ByVal AsyncRunStat_par As Object = True) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            If (AsyncRunStat_par) Then
                ret_val = RunDelegate(New FuncWithParDelegateType(AddressOf Me.SetColCtrlData_Sql_FldPfx_DispFld_ValFld),
                                      {Col_par, DbConn_par, SqlDet_par, FieldPfx_par,
                                       DisplayFieldName_par, ValueFieldName_par},
                                      AsyncRunStat_par)
            Else
                ret_val = Me.SetColCtrlData_Sql_FldPfx_DispFld_ValFld({Col_par, DbConn_par, SqlDet_par, FieldPfx_par,
                                                                       DisplayFieldName_par, ValueFieldName_par})
            End If

            '            Catch 

            '            End Try

end_sub:
            SetColCtrlData = ret_val

        End Function

        Private Function SetColCtrlData_Sql_FldPfx_DispFld_ValFld(ByVal Arg_par) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            Dim Col_lcl As Object = Arg_par(0)
            Dim DbConn_lcl As Object = Arg_par(1)
            Dim SqlDet_lcl As Object = Arg_par(2)
            Dim FieldPfx_lcl As Object = Arg_par(3)
            Dim DisplayFieldName_lcl As Object = Arg_par(4)
            Dim ValueFieldName_lcl As Object = Arg_par(5)

            Col_lcl = ColIndex(Col_lcl)
            DisplayFieldName_lcl = Trim(DisplayFieldName_lcl)
            ValueFieldName_lcl = Trim(ValueFieldName_lcl)

            If (Not IsInitialized Or Not IfValidCol(Col_lcl) Or Len(DisplayFieldName_lcl) = 0 Or
                Len(ValueFieldName_lcl) = 0 Or IsNothing(_LstHeaderRst.Item(Col_lcl))) Then GoTo end_sub


            If (_LstHeader.Item(Col_lcl).DispControlType = EnumDispControlType.ComboBox) Then

                With _LstHeaderRst.Item(Col_lcl)
                    If (Not .Open(DbConn_lcl, SqlDet_lcl, FieldPfx_lcl, ValueFieldName_lcl)) Then
                        GoTo end_sub
                    End If
                End With

                Dim DgvCmbCol As DataGridViewComboBoxColumn = _Ctrl.Columns(Col_lcl)

                With DgvCmbCol
                    .DataSource = _LstHeaderRst.Item(Col_lcl).Dtable
                    .DisplayMember = DisplayFieldName_lcl
                    .ValueMember = ValueFieldName_lcl
                End With

                ret_val = True

            End If

            '            Catch 

            '            End Try

end_sub:
            SetColCtrlData_Sql_FldPfx_DispFld_ValFld = ret_val

        End Function

        Public Function GetHeadCol(ByVal Col_par As Object, ByVal Prop_par As Object,
                                   Optional ByVal AsyncRunStat_par As Object = True) As Object

            Dim ret_val As Object = Nothing

            '            Try

            If (AsyncRunStat_par) Then
                ret_val = RunDelegate(New FuncWithParDelegateType(AddressOf Me.GetCol),
                                      {Col_par, Prop_par, EnumColType.Header},
                                      AsyncRunStat_par)
            Else
                ret_val = Me.GetCol({Col_par, Prop_par, EnumColType.Header})
            End If

            '            Catch

            '            End Try

end_sub:

            GetHeadCol = ret_val

        End Function

        Public Sub SetHeadCol(ByVal Col_par As Object, ByVal Prop_par As Object, ByVal Value_par As Object,
                              Optional ByVal UpdateRelList As Boolean = True,
                              Optional ByVal AsyncRunStat_par As Object = True)

            '            Try

            If (Prop_par = PropDispControlType) Then

            Else

                RunThread(AddressOf Me.SetCol,
                          {Col_par, Prop_par, Value_par, EnumColType.Header, UpdateRelList},
                          AsyncRunStat_par)
            End If

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Function GetDataCol(ByVal Col_par As Object, ByVal Prop_par As Object,
                                   Optional ByVal AsyncRunStat_par As Object = True) As Object

            Dim ret_val As Object = Nothing

            '            Try

            If (AsyncRunStat_par) Then
                ret_val = RunDelegate(New FuncWithParDelegateType(AddressOf Me.GetCol),
                                      {Col_par, Prop_par, EnumColType.Data},
                                      AsyncRunStat_par)
            Else
                ret_val = Me.GetCol({Col_par, Prop_par, EnumColType.Data})
            End If

            '            Catch

            '            End Try

end_sub:
            GetDataCol = ret_val

        End Function

        Public Sub SetDataCol(ByVal Col_par As Object, ByVal Prop_par As Object, ByVal Value_par As Object,
                      Optional ByVal UpdateRelList As Boolean = True,
                      Optional ByVal AsyncRunStat_par As Object = True)

            '            Try

            RunThread(AddressOf Me.SetCol,
                      {Col_par, Prop_par, Value_par, EnumColType.Data, UpdateRelList},
                      AsyncRunStat_par)

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Function GetCol(ByVal Arg_par As Object) As Object

            Dim ret_val As Object = Nothing

            '            Try

            Dim Col_par As Object = Arg_par(0)
            Dim Prop_par As Object = Arg_par(1)
            Dim ColType_par As EnumColType = Arg_par(2)

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidCol(Col_par)) Then GoTo end_sub


            If (Prop_par = PropDataType) Then

                ret_val = ListDefaData(Col_par).DataType

            ElseIf (Prop_par = PropDataAddress) Then

                ret_val = ListDefaData(Col_par).DataAddress

            ElseIf (Prop_par = PropImpStat) Then

                ret_val = ListDefaData(Col_par).ImpStat

            Else

                ret_val = GetColProp({Col_par, Prop_par, ColType_par})

            End If

            '            Catch

            '            End Try

end_sub:
            GetCol = ret_val

        End Function

        Private Sub SetCol(ByVal Arg_par As Object)

            '            Try

            Dim Col_lcl As Object = Arg_par(0)
            Dim Prop_lcl As Object = Arg_par(1)
            Dim Value_lcl As Object = Arg_par(2)
            Dim ColType_lcl As EnumColType = Arg_par(3)
            Dim UpdateRelList As Boolean = Arg_par(4)

            Col_lcl = ColIndex(Col_lcl)

            If (Not IsInitialized Or Not IfValidCol(Col_lcl) Or
                ColType_lcl = EnumColType.Both) Then GoTo end_sub


            Dim lstValue As New List(Of Object)

            If (IsArray(Value_lcl)) Then
                lstValue.AddRange(Value_lcl)
            Else
                lstValue.Add(Value_lcl)
            End If


            If (Prop_lcl = PropDataType) Then

                With ListDefaData(Col_lcl)

                    .DataType = lstValue.Item(0)

                    If (Alignment(.Alignment) <> GetHeadCol(Col_lcl, PropAlignment)) Then
                        SetColProp(CreateList({Col_lcl, PropAlignment, Alignment(.Alignment), EnumColType.Header}))
                        SetColProp(CreateList({Col_lcl, PropAlignment, Alignment(.Alignment), EnumColType.Data}))
                    End If

                End With

            ElseIf (Prop_lcl = PropDataAddress) Then

                ListDefaData(Col_lcl).DataAddress = lstValue

            ElseIf (Prop_lcl = PropImpStat) Then

                ListDefaData(Col_lcl).ImpStat = lstValue.Item(0)

            Else

                If (ColType_lcl = EnumColType.Data And Not IfDefaultCellStyleSet) Then
                    SetDataColList(Col_lcl, Prop_lcl, lstValue.Item(0))
                Else
                    SetColProp(CreateList({Col_lcl, Prop_lcl, lstValue.Item(0), ColType_lcl}))
                End If

            End If


            If (UpdateRelList) Then UpdateHeaderList(Col_lcl)

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Function GetCell(ByVal Row_par As Object, ByVal Col_par As Object, ByVal Prop_par As Object,
                                Optional ByVal AsyncRunStat_par As Object = False) As Object

            Dim ret_val As Object = Nothing

            '            Try

            If (AsyncRunStat_par) Then
                ret_val = RunDelegate(New FuncWithParDelegateType(AddressOf Me.GetCell_Arg),
                                      {Row_par, Col_par, Prop_par},
                                      AsyncRunStat_par)
            Else
                ret_val = Me.GetCell_Arg({Row_par, Col_par, Prop_par})
            End If

            '            Catch

            '            End Try

end_sub:
            GetCell = ret_val

        End Function

        Private Function GetCell_Arg(ByVal Arg_par As Object) As Object

            Dim ret_val As Object = Nothing

            '            Try

            Dim Row_lcl As Object = Arg_par(0)
            Dim Col_lcl As Object = Arg_par(1)
            Dim Prop_lcl As Object = Arg_par(2)

            Col_lcl = ColIndex(Col_lcl)

            If (Not IsInitialized Or Not IfValidRow(Row_lcl) Or Not IfValidCol(Col_lcl)) Then GoTo end_sub


            If (Prop_lcl = PropDataType) Then

                ret_val = _LstData.Item(Row_lcl).Item(Col_lcl).DataType

            ElseIf (Prop_lcl = PropDataAddress) Then

                ret_val = _LstData.Item(Row_lcl).Item(Col_lcl).DataAddress

            ElseIf (Prop_lcl = PropImpStat) Then

                ret_val = _LstData.Item(Row_lcl).Item(Col_lcl).ImpStat

            Else

                ret_val = GetCellProp({Col_lcl, Prop_lcl})

            End If

            '            Catch

            '            End Try

end_sub:
            GetCell_Arg = ret_val

        End Function

        Public Sub SetCell(ByVal Row_par As Object, ByVal Col_par As Object,
                           ByVal Prop_par As Object, ByVal Value_par As Object,
                           Optional ByVal UpdateRelList As Boolean = True,
                           Optional ByVal AsyncRunStat_par As Object = False)

            '            Try

            If (AsyncRunStat_par) Then
                RunDelegate(New SubWithParDelegateType(AddressOf Me.SetCell_Arg),
                            {Row_par, Col_par, Prop_par, Value_par, UpdateRelList, AsyncRunStat_par},
                            AsyncRunStat_par)
            Else
                Me.SetCell_Arg({Row_par, Col_par, Prop_par, Value_par, UpdateRelList, AsyncRunStat_par})
            End If

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Sub SetCell_Arg(ByVal Arg_par As Object)

            '            Try

            Dim Row_lcl As Object = Arg_par(0)
            Dim Col_lcl As Object = Arg_par(1)
            Dim Prop_lcl As Object = Arg_par(2)
            Dim Value_lcl As Object = Arg_par(3)
            Dim UpdateRelList As Boolean = Arg_par(4)
            Dim AsyncRunStat_par As Object = Arg_par(5)

            Col_lcl = ColIndex(Col_lcl)

            If (Not IsInitialized Or Not IfValidRow(Row_lcl) Or Not IfValidCol(Col_lcl)) Then GoTo end_sub


            Dim lstValue As New List(Of Object)

            If (IsArray(Value_lcl)) Then
                lstValue.AddRange(Value_lcl)
            Else
                lstValue.Add(Value_lcl)
            End If


            If (Prop_lcl = PropDataAddress) Then

                _LstData.Item(Row_lcl).Item(Col_lcl).DataAddress = lstValue

            ElseIf (Prop_lcl = PropImpStat) Then

                _LstData.Item(Row_lcl).Item(Col_lcl).ImpStat = lstValue.Item(0)

            Else

                SetCellProp(CreateList({Row_lcl, Col_lcl, Prop_lcl, lstValue.Item(0)}), AsyncRunStat_par)

            End If


            If (UpdateRelList) Then UpdateDataList(Row_lcl, Col_lcl)

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Function GetHeadColList(ByVal Col_par As Object, ByVal Prop_par As Object) As Object

            Dim ret_val As Object = Nothing

            '            Try

            ret_val = GetList(Col_par, Prop_par, EnumColType.Header)

            '            Catch

            '            End Try

end_sub:
            GetHeadColList = ret_val

        End Function

        Public Sub SetHeadColList(ByVal Col_par As Object,
                                  ByVal Prop_var As Object,
                                  ByVal Value_par As Object)

            '            Try

            SetList(Col_par, Prop_var, Value_par, EnumColType.Header)

            '            Catch

            '            End Try
end_sub:

        End Sub

        Public Function GetDataColList(ByVal Col_par As Object, ByVal Prop_par As Object) As Object

            Dim ret_val As Object = Nothing

            '            Try

            ret_val = GetList(Col_par, Prop_par, EnumColType.Data)

            '            Catch

            '            End Try

end_sub:
            GetDataColList = ret_val

        End Function

        Public Sub SetDataColList(ByVal Col_par As Object,
                                  ByVal Prop_var As Object,
                                  ByVal Value_par As Object)

            '            Try

            SetList(Col_par, Prop_var, Value_par, EnumColType.Data)

            '            Catch

            '            End Try
end_sub:

        End Sub

        Private Function GetList(ByVal Col_par As Object, ByVal Prop_var As Object,
                                 ByVal ColType_par As EnumColType) As Object

            Dim ret_val As Object = Nothing

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidCol(Col_par) Or ColType_par = EnumColType.Both) Then GoTo end_sub


            Dim lst_det As New List(Of clsGenData)

            If (ColType_par = EnumColType.Header) Then
                lst_det = _LstHeader
            ElseIf (ColType_par = EnumColType.Data) Then
                lst_det = _LstDefaData
            End If


            With lst_det

                If (Prop_var = PropCode) Then

                    ret_val = .Item(Col_par).Code

                ElseIf (Prop_var = PropText) Then

                    ret_val = .Item(Col_par).Text

                ElseIf (Prop_var = PropOthText) Then

                    ret_val = .Item(Col_par).OthText

                ElseIf (Prop_var = PropDataType) Then

                    ret_val = .Item(Col_par).DataType

                ElseIf (Prop_var = PropDispFormat) Then

                    ret_val = .Item(Col_par).DispFormat

                ElseIf (Prop_var = PropFrozenStat) Then

                    ret_val = .Item(Col_par).FrozenStat

                ElseIf (Prop_var = PropAutoSizeMode) Then

                    ret_val = .Item(Col_par).AutoSizeMode

                ElseIf (Prop_var = PropAlignment) Then

                    ret_val = .Item(Col_par).Alignment

                ElseIf (Prop_var = PropFontName) Then

                    ret_val = .Item(Col_par).FontName

                ElseIf (Prop_var = PropFontSize) Then

                    ret_val = .Item(Col_par).FontSize

                ElseIf (Prop_var = PropFontStyle) Then

                    ret_val = .Item(Col_par).FontStyle

                ElseIf (Prop_var = PropBackColor) Then

                    ret_val = .Item(Col_par).BackColor

                ElseIf (Prop_var = PropForeColor) Then

                    ret_val = .Item(Col_par).ForeColor

                ElseIf (Prop_var = PropLocked) Then

                    ret_val = .Item(Col_par).Locked

                ElseIf (Prop_var = PropDataAddress) Then

                    ret_val = .Item(Col_par).DataAddress

                ElseIf (Prop_var = PropImpStat) Then

                    ret_val = .Item(Col_par).ImpStat

                End If

            End With

            '            Catch

            '            End Try

end_sub:
            GetList = ret_val

        End Function

        Public Function GetList(ByVal Row_par As Object,
                                ByVal Col_par As Object,
                                ByVal Prop_var As Object) As Object

            Dim ret_val As Object = Nothing

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidRow(Row_par) Or Not IfValidCol(Col_par)) Then GoTo end_sub


            With _LstData.Item(Row_par)

                If (Prop_var = PropCode) Then

                    ret_val = .Item(Col_par).Code

                ElseIf (Prop_var = PropText) Then

                    ret_val = .Item(Col_par).Text

                ElseIf (Prop_var = PropOthText) Then

                    ret_val = .Item(Col_par).OthText

                ElseIf (Prop_var = PropDataType) Then

                    ret_val = .Item(Col_par).DataType

                ElseIf (Prop_var = PropDispFormat) Then

                    ret_val = .Item(Col_par).DispFormat

                ElseIf (Prop_var = PropAlignment) Then

                    ret_val = .Item(Col_par).Alignment

                ElseIf (Prop_var = PropFontName) Then

                    ret_val = .Item(Col_par).FontName

                ElseIf (Prop_var = PropFontSize) Then

                    ret_val = .Item(Col_par).FontSize

                ElseIf (Prop_var = PropFontStyle) Then

                    ret_val = .Item(Col_par).FontStyle

                ElseIf (Prop_var = PropBackColor) Then

                    ret_val = .Item(Col_par).BackColor

                ElseIf (Prop_var = PropForeColor) Then

                    ret_val = .Item(Col_par).ForeColor

                ElseIf (Prop_var = PropLocked) Then

                    ret_val = .Item(Col_par).Locked

                ElseIf (Prop_var = PropDataAddress) Then

                    ret_val = .Item(Col_par).DataAddress

                ElseIf (Prop_var = PropImpStat) Then

                    ret_val = .Item(Col_par).ImpStat

                End If

            End With

            '            Catch

            '            End Try

end_sub:
            GetList = ret_val

        End Function

        Private Sub SetList(ByVal Col_par As Object,
                            ByVal Prop_var As Object,
                            ByVal Value_par As Object,
                            ByVal ColType_par As EnumColType)

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidCol(Col_par) Or ColType_par = EnumColType.Both) Then GoTo end_sub


            Dim lstValue As New List(Of Object)

            If (IsArray(Value_par)) Then
                lstValue.AddRange(Value_par)
            Else
                lstValue.Add(Value_par)
            End If


            Dim lst_det As New List(Of clsGenData)

            If (ColType_par = EnumColType.Header) Then
                lst_det = _LstHeader
            ElseIf (ColType_par = EnumColType.Data) Then
                lst_det = _LstDefaData
            End If


            With lst_det

                If (Prop_var = PropCode) Then

                    .Item(Col_par).Code = lstValue.Item(0)

                ElseIf (Prop_var = PropText) Then

                    .Item(Col_par).Text = lstValue.Item(0)

                ElseIf (Prop_var = PropOthText) Then

                    .Item(Col_par).OthText = lstValue.Item(0)

                ElseIf (Prop_var = PropDataType) Then

                    .Item(Col_par).DataType = lstValue.Item(0)

                ElseIf (Prop_var = PropDispFormat) Then

                    .Item(Col_par).DispFormat = lstValue.Item(0)

                ElseIf (Prop_var = PropFrozenStat) Then

                    .Item(Col_par).FrozenStat = lstValue.Item(0)

                ElseIf (Prop_var = PropAutoSizeMode) Then

                    .Item(Col_par).AutoSizeMode = lstValue.Item(0)

                ElseIf (Prop_var = PropAlignment) Then

                    .Item(Col_par).Alignment = lstValue.Item(0)

                ElseIf (Prop_var = PropFontName) Then

                    .Item(Col_par).FontName = lstValue.Item(0)

                ElseIf (Prop_var = PropFontSize) Then

                    .Item(Col_par).FontSize = lstValue.Item(0)

                ElseIf (Prop_var = PropFontStyle) Then

                    .Item(Col_par).FontStyle = lstValue.Item(0)

                ElseIf (Prop_var = PropBackColor) Then

                    .Item(Col_par).BackColor = lstValue.Item(0)

                ElseIf (Prop_var = PropForeColor) Then

                    .Item(Col_par).ForeColor = lstValue.Item(0)

                ElseIf (Prop_var = PropLocked) Then

                    .Item(Col_par).Locked = lstValue.Item(0)

                ElseIf (Prop_var = PropDataAddress) Then

                    .Item(Col_par).DataAddress = lstValue

                ElseIf (Prop_var = PropImpStat) Then

                    .Item(Col_par).ImpStat = lstValue.Item(0)

                End If


                UpdateColFromList(Col_par, ColType_par)

            End With

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub SetList(ByVal Row_par As Integer, ByVal Col_par As Object,
                           ByVal Prop_var As Object, ByVal Value_par As Object)

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidRow(Row_par) Or
                Not IfValidCol(Col_par)) Then GoTo end_sub


            Dim lstValue As New List(Of Object)

            If (IsArray(Value_par)) Then
                lstValue.AddRange(Value_par)
            Else
                lstValue.Add(Value_par)
            End If


            With _LstData.Item(Row_par)

                If (Prop_var = PropCode) Then

                    .Item(Col_par).Code = lstValue.Item(0)

                ElseIf (Prop_var = PropText) Then

                    .Item(Col_par).Text = lstValue.Item(0)

                ElseIf (Prop_var = PropOthText) Then

                    .Item(Col_par).OthText = lstValue.Item(0)

                ElseIf (Prop_var = PropDataType) Then

                    .Item(Col_par).DataType = lstValue.Item(0)

                ElseIf (Prop_var = PropDispFormat) Then

                    .Item(Col_par).DispFormat = lstValue.Item(0)

                ElseIf (Prop_var = PropAlignment) Then

                    .Item(Col_par).Alignment = lstValue.Item(0)

                ElseIf (Prop_var = PropFontName) Then

                    .Item(Col_par).FontName = lstValue.Item(0)

                ElseIf (Prop_var = PropFontSize) Then

                    .Item(Col_par).FontSize = lstValue.Item(0)

                ElseIf (Prop_var = PropFontStyle) Then

                    .Item(Col_par).FontStyle = lstValue.Item(0)

                ElseIf (Prop_var = PropBackColor) Then

                    .Item(Col_par).BackColor = lstValue.Item(0)

                ElseIf (Prop_var = PropForeColor) Then

                    .Item(Col_par).ForeColor = lstValue.Item(0)

                ElseIf (Prop_var = PropLocked) Then

                    .Item(Col_par).Locked = lstValue.Item(0)

                ElseIf (Prop_var = PropDataAddress) Then

                    .Item(Col_par).DataAddress = lstValue

                ElseIf (Prop_var = PropImpStat) Then

                    .Item(Col_par).ImpStat = lstValue.Item(0)

                End If


                UpdateCellFromList(Row_par, Col_par)

            End With

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Function GetHeadRst(ByVal Col_par As Object) As clsRecordset

            Dim ret_val As clsRecordset = Nothing

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidCol(Col_par)) Then GoTo end_sub

            ret_val = _LstHeaderRst(Col_par)

            '            Catch

            '            End Try

end_sub:
            GetHeadRst = ret_val

        End Function

        Public Sub GetHeadRst(ByVal Col_par As Object,
                              ByVal Value_par As clsRecordset)

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfValidCol(Col_par)) Then GoTo end_sub

            _LstHeaderRst(Col_par) = Value_par

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub CopyCol(ByVal SrcCol_par As Object,
                           ByVal TrgCol_par As Object,
                           ByVal ColType_par As EnumColType)

            '            Try

            SrcCol_par = ColIndex(SrcCol_par)
            TrgCol_par = ColIndex(TrgCol_par)

            If (Not IsInitialized Or Not IfValidCol(SrcCol_par) Or Not IfValidCol(TrgCol_par)) Then GoTo end_sub

            With _LstHeader
                .Item(TrgCol_par) = .Item(SrcCol_par).Clone
                UpdateColFromList(TrgCol_par, ColType_par)
            End With

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub CopyCell(ByVal SrcRow_par As Object, ByVal SrcCol_par As Object,
                            ByVal TrgRow_par As Object, ByVal TrgCol_par As Object)

            '            Try

            SrcCol_par = ColIndex(SrcCol_par)
            TrgCol_par = ColIndex(TrgCol_par)

            If (Not IsInitialized Or Not IfValidRow(SrcRow_par) Or Not IfValidCol(SrcCol_par) Or
                Not IfValidRow(TrgRow_par) Or Not IfValidCol(TrgCol_par)) Then GoTo end_sub

            With _LstData
                .Item(TrgRow_par).Item(TrgCol_par) = .Item(SrcRow_par).Item(SrcCol_par).Clone
                UpdateCellFromList(TrgRow_par, TrgCol_par)
            End With

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub Prepare(ByVal Wbook_par As clsExcelWbook,
                           ByVal WbookFromRow_par As Object,
                           ByVal WbookToRow_par As Object,
                           Optional ByVal WsheetNo_par As Object = 1,
                           Optional ByVal AsyncRunStat_par As Object = False)

            '            Try

            If (_TqPrepareRow.IfValidCurrentThread) Then GoTo end_sub

            Hide()

            If (IfNumber(WbookFromRow_par)) Then
                WbookFromRow_par = CInt(WbookFromRow_par)
            Else
                WbookFromRow_par = 0
            End If

            If (IfNumber(WbookToRow_par)) Then
                WbookToRow_par = CInt(WbookToRow_par)
            Else
                WbookToRow_par = 0
            End If

            If (Not IsInitialized Or Not IfHeaderInitialized Or Not Wbook_par.IfValidWsheetNo(WsheetNo_par) Or
                WbookFromRow_par <= 0 Or WbookToRow_par <= 0 Or WbookFromRow_par > WbookToRow_par) Then GoTo end_sub

            Prepare_Arg({Wbook_par, WbookFromRow_par, WbookToRow_par, WsheetNo_par}, AsyncRunStat_par)

            '            Catch

            '            End Try

end_sub:
            If (Rows() > 0) Then Show()

        End Sub

        Private Sub Prepare_Arg(ByVal Arg_par As Object,
                                Optional ByVal AsyncRunStat_par As Object = True)

            '            Try

            Dim Wbook_lcl As clsExcelWbook = Arg_par(0)
            Dim WbookFromRow_lcl As Object = Arg_par(1)
            Dim WbookToRow_lcl As Object = Arg_par(2)
            Dim WsheetNo_lcl As Object = Arg_par(3)

            Dim PrepareRowThreadStart_lcl As ParameterizedThreadStart
            PrepareRowThreadStart_lcl = New ParameterizedThreadStart(AddressOf Me.PrepareRow_Arg)

            Clear()

            _TqPrepareRow.Clear()


            For wbook_row_no = WbookFromRow_lcl To WbookToRow_lcl

                If (AsyncRunStat_par) Then
                    _TqPrepareRow.Enqueue(New Thread(PrepareRowThreadStart_lcl),
                                          {Wbook_lcl, wbook_row_no, WsheetNo_lcl, True, AsyncRunStat_par})
                Else
                    PrepareRow_Arg({Wbook_lcl, wbook_row_no, WsheetNo_lcl, False, AsyncRunStat_par})
                End If

            Next wbook_row_no


            If (AsyncRunStat_par) Then _TqPrepareRow._Start()

            '            Catch ex As Exception

            '            End Try

end_sub:

        End Sub

        Public Function PrepareRow(ByVal Wbook_par As clsExcelWbook,
                                   ByVal WbookRow_par As Object,
                                   Optional ByVal WsheetNo_par As Object = 1,
                                   Optional ByVal AsyncRunStat_par As Object = False) As Object

            Dim ret_val : ret_val = False

            '            Try

            If (_TqPrepareRow.IfValidCurrentThread) Then GoTo end_sub

            If (IfNumber(WbookRow_par)) Then
                WbookRow_par = CInt(WbookRow_par)
            Else
                WbookRow_par = 0
            End If

            If (Not IsInitialized Or Not IfHeaderInitialized Or Not Wbook_par.IfValidWsheetNo(WsheetNo_par) Or
                WbookRow_par <= 0) Then GoTo end_sub

            ret_val = PrepareRow_Arg({Wbook_par, WbookRow_par, WsheetNo_par, False, AsyncRunStat_par})

            '            Catch

            '            End Try

end_sub:
            PrepareRow = ret_val

        End Function

        Private Function PrepareRow_Arg(ByVal Arg_par As Object) As Object

            Dim ret_val : ret_val = False

            '            Try

            Dim Wbook_lcl As clsExcelWbook = Arg_par(0)
            Dim WbookRow_lcl = Arg_par(1)
            Dim WsheetNo_lcl = Arg_par(2)
            Dim EnqueueStat_lcl = Arg_par(3)
            Dim AsyncRunStat_lcl = Arg_par(4)

            Dim lst_data_read As New List(Of clsGenData)
            lst_data_read.AddRange(_LstDefaData)

            If (Wbook_lcl.GetData(WbookRow_lcl, lst_data_read, WsheetNo_lcl, AsyncRunStat_lcl)) Then

                AddRow(AsyncRunStat_lcl)

                With _LstData.Item(Rows() - 1)
                    .Clear()
                    .AddRange(lst_data_read)
                End With

            End If

            If (EnqueueStat_lcl) Then _TqPrepareRow.Dequeue()

            ret_val = True

            '            Catch ex As Exception

            '            End Try

end_sub:
            PrepareRow_Arg = ret_val

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

        Private Sub _Ctrl_CellValueChanged(ByVal sender As Object,
                                           ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) _
                                           Handles _Ctrl.CellValueChanged

            RaiseEvent CellValueChanged(sender, e)

        End Sub

        Private Sub _Ctrl_CurrentCellDirtyStateChanged(ByVal sender As Object,
                                                       ByVal e As System.EventArgs) _
                                                       Handles _Ctrl.CurrentCellDirtyStateChanged

            If (_Ctrl.IsCurrentCellDirty) Then

                '                _Ctrl.CommitEdit(DataGridViewDataErrorContexts.Commit)

                RaiseEvent CurrentCellDirtyStateChanged(sender, e)

            End If

        End Sub

        Private Sub _Ctrl_DataError(ByVal sender As Object,
                                    ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) _
                                    Handles _Ctrl.DataError

            If (e.Context = (DataGridViewDataErrorContexts.Formatting Or
                             DataGridViewDataErrorContexts.PreferredSize)) Then

                e.ThrowException = False

            End If

        End Sub

        Private Sub _Ctrl_EditingControlShowing(ByVal sender As Object,
          ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) _
          Handles _Ctrl.EditingControlShowing

            RaiseEvent EditingControlShowing(sender, e)

        End Sub

        Private Sub clsDataGridView_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles Me.CellValueChanged

        End Sub

    End Class

    Public Sub CopyProp(ByRef Dgrid_par As DataGridView, ByVal Src_par As Object, ByVal Trg_par As Object,
                        ByVal Data_par() As Object, ByVal ColType_par As EnumColType)

        '        Try

        Dim oSrcCol, oTrgCol As Object

        With Dgrid_par

            oSrcCol = .Columns(Src_par)
            oTrgCol = .Columns(Trg_par)

            If (oSrcCol Is Nothing Or oTrgCol Is Nothing) Then GoTo end_sub

        End With


        Dim FromColType, ToColType As EnumColType

        If (ColType_par = EnumColType.Both) Then
            FromColType = 0
            ToColType = 1
        Else
            FromColType = ColType_par
            ToColType = ColType_par
        End If


        Dim lst_of_prop As New List(Of Object)

        lst_of_prop.AddRange(Data_par)


        With lst_of_prop

            For ColType = FromColType To ToColType

                For PropNo = 0 To .Count - 1

                    If (.Item(PropNo) = PropAutoSizeMode) Then

                        With Dgrid_par

                            With .Columns(Trg_par)
                                .AutoSizeMode = Dgrid_par.Columns(Src_par).AutoSizeMode
                            End With

                        End With


                    ElseIf (.Item(PropNo) = PropAlignment) Then

                        With Dgrid_par

                            With .Columns(Trg_par)

                                If (ColType = EnumColType.Header) Then

                                    With Dgrid_par.Columns(Trg_par).HeaderCell.Style
                                        .Alignment = Dgrid_par.Columns(Src_par).HeaderCell.Style.Alignment
                                    End With

                                Else

                                    With Dgrid_par.Columns(Trg_par).DefaultCellStyle
                                        .Alignment = Dgrid_par.Columns(Src_par).DefaultCellStyle.Alignment
                                    End With

                                End If

                            End With

                        End With


                    ElseIf (.Item(PropNo) = PropFont) Then

                        With Dgrid_par

                            With .Columns(Trg_par)

                                If (ColType = EnumColType.Header) Then

                                    With Dgrid_par.Columns(Trg_par).HeaderCell.Style
                                        .Font = Dgrid_par.Columns(Src_par).HeaderCell.Style.Font
                                    End With

                                Else

                                    With Dgrid_par.Columns(Trg_par).DefaultCellStyle
                                        .Font = Dgrid_par.Columns(Src_par).DefaultCellStyle.Font
                                    End With

                                End If

                            End With

                        End With


                    ElseIf (.Item(PropNo) = PropColor) Then

                        With Dgrid_par

                            With .Columns(Trg_par)

                                If (ColType = EnumColType.Header) Then

                                    With Dgrid_par.Columns(Trg_par).HeaderCell.Style
                                        .BackColor = Dgrid_par.Columns(Src_par).HeaderCell.Style.BackColor
                                        .ForeColor = Dgrid_par.Columns(Src_par).HeaderCell.Style.ForeColor
                                    End With

                                Else


                                    With Dgrid_par.Columns(Trg_par).DefaultCellStyle
                                        .BackColor = Dgrid_par.Columns(Src_par).DefaultCellStyle.BackColor
                                        .ForeColor = Dgrid_par.Columns(Src_par).DefaultCellStyle.ForeColor
                                    End With

                                End If

                            End With

                        End With


                    ElseIf (.Item(PropNo) = PropBackColor) Then

                        With Dgrid_par

                            With .Columns(Trg_par)

                                If (ColType = EnumColType.Header) Then

                                    With Dgrid_par.Columns(Trg_par).HeaderCell.Style
                                        .BackColor = Dgrid_par.Columns(Src_par).HeaderCell.Style.BackColor
                                    End With

                                Else

                                    With Dgrid_par.Columns(Trg_par).DefaultCellStyle
                                        .BackColor = Dgrid_par.Columns(Src_par).DefaultCellStyle.BackColor
                                    End With

                                End If

                            End With

                        End With


                    ElseIf (.Item(PropNo) = PropForeColor) Then

                        With Dgrid_par

                            With .Columns(Trg_par)

                                If (ColType = EnumColType.Header) Then

                                    With Dgrid_par.Columns(Trg_par).HeaderCell.Style
                                        .ForeColor = Dgrid_par.Columns(Src_par).HeaderCell.Style.ForeColor
                                    End With

                                Else

                                    With Dgrid_par.Columns(Trg_par).DefaultCellStyle
                                        .ForeColor = Dgrid_par.Columns(Src_par).DefaultCellStyle.ForeColor
                                    End With

                                End If

                            End With

                        End With

                    End If

                Next PropNo

            Next ColType

        End With

        '        Catch

        '        End Try

end_sub:

    End Sub

End Module

Public Module Excel_Related

    Public ExcelApp As Excel.Application

    <Serializable()>
    Public Class clsExcelWbook

        Implements ICloneable

        Dim _IsInitialized As Boolean

        Dim _WorkBook As Excel.Workbook
        Dim _WorkSheet As Excel.Worksheet

        Dim _FileName As String

        Sub New()

            If (IsNothing(ExcelApp)) Then
                ExcelApp = New Excel.Application
            End If

        End Sub

        Protected Overrides Sub Finalize()

            If (IsInitialized) Then

            End If

        End Sub

        Public ReadOnly Property IsInitialized() As Boolean

            Get
                IsInitialized = _IsInitialized
            End Get

        End Property

        Public Property WorkBook() As Excel.Workbook

            Get
                WorkBook = _WorkBook
            End Get

            Set(ByVal Value As Excel.Workbook)
                _WorkBook = Value
            End Set

        End Property

        Public Property WorkSheet() As Excel.Worksheet

            Get
                WorkSheet = _WorkSheet
            End Get

            Set(ByVal Value As Excel.Worksheet)
                _WorkSheet = Value
            End Set

        End Property

        Public Property FileName() As String

            Get
                FileName = _FileName
            End Get

            Set(ByVal Value As String)
                _FileName = Value
            End Set

        End Property

        Public Function Open(ByVal FileName_par As Object,
                             Optional ByVal AsyncRunStat_par As Object = True) As Boolean

            _IsInitialized = False

            '            Try

            If (AsyncRunStat_par) Then
                _IsInitialized = RunDelegate(New FuncWithParDelegateType(AddressOf Me.Open_FileName),
                                             {FileName_par},
                                             AsyncRunStat_par)
            Else
                _IsInitialized = Me.Open_FileName({FileName_par})
            End If

            '            Catch

            '            End Try

end_sub:
            Open = IsInitialized

        End Function

        Private Function Open_FileName(ByVal Arg_par As Object) As Boolean

            _IsInitialized = False

            '            Try

            Dim FileName_lcl As String = Arg_par(0)

            FileName = FileName_lcl

            _WorkBook = ExcelApp.Workbooks.Open(FileName)

            _IsInitialized = True

            '            Catch

            '            End Try

end_sub:
            Open_FileName = IsInitialized

        End Function

        Public Function Close(Optional ByVal FileName_par As Object = "",
                              Optional ByVal Mesg_par As Boolean = False,
                              Optional ByVal AsyncRunStat_par As Object = True) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            If (AsyncRunStat_par) Then
                ret_val = RunDelegate(New FuncWithParDelegateType(AddressOf Me.Close_FileName_Mesg),
                                      {FileName_par, Mesg_par},
                                      AsyncRunStat_par)
            Else
                ret_val = Me.Close_FileName_Mesg({FileName_par, Mesg_par})
            End If

            '            Catch

            '            End Try

end_sub:
            Close = ret_val

        End Function

        Public Function Close_FileName_Mesg(ByVal Arg_par As Object) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            Dim FileName_lcl As Object = Arg_par(0)
            Dim Mesg_lcl As Boolean = Arg_par(1)

            If (Not IsInitialized) Then GoTo end_sub

            WorkBook.Close()

            _IsInitialized = False

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            Close_FileName_Mesg = ret_val

        End Function

        Public Function IfValidWsheetNo(ByVal WsheetNo_par As Object) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            ret_val = WsheetNo_par >= 1 And WsheetNo_par <= WsheetCount()

            '            Catch

            '            End Try

end_sub:
            IfValidWsheetNo = ret_val

        End Function

        Public Function IfValidCol(ByVal Col_par As Object) As Boolean

            IfValidCol = Len(Trim(Col_par)) > 0

        End Function

        Public Function WsheetCount() As Integer

            Dim ret_val As Object = 0

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            ret_val = WorkBook.Sheets.Count

            '            Catch

            '            End Try

end_sub:
            WsheetCount = ret_val

        End Function

        Public Function ColIndex(ByVal ColSub1_par As Object,
                                 Optional ByVal ColSub2_par As Object = "") As Object

            Dim ret_val As Object = ""

            '            Try

            Dim ColSub As Object


            ColSub = ColSub1_par

            If (IfInteger(ColSub)) Then

                If (ColSub >= 1 And ColSub <= 26) Then
                    ColSub = Chr(Asc("a") - 1 + ColSub)
                Else
                    ColSub = ""
                End If

            ElseIf (IfString(ColSub)) Then

                ColSub = LCase(Trim(ColSub))

                If (Len(ColSub) > 0 And Len(ColSub) = 1) Then
                    If (Not (Asc(ColSub) >= Asc("a") And Asc(ColSub) <= Asc("z"))) Then ColSub = ""
                Else
                    ColSub = ""
                End If

            Else

                ColSub = ""

            End If

            ColSub1_par = ColSub


            ColSub = ColSub2_par

            If (Len(ColSub1_par) > 0 And Len(Trim(ColSub)) > 0) Then

                If (IfInteger(ColSub)) Then

                    If (ColSub >= 1 And ColSub <= 26) Then
                        ColSub = Chr(Asc("a") - 1 + ColSub)
                    Else
                        ColSub = ""
                    End If

                ElseIf (IfString(ColSub)) Then

                    ColSub = LCase(Trim(ColSub))

                    If (Len(ColSub) > 0 And Len(ColSub) = 1) Then
                        If (Not (Asc(ColSub) >= Asc("a") And Asc(ColSub) <= Asc("z"))) Then ColSub = ""
                    Else
                        ColSub = ""
                    End If

                Else

                    ColSub = ""

                End If

            Else

                ColSub = ""

            End If


            If (Len(Trim(ColSub2_par)) > 0 And Len(ColSub) = 0) Then ColSub1_par = ""

            ColSub2_par = ColSub


            ret_val = UCase(ColSub1_par & ColSub2_par)

            '            Catch

            '            End Try

end_sub:
            ColIndex = ret_val

        End Function

        Public Function GetData(ByVal Row_par As Object, ByVal Col_par As Object,
                                Optional ByVal WsheetNo_par As Object = 1,
                                Optional ByVal AsyncRunStat_par As Object = True) As Object

            Dim ret_val As Object = Nothing

            '            Try

            Col_par = ColIndex(Col_par)

            If (Not IsInitialized Or Not IfInteger(Row_par) Or Not IfValidCol(Col_par) Or
                Not IfValidWsheetNo(WsheetNo_par)) Then GoTo end_sub

            If (AsyncRunStat_par) Then
                ret_val = RunDelegate(New FuncWithParDelegateType(AddressOf Me.GetData_Row_Col_Wsheet),
                                      {Row_par, Col_par, WsheetNo_par},
                                      AsyncRunStat_par)
            Else
                ret_val = Me.GetData_Row_Col_Wsheet({Row_par, Col_par, WsheetNo_par})
            End If

            '            Catch

            '            End Try

end_sub:
            GetData = ret_val

        End Function

        Private Function GetData_Row_Col_Wsheet(ByVal Arg_par As Object) As Object

            Dim ret_val As Object = Nothing

            '            Try

            Dim Row_lcl As Object = Arg_par(0)
            Dim Col_lcl As Object = Arg_par(1)
            Dim WsheetNo_lcl As Object = Arg_par(2)

            With _WorkBook

                Dim CellNo As Object

                _WorkSheet = .Sheets(WsheetNo_lcl)

                CellNo = Col_lcl & Row_lcl

                With _WorkSheet.Range(CellNo)
                    ret_val = .Value
                End With

            End With

            '            Catch

            '            End Try

end_sub:
            GetData_Row_Col_Wsheet = ret_val

        End Function

        Public Function GetData(ByVal Row_par As Object, ByRef LstCol_par As List(Of clsGenData),
                                Optional ByVal WsheetNo_par As Object = 1,
                                Optional ByVal AsyncRunStat_par As Object = True) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            If (AsyncRunStat_par) Then
                ret_val = RunDelegate(New FuncWithParDelegateType(AddressOf Me.GetData_Row_LstCol),
                                      {Row_par, LstCol_par, WsheetNo_par},
                                      AsyncRunStat_par)
            Else
                ret_val = Me.GetData_Row_LstCol({Row_par, LstCol_par, WsheetNo_par})
            End If

            '            Catch

            '            ret_val = false

            '            End Try

end_sub:
            GetData = ret_val

        End Function

        Private Function GetData_Row_LstCol(ByVal Arg_par As Object) As Boolean

            Dim ret_val As Boolean = False

            '            Try

            Dim Row_lcl As Object = Arg_par(0)
            Dim LstCol_lcl As List(Of clsGenData) = Arg_par(1)
            Dim WsheetNo_lcl As Object = Arg_par(2)

            If (Not IfInteger(Row_lcl) Or Not IsInitialized Or LstCol_lcl.Count = 0 Or
                Not IfValidWsheetNo(WsheetNo_lcl)) Then GoTo end_sub

            Dim lst_data_read As New List(Of clsGenData)

            lst_data_read.AddRange(LstCol_lcl)


            With lst_data_read

                For item_no = 0 To .Count - 1

                    With .Item(item_no)

                        .Text = Nothing

                        For addr_ele_no = 0 To .DataAddress.Count - 1

                            .Text = .Text &
                                    GetData(Row_lcl, .DataAddress.Item(addr_ele_no), WsheetNo_lcl)

                        Next addr_ele_no

                    End With

                Next item_no

                ret_val = True

            End With

            '            Catch

            '            ret_val = false

            '            End Try

end_sub:
            If (ret_val) Then

                With LstCol_lcl
                    .Clear()
                    .AddRange(lst_data_read)
                End With

            End If

            GetData_Row_LstCol = ret_val

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

Public Module Database_Related

#Region "Common Database Fields Dictonary"

    Public Const fldNo = "no"
    Public Const fldOthNo = "oth_no"
    Public Const fldSno = "sno"

    Public Const fldCode = "code"
    Public Const fldMastCode = "mast_code"
    Public Const fldOthCode = "oth_code"
    Public Const fldLectCode = "lect_code"
    Public Const fldActLectCode = "act_lect_code"
    Public Const fldHeadLectCode = "head_lect_code"

    Public Const fldDetTypeNo = "det_type_no"
    Public Const fldDetTypeDet = "det_type_det"

    Public Const fldDetNo = "det_no"
    Public Const fldDet = "det"

    Public Const fldCertificateDetNo = "certificate_det_no"

    Public Const fldShortDet = "short_det"
    Public Const fldShort = "short"
    Public Const fldShort_2 = "short_2"

    Public Const fldOthDet = "oth_det"
    Public Const fldOthDet_2 = "oth_det_2"

    Public Const fldYearNo = "year_no"
    Public Const fldYearDet = "year_det"

    Public Const fldMonNo = "mon_no"
    Public Const fldMonDet = "mon_det"

    Public Const fldDob = "dob"
    Public Const fldDoj = "doj"
    Public Const fldRelieveDt = "relieve_dt"
    Public Const fldDate = "date"
    Public Const fldFromDate = "from_date"
    Public Const fldToDate = "to_date"
    Public Const fldChangeDate = "change_date"
    Public Const fldSubmitDate = "submit_date"
    Public Const fldCollectDate = "collect_date"
    Public Const fldOthDate = "oth_date"

    Public Const fldFromTime = "from_time"
    Public Const fldToTime = "to_time"

    Public Const fldWeekDay = "week_day"

    Public Const fldName = "name"
    Public Const fldMastName = "mast_name"
    Public Const fldName_2 = "name_2"
    Public Const fldMastName_2 = "mast_name_2"
    Public Const fldOthName = "oth_name"
    Public Const fldGender = "gender"
    Public Const fldSex = "sex"
    Public Const fldSalutation = "salutation"
    Public Const fldAddr1 = "addr1"
    Public Const fldAddr2 = "addr2"
    Public Const fldAddr3 = "addr3"
    Public Const fldCity = "city"
    Public Const fldPinNo = "pin_no"
    Public Const fldPhNo = "ph_no"
    Public Const fldResiPhNo = "resi_ph_no"
    Public Const fldCellNo = "cell_no"
    Public Const fldOthCellNo = "oth_cell_no"
    Public Const fldMailId = "mail_id"
    Public Const fldSmsDeliStat = "sms_deli_stat"

    Public Const fldParentName = "parent_name"
    Public Const fldParentQualificationNo = "parent_qualification_no"
    Public Const fldParentOccupationNo = "parent_occupation_no"
    Public Const fldParentOccupationOthDet = "parent_occupation_oth_det"
    Public Const fldParentIncome = "parent_income"
    Public Const fld_parent_ph_no = "parent_ph_no"
    Public Const fldParentCellNo = "parent_cell_no"
    Public Const fldParentMailId = "parent_mail_id"
    Public Const fldParentSmsDeliStat = "parent_sms_deli_stat"

    Public Const fldProfessionNo = "profession_no"
    Public Const fldProfessionDet = "profession_det"
    Public Const fldProfession = "profession"

    Public Const fldRepBy = "rep_by"
    Public Const fldRepByOthDet = "rep_by_oth_det"

    Public Const fldBankAcNo = "bank_ac_no"
    Public Const fldBankAcDet = "bank_ac_det"

    Public Const fldRcNo = "rc_no"
    Public Const fldRcDet = "rc_det"

    Public Const fldPanNo = "pan_no"
    Public Const fldPanDet = "pan_det"

    Public Const fldGstNo = "gst_no"
    Public Const fldGstDet = "gst_det"

    Public Const fldCstNo = "cst_no"
    Public Const fldCstDet = "cst_det"

    Public Const fldTinNo = "tin_no"
    Public Const fldTinDet = "tin_det"

    Public Const fldLfNo = "lf_no"
    Public Const fldLfDet = "lf_det"

    Public Const fldPfAcNo = "pf_ac_no"
    Public Const fldPfAcDet = "pf_ac_det"

    Public Const fldOthAddr1 = "oth_addr1"
    Public Const fldOthAddr2 = "oth_addr2"
    Public Const fldOthAddr3 = "oth_addr3"
    Public Const fldOthCity = "oth_city"

    Public Const fldParti = "parti"
    Public Const fldParti_2 = "parti_2"
    Public Const fldParti_3 = "parti_3"
    Public Const fldParti_4 = "parti_4"
    Public Const fldParti_5 = "parti_5"

    Public Const fldRemarks = "remarks"

    Public Const fldPayScale = "pay_scale"
    Public Const fldSalary = "salary"
    Public Const fldOthSalary = "oth_salary"
    Public Const fldDaPerc = "da_perc"
    Public Const fldDa = "da"
    Public Const fldHraPerc = "hra_perc"
    Public Const fldHra = "hra"
    Public Const fldPf = "pf"
    Public Const fldPtax = "ptax"

    Public Const fldExemptLeaves = "exempt_leaves"

    Public Const fldLocalityNo = "locality_no"
    Public Const fldLocalityCode = "locality_code"
    Public Const fldLocalityDet = "locality_det"

    Public Const fldTehsilNo = "tehsil_no"
    Public Const fldTehsilCode = "tehsil_code"
    Public Const fldTehsilDet = "tehsil_det"

    Public Const fldDistrictNo = "district_no"
    Public Const fldDistrictDet = "district_det"

    Public Const fldStateNo = "state_no"
    Public Const fldState = "state"
    Public Const fldStateDet = "state_det"

    Public Const fldCategoryNo = "category_no"
    Public Const fldCategoryDet = "category_det"

    Public Const fldOthCategoryNo = "oth_category_no"
    Public Const fldOthCategoryDet = "oth_category_det"

    Public Const fldDeptNo = "dept_no"
    Public Const fldDeptDet = "dept_det"
    Public Const fldDeptNo_2 = "dept_no_2"
    Public Const fldDeptDet_2 = "dept_det_2"
    Public Const fldDeptNo_3 = "dept_no_3"
    Public Const fldDeptDet_3 = "dept_det_3"

    Public Const fldDesignationNo = "designation_no"
    Public Const fldDesignation = "designation"
    Public Const fldDesignation_det = "designation_det"

    Public Const fldDesigAndOthRelationNo = "desig_and_oth_relation_no"
    Public Const fldDesigAndOthRelationDet = "desig_and_oth_relation_det"

    Public Const fldNationalityNo = "nationality_no"
    Public Const fldNationalityDet = "nationality_det"

    Public Const fldReligionNo = "religion_no"
    Public Const fldReligionDet = "religion_det"

    Public Const fldCasteNo = "caste_no"
    Public Const fldCasteDet = "caste_det"

    Public Const fldCasteTypeNo = "caste_type_no"
    Public Const fldCasteTypeDet = "caste_type_det"

    Public Const fldSubCasteNo = "sub_caste_no"
    Public Const fldSubCasteDet = "sub_caste_det"
    Public Const fldSubCasteOthDet = "sub_caste_oth_det"

    Public Const fldBloodGroupNo = "blood_group_no"
    Public Const fldBloodGroupDet = "blood_group_det"

    Public Const fldIdentityMarksDet_1 = "identity_marks_det_1"
    Public Const fldIdentityMarksDet_2 = "identity_marks_det_2"

    Public Const fldLastInstDet = "last_inst_det"
    Public Const fldSrcOfPay = "src_of_pay"

    Public Const fldPtaxStat = "ptax_stat"
    Public Const fldInvRelStat = "inv_rel_stat"
    Public Const fldCrDrDetRelStat = "cr_dr_det_rel_stat"
    Public Const fldSalePurcRelStat = "sale_purc_rel_stat"
    Public Const fldRvRelStat = "rv_rel_stat"
    Public Const fldParentStat = "parent_stat"
    Public Const fldPartTimeStat = "part_time_stat"
    Public Const fldLoginStat = "login_stat"
    Public Const fldImpStat = "imp_stat"
    Public Const fldActiveStat = "active_stat"

    Public Const fldAccHeadNo = "acc_head_no"
    Public Const fldRelAccNo = "rel_acc_no"

    Public Const fldSalMainStaffCode = "sal_main_staff_code"
    Public Const fldSalSubStaffCode = "sal_sub_staff_code"

    Public Const fldAllInvEntType = "all_inv_ent_type"

    Public Const fldMastCodeTypeNo = "mast_code_type_no"
    Public Const fldMastCodeTypeDet = "mast_code_type_det"

    Public Const fldRelTypeNo = "rel_type_no"
    Public Const fldRelTypeDet = "rel_type_det"

    Public Const fldEntryTypeNo = "entry_type_no"
    Public Const fldEntryTypeDet = "entry_type_det"

    Public Const fldEntryOthTypeNo = "entry_oth_type_no"
    Public Const fldEntryOthTypeDet = "entry_oth_type_det"

    Public Const flNo = "type_no"
    Public Const flDet = "type_det"

    Public Const fldAmtTypeNo = "amt_type_no"
    Public Const fldAmtTypeDet = "amt_type_det"

    Public Const fldReciTypeNo = "reci_type_no"
    Public Const fldReciTypeDet = "reci_type_det"

    Public Const fldInstCount = "inst_count"

    Public Const fldAmtCalcType = "amt_calc_type"
    Public Const fldAmtCalcOnVal = "amt_calc_on_val"
    Public Const fldAmtCalcByVal = "amt_calc_by_val"

    Public Const fldAmt = "amt"
    Public Const fldAmtDet = "amt_det"

    Public Const fldAmt_2 = "amt_2"
    Public Const fldAmt_2_Det = "amt_2_det"

    Public Const fldOthAmt = "oth_amt"
    Public Const fldOthAmtDet = "oth_amt_det"
    Public Const fldOthAmtDt = "oth_amt_dt"

    Public Const fldCloseAmt = "close_amt"
    Public Const fldCloseAmtDet = "close_amt_det"

    Public Const fldImgDriveDet = "img_drive_det"
    Public Const fldImgPathDet = "img_path_det"
    Public Const fldImgDet = "img_det"

    Public Const fldDispOrder = "disp_order"

    Public Const fldCreateBy = "create_by"
    Public Const fldCreateDate = "create_date"
    Public Const fldCreateTime = "create_time"
    Public Const fldModifyBy = "modify_by"
    Public Const fldModifyDate = "modify_date"
    Public Const fldModifyTime = "modify_time"
    Public Const fldDelStat = "del_stat"

#End Region

    Public Enum EnumDbConnType
        Access
        Sql
    End Enum

    Public Enum EnumDbTableDet
        Field
    End Enum

    Public Enum EnumDbFieldDet
        Length
        Type
        Value
        DefaultValue
        ValForRelCondn
        IndexStat
        ImpStat
    End Enum

    Public Enum EnumDbColDet
        TableName
        ActType
        Type
        MaxLength
        Precision
        Scale
        IsNullable
        DefaulValue
        Value
        ValForRelCondn
    End Enum

    Public Enum EnumDbChangeType
        NewRec
        Update
        None
    End Enum

    <Serializable()>
    Public Class clsDbFieldDet

        Implements ICloneable

        Dim _Heading As Object

        Dim _Name As Object

        Dim _DataFormat As Object
        Dim _Type As Object
        Dim _Length As Object

        Dim _DbTableName As Object

        Dim _DefaDet As New clsGenData
        Dim _PrevDefaDet As New clsGenData

        Dim _Value As New clsGenData
        Dim _PrevValue As New clsGenData

        Dim _IndexStat As Boolean

        Dim _ImpStat As Boolean
        Dim _ActiveStat As Boolean

        Sub New()

        End Sub

        Public Property Heading() As Object

            Get
                Heading = _Heading
            End Get

            Set(ByVal Value As Object)
                _Heading = Value
            End Set

        End Property

        Public Property Name() As Object

            Get
                Name = _Name
            End Get

            Set(ByVal Value As Object)
                _Name = Value
            End Set

        End Property

        Public Property DataFormat() As Object

            Get
                DataFormat = _DataFormat
            End Get

            Set(ByVal Value As Object)
                _DataFormat = Value
            End Set

        End Property

        Public Property Type() As Object

            Get
                Type = _Type
            End Get

            Set(ByVal Value As Object)
                _Type = Value
            End Set

        End Property

        Public Property Length() As Object

            Get
                Length = _Length
            End Get

            Set(ByVal Value As Object)
                _Length = Value
            End Set

        End Property

        Public Property DbTableName() As Object

            Get
                DbTableName = _DbTableName
            End Get

            Set(ByVal Value As Object)
                _DbTableName = Value
            End Set

        End Property

        Public Property DefaDet() As clsGenData

            Get
                DefaDet = _DefaDet
            End Get

            Set(ByVal Value As clsGenData)
                _DefaDet = Value
            End Set

        End Property

        Public Property PrevDefaDet() As clsGenData

            Get
                PrevDefaDet = _PrevDefaDet
            End Get

            Set(ByVal Value As clsGenData)
                _PrevDefaDet = Value
            End Set

        End Property

        Public Property Value() As clsGenData

            Get
                Value = _Value
            End Get

            Set(ByVal Value As clsGenData)
                _Value = Value
            End Set

        End Property

        Public Property PrevValue() As clsGenData

            Get
                PrevValue = _PrevValue
            End Get

            Set(ByVal Value As clsGenData)
                _PrevValue = Value
            End Set

        End Property

        Public Property IndexStat() As Object

            Get
                IndexStat = _IndexStat
            End Get

            Set(ByVal Value As Object)
                _IndexStat = Value
            End Set

        End Property

        Public Property ImpStat() As Object

            Get
                ImpStat = _ImpStat
            End Get

            Set(ByVal Value As Object)
                _ImpStat = Value
            End Set

        End Property

        Public Property ActiveStat() As Object

            Get
                ActiveStat = _ActiveStat
            End Get

            Set(ByVal Value As Object)
                _ActiveStat = Value
            End Set

        End Property

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    <Serializable()>
    Public Class clsRecordset

        Implements ICloneable

        Dim _IsInitialized As Boolean

        Dim _FieldPfx As Object

        Private _Da As Object
        Private _Dset As System.Data.DataSet

        Private _CommBuilder As Object

        Private _DtableNo As Object
        Private _Dtable As System.Data.DataTable

        Private _lstKey As New List(Of System.Data.DataColumn)

        Private _RecColl As Object

        Dim _RecNo As Object

        Dim _FilterSqlDet As Object
        Dim _FilterSortOrder As Object

        Dim _Eof As Boolean

        Public Sub New()

            _IsInitialized = False

            _RecNo = -1
            _Eof = True

        End Sub

        Protected Overrides Sub Finalize()

            If (IsInitialized) Then

                '                _Da.Dispose()
                _Dset.Dispose()
                _CommBuilder.Dispose()

                _lstKey.Clear()

            End If

        End Sub

        Public ReadOnly Property IsInitialized() As Boolean

            Get
                IsInitialized = _IsInitialized
            End Get

        End Property

        Public ReadOnly Property Da() As Object

            Get
                Da = _Da
            End Get

        End Property

        Public ReadOnly Property Dset() As System.Data.DataSet

            Get
                Dset = _Dset
            End Get

        End Property

        Public ReadOnly Property CommBuilder() As Object

            Get
                CommBuilder = _CommBuilder
            End Get

        End Property

        Public ReadOnly Property Dtable() As System.Data.DataTable

            Get
                Dtable = _Dtable
            End Get

        End Property

        Public ReadOnly Property Drow() As System.Data.DataRow

            Get
                Drow = Nothing

                If (Not Eof) Then Drow = _RecColl(_RecNo)

            End Get

        End Property

        Public ReadOnly Property Bof() As Boolean

            Get
                Bof = (_RecNo = 0)
            End Get

        End Property

        Public ReadOnly Property Eof() As Boolean

            Get
                Eof = Not (_RecNo >= 0 And _RecNo <= RecCount() - 1)
            End Get

        End Property

        Public Property RecNo() As Object

            Get
                RecNo = _RecNo
            End Get

            Set(ByVal value As Object)
                _RecNo = value
            End Set

        End Property

        Public Function Open(ByVal DbConn_par As Object,
                             ByVal SqlDet_par As Object,
                             Optional ByVal FieldPfx_par As Object = "",
                             Optional ByVal KeyNameDet_par As Object = "") As Object

            Dim ret_val : ret_val = False

            If (IsInitialized) Then
                '                _Da.Dispose()
                _Dset.Dispose()
                _CommBuilder.Dispose()
            End If


            Dim ConnType = DbConn_par.GetType.ToString()
            Dim OleConnType = GetType(OleDb.OleDbConnection).ToString

#If ConnType = OleConnType Then
            Dim DbComm_lcl As New OleDb.OleDbCommand
            Dim Da_lcl As New OleDb.OleDbDataAdapter
            Dim CommBuilderLocal As New OleDb.OleDbCommandBuilder(Da_lcl)
#Else
        Dim DbComm_lcl As New SqlClient.SqlCommand
        Dim Da_lcl As New SqlClient.SqlDataAdapter
        Dim CommBuilderLocal As New SqlClient.SqlCommandBuilder(Da_lcl)
#End If
            '        Try

            SqlDet_par = Trim(SqlDet_par)
            _FieldPfx = Trim(FieldPfx_par)

            '            If (DbConn_par.State = 0 Or Len(SqlDet_par) = 0 Or _
            '                (Len(Trim(KeyNameDet_par)) = 0 And Not IsArray(KeyNameDet_par))) Then GoTo end_sub
            If (DbConn_par.State = 0 Or Len(SqlDet_par) = 0) Then GoTo end_sub

            If (InStr(SqlDet_par, "select ", CompareMethod.Text) <> 1) Then
                SqlDet_par = "select * from " & SqlDet_par
            End If


            _Da = Da_lcl
            _Dset = New System.Data.DataSet
            _DtableNo = "0"

            With DbComm_lcl
                .CommandTimeout = 100000
                .Connection = DbConn_par
                .CommandText = SqlDet_par
            End With

            With _Da
                .SelectCommand = DbComm_lcl
                .Fill(_Dset, _DtableNo)
            End With

            _CommBuilder = CommBuilderLocal

            _Dtable = _Dset.Tables(_DtableNo)
            _RecColl = _Dtable.Rows

            _FilterSqlDet = ""
            _FilterSortOrder = ""

            MoveFirst()

            ret_val = True

            '        Catch

            '        End Try

end_sub:
            _IsInitialized = ret_val

            SetKey(KeyNameDet_par)

            Open = ret_val

        End Function

        Public Function Reopen(Optional ByVal SqlDet_par As Object = "",
                               Optional ByVal FieldPfx_par As Object = "") As Object

            Dim ret_val As Object = False

            '        Try

            If (Not IsInitialized) Then GoTo end_sub

            With _Da

                _Dtable.Rows.Clear()

                _FilterSqlDet = ""
                _FilterSortOrder = ""

                With .SelectCommand

                    If (Len(Trim(SqlDet_par)) = 0) Then
                        SqlDet_par = .CommandText
                    Else
                        If (InStr(SqlDet_par, "select ", CompareMethod.Text) <> 1) Then
                            SqlDet_par = "select * from " & SqlDet_par
                            _FieldPfx = Trim(FieldPfx_par)
                        End If
                        .CommandText = SqlDet_par
                    End If

                End With

                .Fill(_Dset, _DtableNo)

                _RecColl = _Dtable.Rows

            End With

            MoveFirst()

            ret_val = True

            '        Catch

            '        End Try

end_sub:
            Reopen = ret_val

        End Function

        Public Sub Refresh()

            Dim ret_val : ret_val = False

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            Dim _PrevFilterSqlDet = _FilterSqlDet
            Dim _PrevFilterSortOrder = _FilterSortOrder

            If (Reopen()) Then

                If (Len(_PrevFilterSqlDet) < 0) Then
                    SetFilter(_PrevFilterSqlDet, _PrevFilterSortOrder)
                End If

                ret_val = True

            End If

            '            Catch

            '            End Try

end_sub:

        End Sub

        Private Function GetKey(ByVal Key_par As Object) As System.Data.DataColumn()

            Dim ret_val() As System.Data.DataColumn = {Nothing}

            '        Try

            If (Not IsInitialized) Then GoTo end_sub

            ret_val = _Dtable.PrimaryKey

            '        Catch

            '        End Try

end_sub:
            If (IsNothing(ret_val(0))) Then DelArrItem(ret_val, 0)

            GetKey = ret_val

        End Function

        Private Function SetKey(ByVal KeyNameDet_par As Object) As Object

            Dim ret_val As Object = False

            _lstKey.Clear()

            '        Try

            If (Not IsInitialized) Then GoTo end_sub


            Dim lstKeyNameDet As New List(Of Object)

            If (IsArray(KeyNameDet_par)) Then
                lstKeyNameDet.AddRange(KeyNameDet_par)
            Else
                If (Len(Trim(KeyNameDet_par)) = 0) Then
                    ret_val = True
                    GoTo end_sub
                Else
                    lstKeyNameDet.Add(KeyNameDet_par)
                End If
            End If


            With _Dtable

                For index_no = 0 To lstKeyNameDet.Count - 1

                    With lstKeyNameDet
                        .Item(index_no) = FieldName(.Item(index_no))
                    End With

                    With _lstKey
                        .Add(New System.Data.DataColumn)
                        .Item(.Count - 1) = _Dtable.Columns(lstKeyNameDet.Item(index_no))
                    End With

                Next

                ret_val = True

            End With

            '        Catch

            '        End Try

end_sub:
            With _lstKey
                If (.Count > 0) Then _Dtable.PrimaryKey = .ToArray
            End With

            SetKey = ret_val

        End Function

        Public Function SetKeyVal(ByVal KeyVal_par As Object, _
                                  Optional ByRef Drow_par As System.Data.DataRow = Nothing) As Object

            Dim ret_val As Object = False

            '        Try

            If (IsNothing(Drow_par)) Then Drow_par = Drow

            If (_lstKey.Count = 0) Then
                ret_val = True
                GoTo end_sub
            ElseIf (IsNothing(Drow_par)) Then
                GoTo end_sub
            End If


            Dim lstKeyVal As New List(Of Object)

            If (IsArray(KeyVal_par)) Then
                With lstKeyVal
                    .AddRange(KeyVal_par)
                    If (.Count <> _lstKey.Count) Then GoTo end_sub
                End With
            Else
                If (Len(Trim(KeyVal_par)) = 0) Then
                    GoTo end_sub
                Else
                    lstKeyVal.Add(KeyVal_par)
                End If
            End If


            With _lstKey

                For index_no = 0 To .Count - 1
                    SetDet(.Item(index_no).ColumnName, lstKeyVal.Item(index_no), Drow_par)
                Next

                ret_val = True

            End With

            '        Catch

            '        End Try

end_sub:
            _Dtable.PrimaryKey = _lstKey.ToArray

            SetKeyVal = ret_val

        End Function

        Public Function RecCount() As Object

            Dim ret_val As Object = 0

            '            Try

            If (IsInitialized) Then

                If (_RecColl.GetType.ToString() = GetType(System.Data.DataRowCollection).ToString) Then
                    ret_val = _RecColl.Count
                Else
                    ret_val = _RecColl.Length
                End If

            End If

            '            Catch

            '            End Try

end_sub:
            RecCount = ret_val

        End Function

        Public Sub MoveFirst()

            '            Try

            _RecNo = 0

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub MoveLast()

            '            Try

            _RecNo = RecCount() - 1

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub MoveNext()

            '            Try

            _RecNo = _RecNo + 1

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub MovePrevious()

            '            Try

            _RecNo = _RecNo - 1

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Sub MoveTo(ByVal RecNo_par As Object)

            '            Try

            If (Not IfInteger(RecNo_par)) Then RecNo_par = -1

            _RecNo = RecNo_par

            '            Catch

            '            End Try

end_sub:

        End Sub

        Public Function Find(ByVal Key_par As Object, _
                             ByVal KeyVal_par As Object) As Object

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            _RecNo = -1

            Dim lstKey As New List(Of Object)

            If (IsArray(Key_par)) Then
                lstKey.AddRange(Key_par)
            Else
                If (Len(Trim(Key_par)) > 0) Then lstKey.Add(Key_par)
            End If


            Dim lstKeyVal As New List(Of Object)

            If (IsArray(KeyVal_par)) Then
                lstKeyVal.AddRange(KeyVal_par)
            Else
                If (Len(Trim(KeyVal_par)) > 0) Then lstKeyVal.Add(KeyVal_par)
            End If


            If (Not (lstKey.Count > 0 And lstKeyVal.Count > 0 And lstKey.Count = lstKeyVal.Count)) Then
                GoTo end_sub
            End If


            If (SetKey(lstKey.ToArray)) Then

                If (_RecColl.GetType.ToString() = GetType(System.Data.DataRowCollection).ToString) Then
                    With _RecColl
                        _RecNo = .IndexOf(.Find(lstKeyVal.ToArray))
                    End With
                Else
                    Dim dtbl_lcl As System.Data.DataTable = _RecColl.CopyToDataTable()
                    With dtbl_lcl.Rows
                        _RecNo = .IndexOf(.Find(lstKeyVal.ToArray))
                    End With
                End If

            End If

            '            Catch

            '            End Try

end_sub:
            Find = Not Eof

        End Function

        Public Function FieldName(ByVal FieldName_par) As Object

            Dim ret_val : ret_val = FieldName_par

            FieldName_par = Trim(FieldName_par)

            If (Len(FieldName_par) > 0 And InStr(FieldName_par, _FieldPfx, CompareMethod.Text) = 0) Then
                FieldName_par = _FieldPfx & FieldName_par
            End If

            ret_val = FieldName_par

end_sub:
            FieldName = FieldName_par

        End Function

        Public Function PrepRelCond(ByVal FieldName_par As Object, _
                                    ByVal RelOprtr_par As Object, _
                                    ByVal FieldVal_par As Object) As Object

            Dim ret_val : ret_val = ""

            '            Try

            FieldName_par = FieldName(Trim(FieldName_par))
            RelOprtr_par = Trim(RelOprtr_par)

            Dim FieldDtype = GetDataType(_Dtable.Columns(FieldName_par))

            If (Len(FieldName_par) > 0 And Len(RelOprtr_par) > 0) Then

                If (FieldDtype = EnumDataType.Dt) Then

                    FieldVal_par = IIf(IfNull(FieldVal_par) Or Not IsDate(FieldVal_par), _
                                       DateValue("01/01/1900"), FieldVal_par)
                    FieldVal_par = Format(FieldVal_par, DispFormat_mm_dd_yyyy)

                ElseIf (Not FieldDtype = EnumDataType.Real) Then

                    FieldVal_par = IIf(IfNull(FieldVal_par) Or Len(Trim(FieldVal_par)) = 0, 0, FieldVal_par)

                ElseIf (FieldDtype = EnumDataType.Real) Then

                    FieldVal_par = IIf(IfNull(FieldVal_par) Or Len(Trim(FieldVal_par)) = 0, Val("0.0"), FieldVal_par)

                ElseIf (FieldDtype = EnumDataType.Bool) Then

                    FieldVal_par = IIf(IfNull(FieldVal_par), False, FieldVal_par)

                Else

                    FieldVal_par = IIf(IfNull(FieldVal_par), "", FieldVal_par)

                End If


                If (FieldDtype = EnumDataType.Text Or FieldDtype = EnumDataType.Dt) Then
                    FieldVal_par = "'" & FieldVal_par & "'"
                End If


                ret_val = FieldName_par & " " & RelOprtr_par & " " & FieldVal_par

            End If

            '            Catch

            '            End Try

end_sub:
            PrepRelCond = ret_val

        End Function

        Public Function SetFilter(ByVal Sql_par As Object, _
                                  Optional ByVal SortOrder_par As Object = "") As Object

            Dim ret_val : ret_val = False

            _FilterSqlDet = Trim(Sql_par)
            _FilterSortOrder = Trim(SortOrder_par)

            ret_val = ApplyFilter()

end_sub:
            SetFilter = ret_val

        End Function

        Private Function ApplyFilter() As Object

            Dim ret_val : ret_val = False

            If (Not IsInitialized Or Len(_FilterSqlDet) = 0) Then GoTo end_sub

            With _Dtable

                If (Len(_FilterSortOrder) > 0) Then
                    _RecColl = .Select(_FilterSqlDet, _FilterSortOrder)
                Else
                    _RecColl = .Select(_FilterSqlDet)
                End If

                MoveFirst()

                ret_val = True

            End With

end_sub:
            ApplyFilter = ret_val

        End Function

        Public Function ResetFilter() As Object

            Dim ret_val : ret_val = False

            If (Not IsInitialized Or Len(_FilterSqlDet) = 0) Then GoTo end_sub

            With _Dtable

                _RecColl = .Rows

                _FilterSqlDet = ""
                _FilterSortOrder = ""

                MoveFirst()

                ret_val = True

            End With

end_sub:
            ResetFilter = ret_val

        End Function

        Public Function GetDet(ByVal FieldName_par As Object, _
                               Optional ByVal Drow_par As System.Data.DataRow = Nothing, _
                               Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

            Dim ret_val As Object = Nothing

            '        Try

            If (IsNothing(Drow_par)) Then Drow_par = Drow

            If (IsNothing(Drow_par)) Then GoTo end_sub

            FieldName_par = FieldName(FieldName_par)

            Dim FieldDtype = GetDataType(DirectCast(_Dtable.Columns(FieldName_par), System.Data.DataColumn))


            If (DetType_par = EnumDbFieldDet.Value Or DetType_par = EnumDbFieldDet.ValForRelCondn) Then

                With Drow_par
                    ret_val = .Item(FieldName_par)
                End With

                If (IsNothing(ret_val)) Then
                    If (FieldDtype = EnumDataType.Text) Then
                        ret_val = ""
                    ElseIf (FieldDtype = EnumDataType.Int) Then
                        ret_val = 0
                    ElseIf (FieldDtype = EnumDataType.Real) Then
                        ret_val = 0
                    ElseIf (FieldDtype = EnumDataType.Dt) Then
                        ret_val = DateValue("01/01/1900")
                    End If
                End If

                If (DetType_par = EnumDbFieldDet.ValForRelCondn) Then
                    If (FieldDtype = EnumDataType.Text Or FieldDtype = EnumDataType.Dt) Then
                        ret_val = "'" & ret_val & "'"
                    End If
                End If


            ElseIf (DetType_par = EnumDbFieldDet.Type) Then

                ret_val = FieldDtype

            End If

            '        Catch

            '        End Try

end_sub:
            GetDet = ret_val

        End Function

        Public Function SetDet(ByVal FieldName_par As Object, _
                               ByRef FieldVal_par As Object, _
                               Optional ByRef Drow_par As System.Data.DataRow = Nothing, _
                               Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

            Dim ret_val As Object = False

            '        Try

            If (IsNothing(Drow_par)) Then Drow_par = Drow

            If (IsNothing(Drow_par)) Then GoTo end_sub

            Dim lstFieldName As New List(Of Object)

            If (IsArray(FieldName_par)) Then
                lstFieldName.AddRange(FieldName_par)
            Else
                If (Len(Trim(FieldName_par)) > 0) Then lstFieldName.Add(FieldName_par)
            End If


            Dim lstFieldVal As New List(Of Object)

            If (IsArray(FieldVal_par)) Then
                lstFieldVal.AddRange(FieldVal_par)
            Else
                If (Len(Trim(FieldVal_par)) > 0) Then lstFieldVal.Add(FieldVal_par)
            End If


            If (Not (lstFieldName.Count > 0 And lstFieldVal.Count > 0 And lstFieldName.Count = lstFieldVal.Count)) Then
                GoTo end_sub
            ElseIf (lstFieldName.Count = 0 And lstFieldVal.Count = 0) Then
                ret_val = True
                GoTo end_sub
            End If


            With Drow_par

                .BeginEdit()

                For field_no = 0 To lstFieldName.Count - 1
                    .Item(FieldName(lstFieldName.Item(field_no))) = lstFieldVal.Item(field_no)
                Next field_no

                SetNullDet(Drow_par)

                .EndEdit()

                ret_val = True

            End With

            '        Catch

            '            Drow_par.CancelEdit()

            '        End Try

end_sub:
            SetDet = ret_val

        End Function

        Private Function SetNullDet(ByVal FieldName_par As Object, _
                                    Optional ByRef Drow_par As System.Data.DataRow = Nothing) As Boolean

            Dim ret_val As Object = False

            '        Try

            If (IsNothing(Drow_par)) Then Drow_par = Drow

            If (IsNothing(Drow_par)) Then GoTo end_sub

            Dim lstFieldName As New List(Of Object)

            If (IsArray(FieldName_par)) Then
                lstFieldName.AddRange(FieldName_par)
            Else
                If (Len(Trim(FieldName_par)) > 0) Then lstFieldName.Add(FieldName_par)
            End If


            With Drow_par

                Dim Field_lcl As System.Data.DataColumn

                For field_no = 0 To lstFieldName.Count - 1

                    Field_lcl = _Dtable.Columns(FieldName(lstFieldName.Item(field_no)))

                    Dim field_name = Field_lcl.ColumnName
                    Dim field_type = GetDet(field_name, Drow_par, EnumDbFieldDet.Type)
                    Dim field_val = Drow_par.Item(field_no)

                    If (.IsNull(field_name)) Then

                        If (field_type = EnumDataType.Text) Then
                            field_val = ""
                        ElseIf (field_type = EnumDataType.Int) Then
                            field_val = 0
                        ElseIf (field_type = EnumDataType.Real) Then
                            field_val = 0
                        ElseIf (field_type = EnumDataType.Dt) Then
                            field_val = DateValue("01/01/1900")
                        End If

                        .Item(field_name) = field_val

                    End If

                Next field_no

                ret_val = True

            End With

            '        Catch

            '        End Try

end_sub:
            SetNullDet = ret_val

        End Function

        Private Function SetNullDet(Optional ByRef Drow_par As System.Data.DataRow = Nothing) As Boolean

            Dim ret_val As Object = False

            '        Try

            If (IsNothing(Drow_par)) Then Drow_par = Drow

            If (IsNothing(Drow_par)) Then GoTo end_sub


            With Drow_par

                Dim Field_lcl As System.Data.DataColumn

                For field_no = 0 To _Dtable.Columns.Count - 1

                    Field_lcl = _Dtable.Columns(field_no)

                    Dim field_name = Field_lcl.ColumnName
                    Dim field_type = GetDet(field_name, Drow_par, EnumDbFieldDet.Type)
                    Dim field_val = Drow_par.Item(field_no)

                    If (.IsNull(field_name)) Then ' Checks whether a column contains NULL value

                        If (field_type = EnumDataType.Text) Then
                            field_val = ""
                        ElseIf (field_type = EnumDataType.Int) Then
                            field_val = 0
                        ElseIf (field_type = EnumDataType.Real) Then
                            field_val = 0
                        ElseIf (field_type = EnumDataType.Dt) Then
                            field_val = DateValue("01/01/1900")
                        End If

                        .Item(field_name) = field_val

                    End If

                Next field_no

                ret_val = True

            End With

            '        Catch

            '        End Try

end_sub:
            SetNullDet = ret_val

        End Function

        Public Function Add(ByVal FieldName_par As Object, _
                            ByVal FieldVal_par As Object) As Object

            Dim ret_val : ret_val = False

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            _RecNo = -1

            Dim DrowNew As System.Data.DataRow = _Dtable.NewRow

            If (Not SetDet(FieldName_par, FieldVal_par, DrowNew)) Then GoTo end_sub

            SetNullDet(DrowNew)

            _Dtable.Rows.Add(DrowNew)

            If (_RecColl.GetType.ToString() <> GetType(System.Data.DataRowCollection).ToString) Then
                AddArrItem(_RecColl, DrowNew)
            End If

            MoveLast()

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            Add = Not Eof

        End Function

        Public Function Delete() As Object

            Dim ret_val : ret_val = False

            '            Try

            If (Eof) Then GoTo end_sub

            Drow.Delete()

            If (_RecColl.GetType.ToString() <> GetType(System.Data.DataRowCollection).ToString) Then
                DelArrItem(_RecColl, _RecNo)
            End If

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            Delete = Not Eof

        End Function

        Public Function Update(Optional ByVal Msg_par As Boolean = False) As Object

            Dim ret_val As Object = False

            Dim mpointer_lcl = Cursor.Current

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            If (Msg_par = True) Then

                Cursor.Current = Cursors.Default

                If (MsgBox("Are You Sure ?", vbQuestion + vbYesNo + vbDefaultButton1, "SAVE  CHANGES") = vbNo) Then
                    GoTo end_sub
                End If

                Cursor.Current = mpointer_lcl

            End If


            With _Da
                .Update(_Dtable)
                ret_val = True
            End With

            '            ret_val = UpdateReqdRec(System.Data.DataViewRowState.Added)
            '            If (Not ret_val) Then GoTo end_sub

            '            ret_val = UpdateReqdRec(System.Data.DataViewRowState.ModifiedOriginal)
            '            If (Not ret_val) Then GoTo end_sub

            '            ret_val = UpdateReqdRec(System.Data.DataViewRowState.ModifiedCurrent)
            '            If (Not ret_val) Then GoTo end_sub

            '            ret_val = UpdateReqdRec(System.Data.DataViewRowState.Deleted)
            '            If (Not ret_val) Then GoTo end_sub

            '            Catch

            '            End Try

end_sub:
            Update = ret_val

        End Function

        Private Function UpdateReqdRec(ByVal DataType_par As System.Data.DataViewRowState) As Object

            Dim ret_val As Object = True

            If (_lstKey.Count = 0) Then
                ret_val = False
                GoTo end_sub
            End If

            '            Try

            With _Dtable

                Dim dtblChanges As System.Data.DataTable = .GetChanges(DataType_par)

                If (Not IsNothing(dtblChanges)) Then

                    If (Not dtblChanges.HasErrors) Then
                        _Da.Update(dtblChanges)
                    Else
                        ret_val = False
                    End If

                End If

            End With

            '            Catch

            '            ret_val = False

            '            End Try

end_sub:
            If (Not ret_val) Then
                Rollback()
            Else
                _Dtable.AcceptChanges()
            End If

            UpdateReqdRec = ret_val

        End Function

        Public Function Rollback(Optional ByVal Msg_par As Boolean = False) As Object

            Dim ret_val As Object = False

            '            Try

            If (Not IsInitialized) Then GoTo end_sub

            Refresh()

            ret_val = True

            '            Catch

            '            End Try

end_sub:
            Rollback = ret_val

        End Function

        Public Function Clone() As Object Implements System.ICloneable.Clone

            Dim mem_stream As New MemoryStream()
            Dim bin_formatter As New BinaryFormatter()

            bin_formatter.Serialize(mem_stream, Me)
            mem_stream.Seek(0, SeekOrigin.Begin)

            Return bin_formatter.Deserialize(mem_stream)

        End Function

    End Class

    Public Function ActionOnDb(ByVal DbConn_par As Object, _
                               ByVal SqlDet_par As Object) As Object

        Dim ret_val : ret_val = False

        '        Try

        Dim ConnType = DbConn_par.GetType.ToString()
        Dim OleConnType = GetType(OleDb.OleDbConnection).ToString

#If ConnType = OleConnType Then
        Dim DbComm_lcl As New OleDb.OleDbCommand
#Else
        Dim DbComm_lcl As New SqlClient.SqlCommand
#End If

        SqlDet_par = Trim(SqlDet_par)

        If (DbConn_par.State = 0 Or Len(SqlDet_par) = 0) Then GoTo end_sub

        With DbComm_lcl
            .CommandTimeout = 100000
            .Connection = DbConn_par
            .CommandText = SqlDet_par
            .ExecuteNonQuery()
        End With

        ret_val = True

        '        Catch

        '        End Try

end_sub:

        DbComm_lcl.Dispose()

        ActionOnDb = ret_val

    End Function

    Public Function CreateDataSet(ByVal DbConn_par As Object, _
                                  ByVal SqlDet_par As Object) As System.Data.DataSet

        Dim ConnType = DbConn_par.GetType.ToString()
        Dim OleConnType = GetType(OleDb.OleDbConnection).ToString

#If ConnType = OleConnType Then
        Dim DbComm_lcl As New OleDb.OleDbCommand
        Dim Da_lcl As New OleDb.OleDbDataAdapter
#Else
        Dim DbComm_lcl As New SqlClient.SqlCommand
        Dim Da_lcl As New SqlClient.SqlDataAdapter
#End If

        Dim Ds As New System.Data.DataSet

        '        Try

        SqlDet_par = Trim(SqlDet_par)

        If (DbConn_par.State = 0 Or Len(SqlDet_par) = 0) Then GoTo end_sub

        With DbComm_lcl
            .CommandTimeout = 100000
            .Connection = DbConn_par
            .CommandText = SqlDet_par
        End With

        With Da_lcl
            .SelectCommand = DbComm_lcl
            .Fill(Ds)
        End With

        '        Catch

        '        End Try

end_sub:

        If (Not IsNothing(Da_lcl)) Then Da_lcl.Dispose()
        If (Not IsNothing(DbComm_lcl)) Then DbComm_lcl.Dispose()

        CreateDataSet = Ds

    End Function

    Public Function InitDbField(ByVal LstDbField_par As List(Of clsDbFieldDet)) As Object

        Dim ret_val As Object = False

        '        Try

        If (IsNothing(LstDbField_par)) Then GoTo end_sub


        With LstDbField_par

            For ele_no = 0 To .Count - 1

                With .Item(ele_no)
                    .Value = New clsGenData
                    .PrevValue = New clsGenData
                End With

            Next ele_no

            ret_val = True

        End With

        '        Catch

        '        End Try

end_sub:

        InitDbField = ret_val

    End Function

    Public Function GetDbField(ByVal LstDbField_par As List(Of clsDbFieldDet), _
                               ByVal FieldName_par As Object, _
                               Optional ByVal DetType_par As EnumDbFieldDet = Nothing) As Object

        Dim ret_val As Object = Nothing

        '        Try

        Dim FindIndex_lcl As Object = FindDbField(LstDbField_par, FieldName_par)

        If (FindIndex_lcl >= 0) Then

            With LstDbField_par

                If (IsNothing(DetType_par)) Then
                    ret_val = .Item(FindIndex_lcl)
                Else
                    ret_val = GetDbField(.Item(FindIndex_lcl), DetType_par)
                End If

            End With

        End If

        '        Catch

        '        End Try

end_sub:

        GetDbField = ret_val

    End Function

    Public Function GetDbField(ByVal DbField_par As clsDbFieldDet, _
                               Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

        Dim ret_val As Object = Nothing

        '        Try

        If (Not IsNothing(DbField_par)) Then

            With DbField_par

                If (DetType_par = EnumDbFieldDet.Value) Then
                    ret_val = .Value.Text
                ElseIf (DetType_par = EnumDbFieldDet.DefaultValue) Then
                    ret_val = .DefaDet.Text
                ElseIf (DetType_par = EnumDbFieldDet.Type) Then
                    ret_val = .Type
                ElseIf (DetType_par = EnumDbFieldDet.Length) Then
                    ret_val = .Length
                ElseIf (DetType_par = EnumDbFieldDet.ImpStat) Then
                    ret_val = .ImpStat
                End If

            End With

        End If

        '        Catch

        '        End Try

end_sub:

        GetDbField = ret_val

    End Function

    Public Function SetDbField(ByVal LstDbField_par As List(Of clsDbFieldDet), _
                               ByVal FieldName_par As Object, _
                               ByVal DetVal_par As Object, _
                               Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

        Dim ret_val As Object = False

        '        Try

        Dim FindIndex_lcl As Object = FindDbField(LstDbField_par, FieldName_par)

        If (FindIndex_lcl >= 0) Then
            ret_val = SetDbField(LstDbField_par.Item(FindIndex_lcl), DetVal_par, DetType_par)
        End If

        '        Catch

        '        End Try

end_sub:

        SetDbField = ret_val

    End Function

    Public Function SetDbField(ByVal DbField_par As clsDbFieldDet, _
                               ByVal DetVal_par As Object, _
                               Optional ByVal DetType_par As EnumDbFieldDet = EnumDbFieldDet.Value) As Object

        Dim ret_val As Object = False

        '        Try

        If (Not IsNothing(DbField_par)) Then

            With DbField_par

                If (DetType_par = EnumDbFieldDet.Value) Then
                    .Value.Text = DetVal_par
                ElseIf (DetType_par = EnumDbFieldDet.DefaultValue) Then
                    .DefaDet.Text = DetVal_par
                ElseIf (DetType_par = EnumDbFieldDet.Type) Then
                    .Type = DetVal_par
                ElseIf (DetType_par = EnumDbFieldDet.Length) Then
                    .Length = DetVal_par
                ElseIf (DetType_par = EnumDbFieldDet.ImpStat) Then
                    .ImpStat = DetVal_par
                End If

            End With

            ret_val = True

        End If

        '        Catch

        '        End Try

end_sub:

        SetDbField = ret_val

    End Function

    Public Function FindDbField(ByVal LstDbField_par As List(Of clsDbFieldDet), _
                                ByVal FieldName_par As Object) As Object

        Dim ret_val As Object = -1

        '        Try

        If (IsNothing(LstDbField_par) Or IfNull(FieldName_par) Or Len(Trim(FieldName_par)) = 0) Then GoTo end_sub


        With LstDbField_par

            For ele_no = 0 To .Count - 1

                With .Item(ele_no)

                    If (UCase(.Name) = UCase(FieldName_par)) Then

                        ret_val = ele_no

                        Exit For

                    End If


                End With

            Next ele_no

        End With

        '        Catch

        '        End Try

end_sub:

        FindDbField = ret_val

    End Function

End Module