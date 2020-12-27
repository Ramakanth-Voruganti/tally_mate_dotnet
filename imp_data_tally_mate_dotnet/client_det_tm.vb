

Imports System.Data

Public Class clsFirmDet

    Public _ActiveStat As Boolean

    Public _DealerDetNo As Long
    Public _DealerLoc As String

    Public _No As Integer
    Public _Code As String
    Public _OthCode As String
    Public _Name As String
    Public _ShortName As String
    Public _Addr_1 As String
    Public _Addr_2 As String
    Public _Phone_1 As String

    Public _Caption As String
    Public _CaptionShort As String

    Public _DbDsn As String

    Public _DbPathMulti As String
    Public _DbPath As String
    Public _DbName As String

    Public _UserDbPath As String
    Public _UserDbName As String

    Public _TempFile As String

    Public _MainDbOleConn As New OleDb.OleDbConnection
    Public _UserDbOleConn As New OleDb.OleDbConnection

    Sub New()

        _ActiveStat = True

    End Sub

    Public Function OpenMainDbOleConn() As Boolean

        Dim ret_val = False

        '        Try

        With _MainDbOleConn

            .ConnectionString = "Provider=microsoft.Jet.oledb.4.0;" &
                                "Data Source=" & CombinePaths(Me._DbPath, Me._DbName) &
                                ";user id=admin;password=;"
            .Open()

        End With

        '        Catch

        '        End Try

end_sub:

        OpenMainDbOleConn = ret_val

    End Function

    Private Function CloseMainDbOleConn() As Boolean

        Dim ret_val = False

        '        Try

        With _MainDbOleConn

            If (.State = 1) Then
                .Close()
            End If

        End With

        '        Catch

        '        End Try

end_sub:

        CloseMainDbOleConn = ret_val

    End Function

    Private Function CloseUserDbOleConn() As Boolean

        Dim ret_val = False

        '        Try

        With _MainDbOleConn

            If (.State = 1) Then
                .Close()
            End If

        End With

        '        Catch

        '        End Try

end_sub:

        CloseUserDbOleConn = ret_val

    End Function

    Public Function IfMainDbOleConnOpen() As Boolean

        Dim ret_val As Boolean

        Try

            If (_MainDbOleConn.State = 1) Then
                ret_val = True
            End If

        Catch

        End Try

end_sub:

        IfMainDbOleConnOpen = ret_val

    End Function

    Public Function IfUserDbOleConnOpen() As Boolean

        Dim ret_val As Boolean

        Try

            If (_UserDbOleConn.State = 1) Then
                ret_val = True
            End If

        Catch

        End Try

end_sub:

        IfUserDbOleConnOpen = ret_val

    End Function

    Public Sub Dispose()

        CloseMainDbOleConn()
        CloseUserDbOleConn()

    End Sub

End Class

