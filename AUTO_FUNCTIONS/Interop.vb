Imports Extensibility
Imports Microsoft.Office.Interop

Partial Public Class Functions
    Implements IDTExtensibility2

    Private EXCEL As Excel.Application

    Public Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As Array) Implements IDTExtensibility2.OnConnection

        If TypeOf Application Is Excel.Application Then

            EXCEL = Application

        Else

            Throw New NotImplementedException()

        End If

    End Sub
    Public Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As Array) Implements IDTExtensibility2.OnDisconnection

        Throw New NotImplementedException()

    End Sub

    Public Sub OnAddInsUpdate(ByRef custom As Array) Implements IDTExtensibility2.OnAddInsUpdate

        Throw New NotImplementedException()

    End Sub

    Public Sub OnStartupComplete(ByRef custom As Array) Implements IDTExtensibility2.OnStartupComplete

        Throw New NotImplementedException()

    End Sub

    Public Sub OnBeginShutdown(ByRef custom As Array) Implements IDTExtensibility2.OnBeginShutdown

        Throw New NotImplementedException()

    End Sub

End Class

