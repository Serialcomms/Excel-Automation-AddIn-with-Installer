Option Strict On
Option Explicit On
Option Compare Text
Imports System.Runtime.InteropServices

<ComVisible(True)>
<Guid(Functions.InterfaceId)>
Public Interface ICLASS_FUNCTIONS

    Function IFX() As String

    Function TIMENOW() As String

End Interface

<ComVisible(True)>
<Guid(Functions.ClassId)>
<ProgId("AUTOMATION.Functions")>
<ClassInterface(ClassInterfaceType.None)>
<ComDefaultInterface(GetType(ICLASS_FUNCTIONS))>
Public Class Functions
    Implements ICLASS_FUNCTIONS

#Region "COM GUIDs"
    ' These GUIDs provide the COM identity for this class and its COM interfaces.
    ' If you change them, existing clients will no longer be able to access the class.

    Public Const ClassId As String = "df32904d-DEAD-BEEF-9352-5dc0ceafa03e"
    Public Const EventsId As String = "ca59b115-DEAD-BEEF-8375-a8f652d3258f"
    Public Const InterfaceId As String = "6c409593-DEAD-BEEF-8c16-f3804714e174"
#End Region

    ' A creatable COM class must have a Public Sub New() with no parameters, otherwise, the 
    ' class will not be registered in the COM registry and cannot be created via CreateObject.
    Public Sub New()

        MyBase.New()

    End Sub

    Private Function IFX() As String Implements ICLASS_FUNCTIONS.IFX

        Return "AUTO FX OK"

    End Function

    Private Function TIMENOW() As String Implements ICLASS_FUNCTIONS.TIMENOW

        EXCEL.Volatile()

        Dim DT As Date

        Dim TimeString As String

        DT = Now

        TimeString = DT.ToLongTimeString.ToString() & "." & DT.Millisecond.ToString("D3")

        Return TimeString

    End Function

End Class


