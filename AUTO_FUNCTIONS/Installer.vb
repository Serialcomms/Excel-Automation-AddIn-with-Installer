Imports Microsoft.Win32
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.Remoting.Contexts

<ComVisible(False)>
<ClassInterface(ClassInterfaceType.None)>
Public Class Installer

    Public ASSEMBLY_TYPE As Type
    Public ASSEMBLY_GUID As String
    Public ASSEMBLY_NAME As Assembly
    Public ASSEMBLY_KEYNAME As String
    Public ASSEMBLY_DISPLAY_NAME As String
    Public REGISTRATION_SERVICES As RegistrationServices

    Public Sub New()

        MyBase.New()

        InitializeComponent()       'This call is required by the Component Designer.

        'Add initialization code after the call to InitializeComponent

        ASSEMBLY_TYPE = GetType(Functions)
        ASSEMBLY_GUID = Functions.ClassId.ToUpper
        ASSEMBLY_NAME = GetType(Functions).Assembly
        ASSEMBLY_KEYNAME = "CLSID\{" & ASSEMBLY_GUID & "}\"
        ASSEMBLY_DISPLAY_NAME = ASSEMBLY_NAME.FullName
        REGISTRATION_SERVICES = New RegistrationServices

    End Sub

    Public Overrides Sub Install(stateSaver As IDictionary)

        MyBase.Install(stateSaver)

        REGISTER_ASSEMBLY()

        UPDATE_REGISTRY_KEYS()

    End Sub

    Public Overrides Sub Commit(savedState As IDictionary)

        MyBase.Commit(savedState)

    End Sub

    Public Overrides Sub Rollback(savedState As IDictionary)

        MyBase.Rollback(savedState)

    End Sub

    Public Overrides Sub Uninstall(savedState As IDictionary)

        MyBase.Uninstall(savedState)

        REMOVE_REGISTRY_KEY()

        UNREGISTER_ASSEMBLY()

    End Sub

    Public Sub REGISTER_ASSEMBLY()

        Try

            REGISTRATION_SERVICES.RegisterAssembly(ASSEMBLY_NAME, AssemblyRegistrationFlags.SetCodeBase)

        Catch ex As Exception

            Dim BOXTEXT As String = Nothing
            BOXTEXT &= "Assembly Name = " & ASSEMBLY_DISPLAY_NAME & vbCrLf
            BOXTEXT &= "Assembly GUID = " & ASSEMBLY_GUID & vbCrLf
            BOXTEXT &= "Message = " & ex.Message

            MsgBox(BOXTEXT, vbCritical + vbOKOnly, " Error Registering Assembly")

        End Try

    End Sub

    Public Sub UPDATE_REGISTRY_KEYS()

        Dim SetKey As RegistryKey = Nothing

        Try

            Registry.ClassesRoot.CreateSubKey(ASSEMBLY_KEYNAME & "Programmable")

            SetKey = Registry.ClassesRoot.OpenSubKey(ASSEMBLY_KEYNAME & "InprocServer32", True)

            SetKey.SetValue("", Environment.SystemDirectory & "\mscoree.dll", RegistryValueKind.String)

            ' Environment.SystemDirectory = auto 32/64 bit 

            Registry.ClassesRoot.Close()

        Catch ex As Exception

            Dim BOXTEXT As String = Nothing
            BOXTEXT &= "Assembly Name = " & ASSEMBLY_DISPLAY_NAME & vbCrLf
            BOXTEXT &= "Assembly GUID = " & ASSEMBLY_GUID & vbCrLf
            BOXTEXT &= "Registry Key = " & SetKey.ToString & vbCrLf
            BOXTEXT &= "Message = " & ex.Message

            MsgBox(BOXTEXT, vbCritical + vbOKOnly, " Error updating registry")

        End Try


    End Sub

    Public Sub REMOVE_REGISTRY_KEY()

        Try

            Registry.ClassesRoot.DeleteSubKey(ASSEMBLY_KEYNAME & "Programmable")

            Registry.ClassesRoot.Flush()

        Catch ex As Exception

            Dim BOXTEXT As String = Nothing
            BOXTEXT &= "Assembly Name = " & ASSEMBLY_DISPLAY_NAME & vbCrLf
            BOXTEXT &= "Assembly GUID = " & ASSEMBLY_GUID & vbCrLf
            BOXTEXT &= "Registry SubKey = " & ASSEMBLY_KEYNAME & "Programmable"
            BOXTEXT &= "Message = " & ex.Message

            MsgBox(BOXTEXT, vbCritical + vbOKOnly, " Error deleting subkey")

        End Try

    End Sub

    Public Sub UNREGISTER_ASSEMBLY()

        Try

            REGISTRATION_SERVICES.UnregisterAssembly(ASSEMBLY_NAME)

        Catch ex As Exception

            Dim BOXTEXT As String = Nothing
            BOXTEXT &= "Assembly Name = " & ASSEMBLY_DISPLAY_NAME & vbCrLf
            BOXTEXT &= "Assembly GUID = " & ASSEMBLY_GUID & vbCrLf
            BOXTEXT &= "Message = " & ex.Message

            MsgBox(BOXTEXT, vbCritical + vbOKOnly, " Error UnRegistering Assembly")

        End Try

    End Sub

End Class
