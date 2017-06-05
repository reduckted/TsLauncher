Imports System.Runtime.InteropServices
Imports System.Text


Friend NotInheritable Class NativeMethods

    Private Sub New()
        ' Prevents instance creation.
    End Sub


    Public Const S_OK = &H0
    Public Const S_FALSE As Integer = &H1


    <Flags()>
    Public Enum AssocF
        None = 0
        Init_NoRemapCLSID = &H1
        Init_ByExeName = &H2
        Open_ByExeName = &H2
        Init_DefaultToStar = &H4
        Init_DefaultToFolder = &H8
        NoUserSettings = &H10
        NoTruncate = &H20
        Verify = &H40
        RemapRunDll = &H80
        NoFixUps = &H100
        IgnoreBaseClass = &H200
        Init_IgnoreUnknown = &H400
        Init_FixedProgId = &H800
        IsProtocol = &H1000
        InitForFile = &H2000
    End Enum


    Public Enum AssocStr
        Command = 1
        Executable
        FriendlyDocName
        FriendlyAppName
        NoOpen
        ShellNewValue
        DDECommand
        DDEIfExec
        DDEApplication
        DDETopic
        InfoTip
        QuickTip
        TileInfo
        ContentType
        DefaultIcon
        ShellExtension
        DropTarget
        DelegateExecute
        SupportedUriProtocols
        Max
    End Enum


    <DllImport("Shlwapi.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function AssocQueryString(
            flags As AssocF,
            str As AssocStr,
            pszAssoc As String,
            pszExtra As String,
            pszOut As StringBuilder,
            ByRef pcchOut As UInteger
        ) As UInteger

    End Function

End Class
