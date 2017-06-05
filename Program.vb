Imports System.IO
Imports System.Text


Public Class Program

    Private Enum FileType
        TypeScript
        TransportStream
    End Enum


    Public Shared Sub Main(args() As String)
        If args.Length = 1 Then
            Dim fileName As String


            fileName = args(0)

            If Path.GetExtension(fileName).Equals(".ts", StringComparison.OrdinalIgnoreCase) Then
                If File.Exists(fileName) Then
                    Try
                        OpenFile(fileName)

                    Catch ex As Exception
                        MessageBox.Show(
                            "Failed to open TS file: " & Environment.NewLine & ex.Message,
                            My.Application.Info.Title,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Exclamation
                        )
                    End Try
                End If
            End If
        End If
    End Sub


    Private Shared Sub OpenFile(fileName As String)
        Dim type As FileType


        type = GetFileType(fileName)

        Select Case type
            Case FileType.TypeScript
                ' Open TypeScript files with the same
                ' thing that opens JavaScript files.
                OpenFileAsType(fileName, ".js")

            Case FileType.TransportStream
                ' Open Transport Streams with the 
                ' same thing that opens MP4 files.
                OpenFileAsType(fileName, ".mp4")

        End Select
    End Sub


    Private Shared Function GetFileType(fileName As String) As FileType
        Dim bytes() As Byte
        Dim length As Integer


        ' Read the first kilobyte of data from the file.
        bytes = New Byte(1023) {}

        Using stream As New FileStream(fileName, FileMode.Open, FileAccess.Read)
            length = stream.Read(bytes, 0, bytes.Length)
        End Using

        ' If the file is empty, then it
        ' can't be a Transport stream.
        If length = 0 Then
            Return FileType.TypeScript
        End If

        ' Transport stream packet headers start
        ' with &H47 (which is also a capital "G").
        If bytes(0) <> &H47 Then
            Return FileType.TypeScript
        End If

        ' If there are any characters below a horizontal tab (ASCII
        ' value of 9), then we will assume that it's a transport stream.
        If bytes.Take(length).Any(Function(x) x < 9) Then
            Return FileType.TransportStream
        End If

        ' The file starts with a capital G, but doesn't have any low ASCII characters,
        ' so it's probably a TypeScript file. However, we didn't look through the entire
        ' file for the low ASCII values, so it might be that we just didn't see any in
        ' the sample that we took. TypeScript files should be small, whereas Transport
        ' Streams should be large. If the file is at least 5MB in size, then we will
        ' call it a Transport Stream; otherwise, we can call it a TypeScript file.
        If (New FileInfo(fileName)).Length >= 5242880L Then
            Return FileType.TransportStream
        End If

        Return FileType.TypeScript
    End Function


    Private Shared Sub OpenFileAsType(
            fileName As String,
            extension As String
        )

        Dim command As String


        command = GetCommandForFileExtension(extension)

        If command IsNot Nothing Then
            Shell(command.Replace("%1", fileName))
        Else
            Throw New Exception($"Could not file a file association for '{extension}' files.")
        End If
    End Sub


    Private Shared Function GetCommandForFileExtension(extension As String) As String
        Dim length As UInteger
        Dim buffer As StringBuilder


        length = 0
        buffer = Nothing

        Do
            Dim result As UInteger


            result = NativeMethods.AssocQueryString(
                NativeMethods.AssocF.None,
                NativeMethods.AssocStr.Command,
                extension,
                Nothing,
                buffer,
                length
            )

            If buffer IsNot Nothing Then
                ' When we supply a buffer, the result
                ' should be `S_OK` if everything worked.
                If result = NativeMethods.S_OK Then
                    Return buffer.ToString()
                Else
                    Return Nothing
                End If

            Else
                ' `S_FALSE` should be returned when there isn't a buffer. The
                ' required buffer size will be stored in the `length` argument.
                If (result <> NativeMethods.S_FALSE) OrElse (length = 0) Then
                    Return Nothing
                End If

                buffer = New StringBuilder(CInt(length))
            End If
        Loop
    End Function

End Class
