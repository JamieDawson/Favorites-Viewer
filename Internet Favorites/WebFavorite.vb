Imports System.IO
Imports System.Windows.Forms
Public Class WebFavorite
    Implements IDisposable

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region

    'public variables 
    Public Name As String
    Public URL As String

    Public Sub Load(ByVal FileName As String)
        Dim strData As String
        Dim strLines() As String
        Dim strLine As String

        Dim objFileInfo As New FileInfo(FileName)
        Name = objFileInfo.Name.Substring(0, objFileInfo.Name.Length - objFileInfo.Extension.Length)

        Try
            strData = My.Computer.FileSystem.ReadAllText(FileName)
            strLines = strData.Split(New String() {ControlChars.CrLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each strLine In strLines
                If strLine.StartsWith("URL=") Then
                    URL = strLine.Substring(4)
                    Exit For
                End If
            Next
        Catch ex As IOException
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
