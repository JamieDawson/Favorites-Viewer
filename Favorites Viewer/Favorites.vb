Public Class Favorites
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


    Public FavoritesCollection As WebFavoriteCollection

    Public ReadOnly Property FavoritesFolder() As String
        Get
            Return Environment.GetFolderPath(Environment.SpecialFolder.Favorites)
        End Get
    End Property

    Public Sub ScanFavorites()
        ScanFavorites(FavoritesFolder)
    End Sub

    Public Sub ScanFavorites(ByVal FolderName As String)
        If FavoritesCollection Is Nothing Then
            FavoritesCollection = New WebFavoriteCollection
        End If

        For Each strFile As String In My.Computer.FileSystem.GetFiles(FolderName)
            If strFile.EndsWith(".url", True, Nothing) Then
                Try
                    Using objWebFavorite As New WebFavorite
                        objWebFavorite.Load(strFile)
                        FavoritesCollection.Add(objWebFavorite)
                    End Using
                Catch ex As Exception
                    MessageBox.Show(ex.Message)
                End Try
            End If
        Next
    End Sub
End Class
