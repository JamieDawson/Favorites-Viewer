

Public Class WebFavoriteMenuItem
    Inherits MenuItem

    Public Favorite As WebFavorite

    Public Sub New(ByVal newFavorite As WebFavorite)
        Favorite = newFavorite
        Text = Favorite.Name
    End Sub

    Private Sub WebFavoriteMenuItem_Click(sender As Object, e As EventArgs) Handles Me.Click
        If Not Favorite Is Nothing Then
            Process.Start(Favorite.URL)
        End If
    End Sub
End Class


Public Class ExitMenuItem
    Inherits MenuItem
    Public Sub New()
        Text = "Exit"
    End Sub

    Private Sub ExitMenuItem_Click(sender As Object, e As EventArgs) Handles Me.Click
        Application.Exit()
    End Sub
End Class