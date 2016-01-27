Public Class Form1

    Public Favorites As New Internet_Favorites.Favorites()
    Private blnLoadCalled As Boolean = False

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        Favorites.ScanFavorites()

        Dim objMenu As New ContextMenu()

        For Each objWebFavorite As Internet_Favorites.WebFavorite In Favorites.FavoritesCollection
            Dim objItem As New WebFavoriteMenuItem(objWebFavorite)
            objMenu.MenuItems.Add(objItem)
        Next

        objMenu.MenuItems.Add("-")
        objMenu.MenuItems.Add(New ExitMenuItem())
        icnNotify.ContextMenu = objMenu

        blnLoadCalled = True
        Me.Hide()
    End Sub

    Private Sub Form1_VisibleChanged(sender As Object, e As EventArgs) Handles MyBase.VisibleChanged
        If blnLoadCalled = False Then
            Return
        End If
        If Me.Visible = True Then Me.Visible = False

    End Sub



End Class
