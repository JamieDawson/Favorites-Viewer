Public Class WebFavoriteCollection
    Inherits CollectionBase

    Public Sub Add(ByVal Favorite As WebFavorite)
        List.Add(Favorite)
    End Sub

    Public Sub Remove(ByVal Index As Integer)
        If Index >= 0 And Index < Count Then
            List.RemoveAt(Index)

        End If
    End Sub

    Public ReadOnly Property item(ByVal Index As Integer) As WebFavorite
        Get
            Return CType(List.Item(Index), WebFavorite)
        End Get
    End Property
End Class

