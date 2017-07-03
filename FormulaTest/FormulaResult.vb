Imports System.Collections.ObjectModel

Public Class FormulaResult

    Property PosicionFormula As PositionCell

    ReadOnly Property listaCeldasUtilizadas As ReadOnlyCollection(Of PositionCell)
        Get

            Return New ReadOnlyCollection(Of PositionCell)(listaCeldas)
        End Get
    End Property

    Private Property listaCeldas As Collection(Of PositionCell)

    Property Resultado As Double

    Sub New()
        listaCeldas = New Collection(Of PositionCell)
        PosicionFormula = New PositionCell


    End Sub


    Public Sub AgregaNuevoItemALista(ByVal objetoNuevo As Object)
        listaCeldas.Add(objetoNuevo)
    End Sub

    Public Sub AgregaListaCompleta(ByVal objetoNuevo As Collection(Of PositionCell))
        listaCeldas = objetoNuevo
    End Sub


End Class
