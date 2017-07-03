

Imports System.Collections.ObjectModel
Imports System.Text
Imports FarPoint.CalcEngine

<Serializable()>
Public Class SUMIFS
    Inherits FarPoint.CalcEngine.SumIfsFunctionInfo


    Private Property PositionOriginCell As PositionCell
    Public Sub New()
        PositionOriginCell = New PositionCell
    End Sub

    Public Overrides ReadOnly Property Name As String
        Get
            Return MyBase.Name
        End Get
    End Property

    Public Overrides Function AcceptsError(i As Integer) As Boolean
        Return MyBase.AcceptsError(i)
    End Function

    Public Overrides Function AcceptsReferenceReturn() As Boolean
        Return MyBase.AcceptsReferenceReturn()
    End Function


    Public Overrides ReadOnly Property MaxArgs As Integer
        Get
            Return MyBase.MaxArgs
        End Get
    End Property

    Public Overrides ReadOnly Property MinArgs As Integer
        Get
            Return MyBase.MinArgs
        End Get
    End Property


    Public Overrides ReadOnly Property IsContextSensitive As Boolean
        Get
            Return True
        End Get
    End Property


    Public Overrides Function Evaluate(args() As Object, context As Object) As Object
        Dim reference As FarPoint.CalcEngine.CalcReference = context
        PositionOriginCell.Column = reference.Column
        PositionOriginCell.Row = reference.Row

        Return MyBase.Evaluate(args, context)
    End Function

    Public Overrides Function Evaluate(args() As Object) As Object
        Dim resultadoDeFormula As New FormulaResult
        Try
            If args Is Nothing Then
                Return FarPoint.CalcEngine.CalcError.Value
            End If
            Dim rangoSuma As CalcReference = args(0)

            Dim ElementosPorVerificarTodosCriterios = CreaListaDeRangoYCriterioParaVerificar(args)

            If rangoSuma IsNot Nothing AndAlso ElementosPorVerificarTodosCriterios IsNot Nothing Then

                Dim elementosCumplieronCriterios As Collection(Of PositionCell) = ProcesarCriterios(ElementosPorVerificarTodosCriterios)
                '' Search the range
                resultadoDeFormula.AgregaListaCompleta(elementosCumplieronCriterios)
                resultadoDeFormula.PosicionFormula = PositionOriginCell

                Dim resultadoFormulaOriginal = MyBase.Evaluate(args)
                resultadoDeFormula.Resultado = FarPoint.CalcEngine.CalcConvert.ToDouble(resultadoFormulaOriginal)

            End If
            Return resultadoDeFormula

        Catch ex As Exception

            System.Diagnostics.Debug.WriteLine(ex)
            Return FarPoint.CalcEngine.CalcError.Value
        Finally

        End Try


    End Function

    Private Function ProcesarCriterios(ElementosPorVerificarTodosCriterios As List(Of EstructuraRangoCriterio)) As Collection(Of PositionCell)
        'Paso 1 obtener filtro uno para descartar filas, el primer metodo recorre todas las celdas
        Dim elementosEvaluados As Collection(Of PositionCell) =
            ObtieneElementosEvaluadosEnFormula(ElementosPorVerificarTodosCriterios(0).RangoCriterio, ElementosPorVerificarTodosCriterios(0).Criterio)
        'Eliminar el primer rangoCriterio
        ElementosPorVerificarTodosCriterios.RemoveAt(0)

        'Luego revisa el rango reducido solo tomando en cuenta las filas que cumplieron el criterio anterior
        For Each itemRangoCriterio In ElementosPorVerificarTodosCriterios
            Dim rangoCriterio As CalcReference = itemRangoCriterio.RangoCriterio
            Dim criterio As String = itemRangoCriterio.Criterio
            elementosEvaluados = ProcesarSoloFilasPosibles(elementosEvaluados, rangoCriterio, criterio)
        Next
        Return elementosEvaluados


    End Function

    Private Function ProcesarSoloFilasPosibles(ByVal filasEvaluadas As Collection(Of PositionCell), ByVal rangoCriterio As CalcReference,
                                               ByVal criterio As String) As Collection(Of PositionCell)

        Try


            Dim columna As Integer = rangoCriterio.Column

            For Each indice In filasEvaluadas.Reverse
                Dim valorParaAplicarCriterio As String = rangoCriterio.GetValue(indice.Row, columna).ToString
                'Aqui podria evaluar mas condiciones como mayor que >, etc por el momento solo esta si es igual al criterio
                If String.Equals(valorParaAplicarCriterio, criterio) = False Then
                    'Elimina el item de la lista porque no cumple con el criterio actual. Solo cumplió el criterio pasado.
                    filasEvaluadas.Remove(indice)
                End If
            Next
            Return filasEvaluadas
        Catch ex As Exception
            Throw
        End Try
    End Function



    Private Function CreaListaDeRangoYCriterioParaVerificar(args() As Object) As List(Of EstructuraRangoCriterio)
        Dim listaTemp As New List(Of EstructuraRangoCriterio)
        'Despues del args(4) el args(6),args(8),args(10) los pares son criterios
        'Despues del args(3) el args(5),args(7),args(9) los pares son rangos
        For indiceContador = 1 To args.Count - 1 Step 2
            Dim nuevaEstructuraRangoCriterio As New EstructuraRangoCriterio

            'El par si
            'If esNumeroImpar(indiceContador) Then
            Dim rangoCriterio As CalcReference = args(indiceContador)
            nuevaEstructuraRangoCriterio.RangoCriterio = rangoCriterio



            Dim criterio As String = args(indiceContador + 1).ToString
            nuevaEstructuraRangoCriterio.Criterio = criterio
            listaTemp.Add(nuevaEstructuraRangoCriterio)

        Next

        Return listaTemp

    End Function


    Private Function esNumeroImpar(ByVal numero As Decimal)

        If (numero Mod 2) = 0 Then
            'El número es par.
            Return True
        Else
            'El número es impar.
            Return False
        End If
        Return True

    End Function
    'Private Function CalculaCantidadRangosPorEvaluar() As Integer

    'End Function


    Private Shared Function ObtieneElementosEvaluadosEnFormula(range As CalcReference, criteria As String) As Collection(Of PositionCell)
        Dim coordenadasItemsEvaluados As New Collection(Of PositionCell)

        For indiceFila As Integer = range.Row To range.Row + (range.RowCount - 1)
            For indiceColumna As Integer = range.Column To range.Column + (range.ColumnCount - 1)
                Dim valor = range.GetValue(indiceFila, indiceColumna)
                If EsCeldaConValor(valor) Then
                    If EsValorString(valor) Then
                        If valor.ToString = criteria Then
                            'StringBuilderlocal.Append([String].Format("[{0}, {1}]", i, j))
                            coordenadasItemsEvaluados.Add(New PositionCell With {.Column = indiceColumna, .Row = indiceFila})

                            'Sum if puede usarse de diferentes maneras no solo texto puede ser mayor que, etc. https://exceljet.net/excel-functions/excel-sumifs-function
                        End If
                    End If
                End If
            Next
        Next
        Return coordenadasItemsEvaluados
    End Function



    Private Shared Function EsCeldaConValor(ByRef valor As Object) As Boolean


        If valor Is Nothing Then
            Return False
        End If
        Return True
    End Function

    Private Shared Function EsValorString(ByVal valor As Object) As Boolean

        If TypeOf valor Is String Then
            Return True
        End If

        Return False

    End Function




End Class

Enum OperatorType
    Equal
    NotEqual
    GreaterThanEqual
    GreaterThan
    LessThanEqual
    LessThan


End Enum


Friend Class EstructuraRangoCriterio

    Property Criterio As String
    Property RangoCriterio As CalcReference

End Class
