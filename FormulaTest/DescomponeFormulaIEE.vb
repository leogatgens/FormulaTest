
Imports System.Globalization

Imports FarPoint.CalcEngine
Imports FarPoint.Win.Spread
<Assembly: CLSCompliant(True)>
Public Class DescomponeFormulaIee



    Private listaMaesResultadoFormulas As List(Of IeeResult)

    Private listaFormulasSinCodificacion As List(Of PositionCell)


    Public Sub New()

        listaMaesResultadoFormulas = New List(Of IeeResult)
        listaFormulasSinCodificacion = New List(Of PositionCell)
    End Sub
    Public Function DescomponerFormula(ByVal objetoSpread As FpSpread) As DataTable
        If objetoSpread Is Nothing Then
            Throw New ArgumentException("SpreadSheet null")
        End If
        objetoSpread.ActiveSheetIndex = 0
        Dim vistaActualSpread As SheetView = objetoSpread.GetRootWorkbook().GetActiveWorkbook().GetSheetView()
        Dim dataModel As Model.DefaultSheetDataModel = vistaActualSpread.Models.Data


        dataModel = IterateModelToEvaluteFormula(dataModel)

        Return ConvertToDataTable()
    End Function



    Private Function IterateModelToEvaluteFormula(dataModel As Model.DefaultSheetDataModel) As Model.DefaultSheetDataModel
        Try




            For rowINdex = 0 To dataModel.RowCount - 1
                IterateColumnSearchingData(rowINdex, dataModel)
            Next




            Return dataModel


        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex)
            Throw
        End Try

    End Function


    Private Sub IterateColumnSearchingData(ByRef indiceFila As Integer, ByRef dataModel As Model.DefaultSheetDataModel)
        Dim nextColumnWithData As Integer = 0

        For indiceDelFor = 0 To dataModel.ColumnCount - 1
            ObtainNextColumnIndex(indiceFila, dataModel, indiceDelFor, nextColumnWithData)

            If nextColumnWithData = -1 Then
                Exit Sub
            Else

                If ValidateAndSetFormula(indiceFila, nextColumnWithData, dataModel) Then
                    Dim resultValue As Object = EvaluateFormula(indiceFila, dataModel, nextColumnWithData)
                    ManageFormulaResult(resultValue, indiceFila, nextColumnWithData)
                End If

            End If





        Next


    End Sub

    Private Shared Sub ObtainNextColumnIndex(indiceFila As Integer, dataModel As Model.DefaultSheetDataModel,
                                                     ByRef indiceDelFor As Integer, ByRef IndiceProximaColumnaConDato As Integer)
        If indiceDelFor < IndiceProximaColumnaConDato Then
            indiceDelFor = IndiceProximaColumnaConDato
        End If

        IndiceProximaColumnaConDato = dataModel.NextNonEmptyColumnInRow(indiceFila, indiceDelFor)
    End Sub

    Private Function ValidateAndSetFormula(indiceFila As Integer, nextColumnIndex As Integer, dataModel As Model.DefaultSheetDataModel) As Boolean
        Dim resultado As Boolean = False
        Try


            Dim cellformula = dataModel.GetFormula(indiceFila, nextColumnIndex)

            If cellformula <> String.Empty Then
                ''Limpia las formulas y las vuelve a setear para que tome la formula custom sumif
                dataModel.SetFormula(indiceFila, nextColumnIndex, String.Empty)
                dataModel.SetFormula(indiceFila, nextColumnIndex, cellformula)
                resultado = True
                System.Diagnostics.Debug.WriteLine(cellformula + "Asignada" + "___Coordenada:" + indiceFila.ToString + "-" + nextColumnIndex.ToString)
            End If
            Return resultado
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex)
            Throw
        End Try
    End Function


    Private Sub ManageFormulaResult(valorResultado As Object, indiceFila As Integer, IndiceColumna As Integer)

        If TypeOf valorResultado Is FormulaResult Then
            CreaNuevaFilaResultado(valorResultado)
        ElseIf IsNumeric(valorResultado) Then
            listaFormulasSinCodificacion.Add(New PositionCell With {.Row = indiceFila, .Column = IndiceColumna})
        Else
            Throw New ArgumentException(" resultado of formula is not type:ResultadoFormula. An errror must ocurred")

        End If
    End Sub



    Private Function EvaluateFormula(rowIndex As Integer, dataModel As Model.DefaultSheetDataModel, columnIndex As Integer) As Object
        Dim result As Object = Nothing

        Try


            Dim formulaexpression As Expression = dataModel.GetExpression(rowIndex, columnIndex)

            '''''NEED HELP HERE
            decomposeFormulaExpresion(formulaexpression)

            result = dataModel.EvaluateExpression(rowIndex, columnIndex, formulaexpression)

        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex)
            Throw
        End Try


        Return result
    End Function

    ''' <summary>
    ''' I need Help here, how to evaluate formulas retuning objects with plus operator???
    ''' </summary>
    ''' <param name="formulaexpresion"></param>
    Private Sub decomposeFormulaExpresion(formulaexpresion As Expression)
        Try

            Dim listaFormulas As New List(Of Expression)


            If TypeOf formulaexpresion Is FunctionExpression Then
                Dim formulas = TryCast(formulaexpresion, FarPoint.CalcEngine.FunctionExpression)

            ElseIf TypeOf formulaexpresion Is BinaryOperatorExpression Then

                Dim formulas = TryCast(formulaexpresion, FarPoint.CalcEngine.BinaryOperatorExpression)

                listaFormulas.Add(formulas.Arg0)
                listaFormulas.Add(formulas.Arg1)
            End If








        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex)
            'Throw
        End Try


    End Sub

    Private Sub CreaNuevaFilaResultado(ByVal dataModel As Model.DefaultSheetDataModel, ByVal indiceFila As Integer, ByVal indiceColumna As Integer)

        If dataModel Is Nothing Then
            Throw New ArgumentException("El modelo es invalido")
        End If
        Dim valor = dataModel.GetValue(indiceFila, indiceColumna)

        Dim respuesta As New IeeResult

        respuesta.Valor = valor
        'respuesta.Plantilla = "Empresa"
        'respuesta.IndiceHoja = 1

        listaMaesResultadoFormulas.Add(respuesta)
    End Sub



    Private Sub CreaNuevaFilaResultado(ByVal valorResultado As Object)

        If valorResultado Is Nothing Then
            Throw New ArgumentException("Metodo CreaNuevaFilaResultado fallo porque el resultado de la formula es invalido")
        End If

        Dim respuesta As New IeeResult
        respuesta.Valor = valorResultado
        'respuesta.Plantilla = "Empresa"
        'respuesta.IndiceHoja = 1
        listaMaesResultadoFormulas.Add(respuesta)


    End Sub



    Private Function ConvertToDataTable() As DataTable


        If listaMaesResultadoFormulas Is Nothing Then
            Throw New ArgumentException("Lista vacia")
        End If
        Using table = New DataTable()
            table.Locale = CultureInfo.InvariantCulture
            If Not listaMaesResultadoFormulas.Any Then
                'don't know schema ....
                Return table
            End If
            Dim fields() = listaMaesResultadoFormulas.First.GetType.GetProperties
            For Each field In fields
                table.Columns.Add(field.Name, field.PropertyType)
            Next
            For Each item In listaMaesResultadoFormulas
                Dim row As DataRow = table.NewRow()
                For Each field In fields
                    Dim p = item.GetType.GetProperty(field.Name)
                    row(field.Name) = p.GetValue(item, Nothing)
                Next
                table.Rows.Add(row)
            Next
            Return table
        End Using


    End Function






End Class




