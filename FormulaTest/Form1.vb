
Imports System.Text.RegularExpressions
Imports FarPoint.Excel
Imports FarPoint.Win
Imports FarPoint.Win.Spread

Public Class Form1



    Private Property FpSpreadInstance As Spread.FpSpread



    Private editorNameBox As New FarPoint.Win.Spread.NameBox


    Private editorFormulaTextBox As New FarPoint.Win.Spread.FormulaTextBox

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.


        FpSpreadInstance = New FpSpread()
        OpenSpreadFile("E:\documentos usuarios\melendezgl\mis documentos\visual studio 2015\Projects\FormulaTest\FormulaTest\files\Prueba de modelo AE.xlsx")


        'FpSpread1 = hojaSpreadSheet
        FpSpreadInstance.Dock = DockStyle.Fill
        FpSpreadInstance.AllowUserFormulas = True




        'Dim dataModel As FarPoint.Win.Spread.Model.DefaultSheetDataModel = FpSpread1.Sheets(0).Models.Data
        'Dim fi As FarPoint.CalcEngine.FunctionInfo
        'fi = dataModel.GetCustomFunction("SUMIF")

        Me.Panel3.Controls.Add(FpSpreadInstance)


        'FpSpread1.LoadFormulas(True)
        CreateFormulaTextBox()
        CreateFormulaNameBox()





    End Sub

    Private Sub CreateFormulaTextBox()
        Try




            editorFormulaTextBox.Name = "EditorFormula"
            editorFormulaTextBox.Dock = DockStyle.Fill
            Me.SplitContainer1.Panel2.Controls.Add(editorFormulaTextBox)
            editorFormulaTextBox.Attach(FpSpreadInstance)


        Catch ex As Exception
            Throw
        End Try
    End Sub

    Private Sub CreateFormulaNameBox()


        editorNameBox.Name = "EditorNameBox"
        editorNameBox.Dock = DockStyle.Fill
        Me.SplitContainer1.Panel1.Controls.Add(editorNameBox)
        editorNameBox.Attach(FpSpreadInstance)



    End Sub





    Private Sub btnAbrir_Click(sender As Object, e As EventArgs) Handles btnRecalcular.Click
        'Dim OpenFileDialog1 As New OpenFileDialog
        'Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        'If result = DialogResult.OK Then

        '    'Me.instanciaFormuladorSpread.AbrirArchivoDesdeRutaFisica(OpenFileDialog1.FileName)
        'End If
        Dim horainicio As DateTime = DateTime.Now




        Me.FpSpreadInstance.Sheets(0).RecalculateAll()

        Dim horaFin As DateTime = DateTime.Now

        Dim duracion = horaFin.Subtract(horainicio)

        MsgBox(duracion.ToString)
    End Sub



    Private Sub btnMostrar_Click(sender As Object, e As EventArgs) Handles btnMostrar.Click


        Dim ems As FarPoint.Win.Spread.Model.IExpressionSupport
        ems = FpSpreadInstance.ActiveSheet.Models.Data
        Dim sformula As String
        sformula = ems.GetFormula(Me.FpSpreadInstance.ActiveSheet.ActiveRowIndex, Me.FpSpreadInstance.ActiveSheet.ActiveColumnIndex)
        Dim exp As FarPoint.CalcEngine.Expression


        Dim es As FarPoint.Win.Spread.Model.IExpressionSupport2
        es = FpSpreadInstance.ActiveSheet.Models.Data
        exp = es.ParseFormula(Me.FpSpreadInstance.ActiveSheet.ActiveRowIndex, Me.FpSpreadInstance.ActiveSheet.ActiveColumnIndex, sformula)
        MsgBox("The parsed formula is " & exp.ToString())
        Dim ret As String
        ret = es.UnparseFormula(Me.FpSpreadInstance.ActiveSheet.ActiveRowIndex, Me.FpSpreadInstance.ActiveSheet.ActiveColumnIndex, exp)
        MessageBox.Show("The unparsed formula is " & ret)



        'FpSpread1.ShowDependents(0, 0, 1)
        'FpSpread1.ShowPrecedents(0, 0, 1)

    End Sub






    Shared Function ObtengaLasCeldasDeLaFormula(ByVal formula As String) As IList(Of String)
        Dim celdasDeLaFormula As New List(Of String)
        Dim expresionRegular As New Regex("[a-zA-Z]*\d+")
        Dim coleccionDeEquivalencias As MatchCollection = expresionRegular.Matches(formula)

        For contador As Integer = 0 To coleccionDeEquivalencias.Count - 1
            celdasDeLaFormula.Add(coleccionDeEquivalencias(contador).Value)
        Next

        Return celdasDeLaFormula.OrderByDescending(Function(a) a).ToList()
    End Function

    Private Sub btnTipoReferencia_Click(sender As Object, e As EventArgs)
        Me.FpSpreadInstance.Sheets(0).ReferenceStyle = Spread.Model.ReferenceStyle.R1C1
        Me.FpSpreadInstance.Sheets(1).ReferenceStyle = Spread.Model.ReferenceStyle.R1C1

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        editorNameBox.Detach()


        editorFormulaTextBox.Detach()


    End Sub

    Public Function OpenSpreadFile(ByVal path As String) As Boolean
        Try

            Dim esCargaExitosa As Boolean = FpSpreadInstance.OpenExcel(path, ExcelOpenFlags.DoNotRecalculateAfterLoad Or ExcelOpenFlags.TruncateEmptyRowsAndColumns)

            Return True

        Catch ex As Exception
            Throw
        End Try

    End Function


    Public Sub RegisterCustomFormula()
        Dim cfs As FarPoint.Win.Spread.Model.ICustomFunctionSupport
        cfs = FpSpreadInstance.Sheets(0).Models.Data
        cfs.AddCustomFunction(New SUMIFS())



    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Dim intanceExcute As New DescomponeFormulaIee

        intanceExcute.DescomponerFormula(FpSpreadInstance)

    End Sub
End Class


