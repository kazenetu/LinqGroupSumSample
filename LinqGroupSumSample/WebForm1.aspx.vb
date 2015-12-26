Imports System.Linq

Public Class WebForm1
    Inherits System.Web.UI.Page

    ''' <summary>
    ''' データソースアイテム
    ''' </summary>
    ''' <remarks></remarks>
    Public Class ItemClass
        Public Property Key1 As String
        Public Property Key2 As String
        Public Property Key3 As String
        Public Property Qty As Decimal
    End Class

    ''' <summary>
    ''' ページロード
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'ダミーデータを取得
        Dim dataSource As List(Of ItemClass) = Me.getDataSource()

        'キー1～3でグルーピングしてQtyの集計を行う
        Dim enumerable As IEnumerable(Of ItemClass) = dataSource.AsEnumerable()
        Dim sumQuery = enumerable.GroupBy(Function(item) Tuple.Create(item.Key1, item.Key2, item.Key3), _
                                          Function(key, item) New With {.Key = key, .QtySum = item.Sum(Function(i) i.Qty)})


        '結果を出力
        Dim sb = New StringBuilder()
        For Each item In sumQuery
            sb.Append("[" + item.Key.Item1 + ",")
            sb.Append(item.Key.Item2 + ",")
            sb.Append(item.Key.Item3 + "] = ")
            sb.Append(item.QtySum)
            sb.AppendLine("<br>")
        Next
        Me.Label1.Text = sb.ToString()


    End Sub

    ''' <summary>
    ''' データソースを取得
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getDataSource() As List(Of ItemClass)
        Dim result As New List(Of ItemClass)

        result.Add(New ItemClass() With {.Key1 = "K1_A", .Key2 = "K2_A", .Key3 = "K3_A", .Qty = 10})
        result.Add(New ItemClass() With {.Key1 = "K1_A", .Key2 = "K2_A", .Key3 = "K3_A", .Qty = 20})
        result.Add(New ItemClass() With {.Key1 = "K1_B", .Key2 = "K2_B", .Key3 = "K3_A", .Qty = 10})
        result.Add(New ItemClass() With {.Key1 = "K1_A", .Key2 = "K2_A", .Key3 = "K3_A", .Qty = 70})

        Return result
    End Function
End Class