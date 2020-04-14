Imports System.Data
Imports System.Data.OleDb
Module Module1
    Public sqlnya As String
    Public bagian1 As Integer
    Public bagian2 As Integer
    Public bagian1a As Integer = Form1.CheckedListBox1.CheckedItems.Count
    Public bagian2a As Integer = Form1.CheckedListBox1.CheckedItems.Count
    Public bagian1b As Integer = Form1.CheckedListBox2.CheckedItems.Count
    Public bagian2b As Integer = Form1.CheckedListBox2.CheckedItems.Count
    Public bagian1c As Integer = Form1.CheckedListBox3.CheckedItems.Count
    Public bagian2c As Integer = Form1.CheckedListBox3.CheckedItems.Count
    Public hasil As String
    Public jk As String
    Public conn As OleDbConnection
    Public CMD As OleDbCommand
    Public DS As New DataSet
    Public DA As OleDbDataAdapter
    Public RD As OleDbDataReader
    Public lokasidata As String
    Public Sub konek()
        lokasidata = "provider=microsoft.jet.oledb.4.0;data source=survey.mdb"
        conn = New OleDbConnection(lokasidata)
        If conn.State = ConnectionState.Closed Then
            conn.Open()
        End If
    End Sub
    Public Sub summon()
        konek()
        DA = New OleDb.OleDbDataAdapter("SELECT * FROM tb_survey", conn)
        DS = New DataSet
        DS.Clear()
        DA.Fill(DS, "tb_survey")
        Form3.DataGridView1.DataSource = DS.Tables("tb_survey")
        Form3.DataGridView1.Enabled = True
    End Sub
    Public Sub go()
        Dim objcmd As New System.Data.OleDb.OleDbCommand
        Call konek()
        objcmd.Connection = conn
        objcmd.CommandType = CommandType.Text
        objcmd.CommandText = sqlnya
        objcmd.ExecuteNonQuery()
        objcmd.Dispose()
    End Sub
End Module

