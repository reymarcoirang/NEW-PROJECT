Public Class Form1
    Dim lv As ListViewItem
    Dim selected As String


    Public Sub PoplistView()
        ListView1.Clear()
        With ListView1
            .View = View.Details
            .GridLines = True
            .Columns.Add("id", 40)
            .Columns.Add("lastname", 110)
            .Columns.Add("firstname", 110)
            .Columns.Add("mi", 110)
            .Columns.Add("Address", 110)
            .Columns.Add("gender", 110)
            .Columns.Add("contactno", 110)
            .Columns.Add("position", 110)
        End With

        openCon()
        sql = "Select * from tblempinfo"
        rs = New ADODB.Recordset
        rs.CursorLocation = 3
        rs.Open(sql, cn, 3, 3)


        If rs.RecordCount > 0 Then
            rs.MoveFirst()
            Do Until rs.EOF
                lv = New ListViewItem(rs.Fields("empid").Value.ToString)
                lv.SubItems.Add(rs.Fields("emplastname").Value)
                lv.SubItems.Add(rs.Fields("empfirstname").Value)
                lv.SubItems.Add(rs.Fields("empmi").Value)
                lv.SubItems.Add(rs.Fields("empaddress").Value)
                lv.SubItems.Add(rs.Fields("empgender").Value)
                lv.SubItems.Add(rs.Fields("empcontactno").Value)
                lv.SubItems.Add(rs.Fields("empposition").Value)
                ListView1.Items.Add(lv)
                rs.MoveNext()


            Loop
        End If
        rs.Close()
        cn.Close()


    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PoplistView()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        If MsgBox("Are you sure to save this record?", (vbYesNo + vbQuestion)) = DialogResult.Yes Then
            openCon()
            sql = "Select * from tblempinfo"
            rs = New ADODB.Recordset
            rs.Open(sql, cn, 3, 3)
            rs.AddNew()
            rs.Fields("empid").Value = Me.txtemployeeno.Text
            rs.Fields("emplastname").Value = Me.txtlastname.Text
            rs.Fields("empfirstname").Value = Me.txtfirstname.Text
            rs.Fields("empmi").Value = Me.txtmi.Text
            rs.Fields("empaddress").Value = Me.txtaddress.Text
            rs.Fields("empgender").Value = Me.cmbgender.Text
            rs.Fields("empcontactno").Value = Me.txtcontactno.Text
            rs.Fields("empposition").Value = Me.cmbposition.Text
            rs.Update()
            rs.Close()
            cn.Close()





        End If

        PoplistView()

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim i As Integer
        For i = 0 To ListView1.SelectedItems.Count - 1
            selected = ListView1.SelectedItems(i).Text

            openCon()
            sql = "Select * from tblempinfo WHERE empinfo"
            rs = New ADODB.Recordset
            rs.CursorLocation = 3
            rs.Open(sql, cn, 3, 3)

            Me.txtemployeeno.Text = rs.Fields("empid").Value
            Me.txtlastname.Text = rs.Fields("emplastname").Value
            Me.txtfirstname.Text = rs.Fields("empfirstname").Value
            Me.txtmi.Text = rs.Fields("empmi").Value
            Me.txtaddress.Text = rs.Fields("empaddress").Value
            Me.cmbgender.Text = rs.Fields("empgender").Value
            Me.txtcontactno.Text = rs.Fields("empcontactno").Value
            Me.cmbposition.Text = rs.Fields("empposition").Value
            rs.MoveNext()




            End If


            rs.Close()
            cn.Close()





    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub
End Class
