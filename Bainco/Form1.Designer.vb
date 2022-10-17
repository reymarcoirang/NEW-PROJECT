Public Class Form1
    Dim lv As New ListViewItem
    Dim selected As String

    Private Sub ClearAllTextbox()
        Me.txtemployeeno.Text = ""
        Me.txtlastname.Text = ""
        Me.txtfirstname.Text = ""
        Me.txtmi.Text = ""
        Me.txtaddress.Text = ""
        Me.cmbgender.Text = ""
        Me.txtcontact.Text = ""
        Me.cmbposition.Text = ""

    End Sub

    Private Sub PoplistView()
        ListView1.Clear()

        With ListView1
            .HideSelection = True
            .FullRowSelect = True
            .View = View.Details
            .GridLines = True
            .Columns.Add("ID", 40)
            .Columns.Add("Lastname", 110)
            .Columns.Add("Firstname", 110)
            .Columns.Add("Middlename", 110)
            .Columns.Add("Address", 100)
            .Columns.Add("Gender", 150)
            .Columns.Add("Contact no.", 110)
            .Columns.Add("Position", 110)
        End With
        openCon()
        sql = "Select * FROM tblempinfo"
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
                lv.SubItems.Add(rs.Fields("empcontact").Value)
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

    Private Sub addbtn_Click(sender As Object, e As EventArgs) Handles addbtn.Click
        If MsgBox("Are you sure to save this record?", (vbYesNo + vbQuestion)) = DialogResult.Yes Then
            openCon()
            sql = "Select * FROM tblempinfo"
            rs = New ADODB.Recordset
            rs.Open(sql, cn, 3, 3)
            rs.AddNew()
            rs.Fields("empid").Value = Me.txtemployeeno.Text
            rs.Fields("emplastname").Value = Me.txtlastname.Text
            rs.Fields("empfirstname").Value = Me.txtfirstname.Text
            rs.Fields("empmi").Value = Me.txtmi.Text
            rs.Fields("empaddress").Value = Me.txtaddress.Text
            rs.Fields("empgender").Value = Me.cmbgender.Text
            rs.Fields("empcontact").Value = Me.txtcontact.Text
            rs.Fields("empposition").Value = Me.cmbposition.Text
            rs.Update()
            rs.Close()
            cn.Close()
            ClearAllTextbox()

        End If
        If Not (txtemployeeno.Text.Length = 0 Or txtaddress.TextLength = 0 Or txtfirstname.TextLength = 0 Or txtmi.TextLength = 0 Or txtlastname.TextLength = 0 Or txtcontact.TextLength = 0 Or cmbposition.Text.Length = 0 Or cmbgender.Text.Length = 0) Then
            If MsgBox("Are you sure to save this record?", (vbYesNo + vbQuestion)) = DialogResult.Yes Then
                openCon()
                sql = "Select * FROM tblempinfo"
                rs = New ADODB.Recordset
                rs.Open(sql, cn, 3, 3)
                rs.AddNew()
                rs.Fields("empid").Value = Me.txtemployeeno.Text
                rs.Fields("emplastname").Value = Me.txtlastname.Text
                rs.Fields("empfirstname").Value = Me.txtfirstname.Text
                rs.Fields("empmi").Value = Me.txtmi.Text
                rs.Fields("empaddress").Value = Me.txtaddress.Text
                rs.Fields("empgender").Value = Me.cmbgender.Text
                rs.Fields("empcontact").Value = Me.txtcontact.Text
                rs.Fields("empposition").Value = Me.cmbposition.Text
                rs.Update()
                rs.Close()
                cn.Close()
                ClearAllTextbox()

            End If

            PoplistView()
        Else
            MsgBox("Please Fill Up! WAG NA MAKULET!!!", vbOKOnly)

        End If
        PoplistView()

    End Sub

    Private Sub btnupdate_Click(sender As Object, e As EventArgs) Handles btnupdate.Click
        If MsgBox("Are you sure to Update this record?", (vbYesNo + vbQuestion)) = DialogResult.Yes Then
            openCon()
            sql = "Select * FROM tblempinfo WHERE empid = " & selected & " "
            rs = New ADODB.Recordset
            rs.Open(sql, cn, 3, 3)
            rs.Fields("empid").Value = Me.txtemployeeno.Text
            rs.Fields("emplastname").Value = Me.txtlastname.Text
            rs.Fields("empfirstname").Value = Me.txtfirstname.Text
            rs.Fields("empmi").Value = Me.txtmi.Text
            rs.Fields("empaddress").Value = Me.txtaddress.Text
            rs.Fields("empgender").Value = Me.cmbgender.Text
            rs.Fields("empcontact").Value = Me.txtcontact.Text
            rs.Fields("empposition").Value = Me.cmbposition.Text
            rs.Update()
            rs.Close()
            cn.Close()
            ClearAllTextbox()

        End If
        If Not (txtemployeeno.Text.Length = 0 Or txtaddress.TextLength = 0 Or txtfirstname.TextLength = 0 Or txtmi.TextLength = 0 Or txtlastname.TextLength = 0 Or txtcontact.TextLength = 0 Or cmbposition.Text.Length = 0 Or cmbgender.Text.Length = 0) Then
            If MsgBox("Are you sure to Update this record?", (vbYesNo + vbQuestion)) = DialogResult.Yes Then
                openCon()
                sql = "Select * FROM tblempinfo WHERE empid = " & selected & " "
                rs = New ADODB.Recordset
                rs.Open(sql, cn, 3, 3)
                rs.Fields("empid").Value = Me.txtemployeeno.Text
                rs.Fields("emplastname").Value = Me.txtlastname.Text
                rs.Fields("empfirstname").Value = Me.txtfirstname.Text
                rs.Fields("empmi").Value = Me.txtmi.Text
                rs.Fields("empaddress").Value = Me.txtaddress.Text
                rs.Fields("empgender").Value = Me.cmbgender.Text
                rs.Fields("empcontact").Value = Me.txtcontact.Text
                rs.Fields("empposition").Value = Me.cmbposition.Text
                rs.Update()
                rs.Close()
                cn.Close()
                ClearAllTextbox()

            End If
        Else
            MsgBox("Please Select You Wan't To Update!", vbOKOnly)
        End If
        PoplistView()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListView1.SelectedIndexChanged
        Dim i As Integer
        For i = 0 To ListView1.SelectedItems.Count - 1
            selected = ListView1.SelectedItems(i).Text

            openCon()
            sql = "Select * FROM tblempinfo WHERE empid = " & selected & " "
            rs = New ADODB.Recordset
            rs.CursorLocation = 3
            rs.Open(sql, cn, 3, 3)


            If rs.RecordCount > 0 Then
                rs.MoveFirst()

                Me.txtemployeeno.Text = rs.Fields("empid").Value
                Me.txtlastname.Text = rs.Fields("emplastname").Value
                Me.txtfirstname.Text = rs.Fields("empfirstname").Value
                Me.txtmi.Text = rs.Fields("empmi").Value
                Me.txtaddress.Text = rs.Fields("empaddress").Value
                Me.cmbgender.Text = rs.Fields("empgender").Value
                Me.txtcontact.Text = rs.Fields("empcontact").Value
                Me.cmbposition.Text = rs.Fields("empposition").Value
                rs.MoveNext()


            End If

            rs.Close()
            cn.Close()
        Next


    End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        If MsgBox("Are you sure to Delete this record?", (vbYesNo + vbQuestion)) = DialogResult.Yes Then
            openCon()
            sql = "Select * FROM tblempinfo WHERE empid = " & selected & " "
            rs = New ADODB.Recordset
            rs.Open(sql, cn, 3, 3)
            rs.Delete()
            rs.Update()
            rs.Close()
            cn.Close()
            ClearAllTextbox()

            PoplistView()

        End If

        If ListView1.SelectedItems.Count > 0 Then
            If MsgBox("Are you sure to Delete this record?", (vbYesNo + vbQuestion)) = DialogResult.Yes Then
                openCon()
                sql = "Select * FROM tblempinfo WHERE empid = " & selected & " "
                rs = New ADODB.Recordset
                rs.Open(sql, cn, 3, 3)
                rs.Delete()
                rs.Update()
                rs.Close()
                cn.Close()
                ClearAllTextbox()

                PoplistView()

            End If
        Else
            MsgBox("Please Select Do You Wan't To Delete!", vbOKOnly)
        End If

    End Sub
End Class
