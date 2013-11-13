Imports System.Data.SqlClient

Public Class Form1
    Dim codes(22) As String
    Dim cn As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=resources\db.mdb;Jet OLEDB:Database Password=prozergman;")
    Dim ds As New DataSet()
    Dim sql As String
    Dim adaptor As OleDb.OleDbDataAdapter
    Dim k As Integer
    Dim checkboxes(22) As CheckBox
    Dim texts(22) As TextBox
    Dim builder As OleDb.OleDbCommandBuilder

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        k = 1
        cn.Open()
        sql = "SELECT * FROM codes"
        adaptor = New OleDb.OleDbDataAdapter(sql, cn)
        adaptor.Fill(ds, "getcodes")
        PopulateCodes()
        PopulateChk()
        PopulateNames()
        sql = "SELECT * FROM pass"
        adaptor = New OleDb.OleDbDataAdapter(sql, cn)
        adaptor.Fill(ds, "getpass")
    End Sub

    Private Function PopulateNames()
        agam.Text = ds.Tables("getcodes").Rows(0).Item(1)
        hamal_alof.Text = ds.Tables("getcodes").Rows(1).Item(1)
        hadan_alof.Text = ds.Tables("getcodes").Rows(2).Item(1)
        modiin.Text = ds.Tables("getcodes").Rows(3).Item(1)
        logistica.Text = ds.Tables("getcodes").Rows(4).Item(1)
        matnap.Text = ds.Tables("getcodes").Rows(5).Item(1)
        himush.Text = ds.Tables("getcodes").Rows(6).Item(1)
        fire_center.Text = ds.Tables("getcodes").Rows(7).Item(1)
        shalishut.Text = ds.Tables("getcodes").Rows(8).Item(1)
        merkaz_hamishi.Text = ds.Tables("getcodes").Rows(9).Item(1)
        medicine.Text = ds.Tables("getcodes").Rows(10).Item(1)
        handasa.Text = ds.Tables("getcodes").Rows(11).Item(1)
        kesher.Text = ds.Tables("getcodes").Rows(12).Item(1)
        lishkat_agam.Text = ds.Tables("getcodes").Rows(13).Item(1)
        lishkat_alof.Text = ds.Tables("getcodes").Rows(14).Item(1)
        ta_pisga.Text = ds.Tables("getcodes").Rows(15).Item(1)
        ta_zitar.Text = ds.Tables("getcodes").Rows(16).Item(1)
        mashlak.Text = ds.Tables("getcodes").Rows(17).Item(1)
        malnap.Text = ds.Tables("getcodes").Rows(18).Item(1)
        barak.Text = ds.Tables("getcodes").Rows(19).Item(1)
        hadat.Text = ds.Tables("getcodes").Rows(20).Item(1)
        temp.Text = ds.Tables("getcodes").Rows(21).Item(1)
    End Function

    Private Function PopulateCodes()
        codes(1) = ds.Tables("getcodes").Rows(0).Item(2)
        txt1.Text = codes(1)
        texts(1) = txt1
        codes(2) = ds.Tables("getcodes").Rows(1).Item(2)
        txt2.Text = codes(2)
        texts(2) = txt2
        codes(3) = ds.Tables("getcodes").Rows(2).Item(2)
        txt3.Text = codes(3)
        texts(3) = txt3
        codes(4) = ds.Tables("getcodes").Rows(3).Item(2)
        txt4.Text = codes(4)
        texts(4) = txt4
        codes(5) = ds.Tables("getcodes").Rows(4).Item(2)
        txt5.Text = codes(5)
        texts(5) = txt5
        codes(6) = ds.Tables("getcodes").Rows(5).Item(2)
        txt6.Text = codes(6)
        texts(6) = txt6
        codes(7) = ds.Tables("getcodes").Rows(6).Item(2)
        txt7.Text = codes(7)
        texts(7) = txt7
        codes(8) = ds.Tables("getcodes").Rows(7).Item(2)
        txt8.Text = codes(8)
        texts(8) = txt8
        codes(9) = ds.Tables("getcodes").Rows(8).Item(2)
        txt9.Text = codes(9)
        texts(9) = txt9
        codes(10) = ds.Tables("getcodes").Rows(9).Item(2)
        txt10.Text = codes(10)
        texts(10) = txt10
        codes(11) = ds.Tables("getcodes").Rows(10).Item(2)
        txt11.Text = codes(11)
        texts(11) = txt11
        codes(12) = ds.Tables("getcodes").Rows(11).Item(2)
        txt12.Text = codes(12)
        texts(12) = txt12
        codes(13) = ds.Tables("getcodes").Rows(12).Item(2)
        txt13.Text = codes(13)
        texts(13) = txt13
        codes(14) = ds.Tables("getcodes").Rows(13).Item(2)
        txt14.Text = codes(14)
        texts(14) = txt14
        codes(15) = ds.Tables("getcodes").Rows(14).Item(2)
        txt15.Text = codes(15)
        texts(15) = txt15
        codes(16) = ds.Tables("getcodes").Rows(15).Item(2)
        txt16.Text = codes(16)
        texts(16) = txt16
        codes(17) = ds.Tables("getcodes").Rows(16).Item(2)
        txt17.Text = codes(17)
        texts(17) = txt17
        codes(18) = ds.Tables("getcodes").Rows(17).Item(2)
        txt18.Text = codes(18)
        texts(18) = txt18
        codes(19) = ds.Tables("getcodes").Rows(18).Item(2)
        txt19.Text = codes(19)
        texts(19) = txt19
        codes(20) = ds.Tables("getcodes").Rows(19).Item(2)
        txt20.Text = codes(20)
        texts(20) = txt20
        codes(21) = ds.Tables("getcodes").Rows(20).Item(2)
        txt21.Text = codes(21)
        texts(21) = txt21
        codes(22) = ds.Tables("getcodes").Rows(21).Item(2)
        txt22.Text = codes(22)
        texts(22) = txt22
    End Function

    Private Function Charge()
        Dim k = 1
        Dim len
        Dim temp
        Dim added
        result3.Text = ""
        Do While k < 23
            temp = codes(k)
            len = 0
            Do While temp > 0
                len = len + 1
                temp = Fix(temp / 10)
            Loop
            If checkboxes(k).Checked = True Then
                result3.Text = CStr(result3.Text) + CStr(codes(k))
            Else
                added = codes(k)
                Do While added = codes(k)
                    added = Fix(Rnd() * (System.Math.Pow(10, len) - System.Math.Pow(10, len - 1))) + System.Math.Pow(10, len - 1)
                Loop
                result3.Text = CStr(result3.Text) + CStr(added)
            End If
            k = k + 1
        Loop
    End Function

    Private Sub create_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles create2.Click
        Charge()
        Clipboard.SetDataObject(result3.Text)
    End Sub

    Public Function PopulateChk()
        checkboxes(1) = agam
        checkboxes(2) = hamal_alof
        checkboxes(3) = hadan_alof
        checkboxes(4) = modiin
        checkboxes(5) = logistica
        checkboxes(6) = matnap
        checkboxes(7) = himush
        checkboxes(8) = fire_center
        checkboxes(9) = shalishut
        checkboxes(10) = merkaz_hamishi
        checkboxes(11) = medicine
        checkboxes(12) = handasa
        checkboxes(13) = kesher
        checkboxes(14) = lishkat_agam
        checkboxes(15) = lishkat_alof
        checkboxes(16) = ta_pisga
        checkboxes(17) = ta_zitar
        checkboxes(18) = mashlak
        checkboxes(19) = malnap
        checkboxes(20) = barak
        checkboxes(21) = hadat
        checkboxes(22) = temp
    End Function

    Private Sub edit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles edit2.Click
        If pass2.Text = ds.Tables("getpass").Rows(0).Item(0) Then
            txt1.Visible = True
            txt2.Visible = True
            txt3.Visible = True
            txt4.Visible = True
            txt5.Visible = True
            txt6.Visible = True
            txt7.Visible = True
            txt8.Visible = True
            txt9.Visible = True
            txt10.Visible = True
            txt11.Visible = True
            txt12.Visible = True
            txt13.Visible = True
            txt14.Visible = True
            txt15.Visible = True
            txt16.Visible = True
            txt17.Visible = True
            txt18.Visible = True
            txt19.Visible = True
            txt20.Visible = True
            txt21.Visible = True
            txt22.Visible = True
            create2.Visible = False
            result3.Visible = False
            pass2.Text = ""
            pass2.Visible = False
            edit2.Visible = False
            Label3.Visible = False
            commitcodes2.Visible = True
            back2.Visible = True
            LineShape1.Visible = True
            namenum.Visible = True
            changename2.Visible = True
            commitname.Visible = True
            RectangleShape1.Visible = True
            RectangleShape2.Visible = True
            RectangleShape3.Visible = True
            RectangleShape4.Visible = False
            commitpass2.Visible = True
            topass2.Visible = True
            Label5.Visible = True
            Label6.Visible = True
            Label7.Visible = True
            Label8.Visible = False
            t1.Visible = True
            t2.Visible = True
            t3.Visible = True
            t4.Visible = True
            t5.Visible = True
            t6.Visible = True
            t7.Visible = True
            t8.Visible = True
            t9.Visible = True
            t10.Visible = True
            t11.Visible = True
            t12.Visible = True
            t13.Visible = True
            t14.Visible = True
            t15.Visible = True
            t16.Visible = True
            t17.Visible = True
            t18.Visible = True
            t19.Visible = True
            t20.Visible = True
            t21.Visible = True
            t22.Visible = True
        End If
    End Sub

    Private Sub back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles back2.Click
        txt1.Visible = False
        txt2.Visible = False
        txt3.Visible = False
        txt4.Visible = False
        txt5.Visible = False
        txt6.Visible = False
        txt7.Visible = False
        txt8.Visible = False
        txt9.Visible = False
        txt10.Visible = False
        txt11.Visible = False
        txt12.Visible = False
        txt13.Visible = False
        txt14.Visible = False
        txt15.Visible = False
        txt16.Visible = False
        txt17.Visible = False
        txt18.Visible = False
        txt19.Visible = False
        txt20.Visible = False
        txt21.Visible = False
        txt22.Visible = False
        create2.Visible = True
        result3.Visible = True
        pass2.Visible = True
        edit2.Visible = True
        Label3.Visible = True
        commitcodes2.Visible = False
        back2.Visible = False
        namenum.Visible = False
        LineShape1.Visible = False
        changename2.Visible = False
        commitname.Visible = False
        RectangleShape1.Visible = False
        RectangleShape2.Visible = False
        RectangleShape3.Visible = False
        RectangleShape4.Visible = True
        commitpass2.Visible = False
        topass2.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = True
        t1.Visible = False
        t2.Visible = False
        t3.Visible = False
        t4.Visible = False
        t5.Visible = False
        t6.Visible = False
        t7.Visible = False
        t8.Visible = False
        t9.Visible = False
        t10.Visible = False
        t11.Visible = False
        t12.Visible = False
        t13.Visible = False
        t14.Visible = False
        t15.Visible = False
        t16.Visible = False
        t17.Visible = False
        t18.Visible = False
        t19.Visible = False
        t20.Visible = False
        t21.Visible = False
        t22.Visible = False
    End Sub

    Private Sub commitcodes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles commitcodes2.Click
        sql = "SELECT * FROM codes"
        adaptor = New OleDb.OleDbDataAdapter(sql, cn)
        adaptor.Fill(ds, "getcodes")
        builder = New OleDb.OleDbCommandBuilder(adaptor)
        Dim k As Integer
        k = 1
        Do While k < 23
            ds.Tables("getcodes").Rows(k - 1).Item(2) = texts(k).Text
            k = k + 1
        Loop
        adaptor.Update(ds, "getcodes")
        MsgBox("Click OK to proceed", 0, "Done!")
    End Sub

    Private Sub commitpass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles commitpass2.Click
        sql = "SELECT * FROM pass"
        adaptor = New OleDb.OleDbDataAdapter(sql, cn)
        adaptor.Fill(ds, "getpass")
        builder = New OleDb.OleDbCommandBuilder(adaptor)
        ds.Tables("getpass").Rows(0).Item(0) = topass2.Text
        adaptor.Update(ds, "getpass")
        MsgBox("Click OK to proceed", 0, "Done!")
    End Sub

    Private Sub pass2_Validated(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If pass2.Text = ds.Tables("getpass").Rows(0).Item(0) Then
            txt1.Visible = True
            txt2.Visible = True
            txt3.Visible = True
            txt4.Visible = True
            txt5.Visible = True
            txt6.Visible = True
            txt7.Visible = True
            txt8.Visible = True
            txt9.Visible = True
            txt10.Visible = True
            txt11.Visible = True
            txt12.Visible = True
            txt13.Visible = True
            txt14.Visible = True
            txt15.Visible = True
            txt16.Visible = True
            txt17.Visible = True
            txt18.Visible = True
            txt19.Visible = True
            txt20.Visible = True
            txt21.Visible = True
            txt22.Visible = True
            create2.Visible = False
            result3.Visible = False
            pass2.Text = ""
            pass2.Visible = False
            edit2.Visible = False
            Label3.Visible = False
            commitcodes2.Visible = True
            back2.Visible = True
            LineShape1.Visible = True
            namenum.Visible = True
            changename2.Visible = True
            commitname.Visible = True
            RectangleShape1.Visible = True
            RectangleShape2.Visible = True
            RectangleShape3.Visible = True
            RectangleShape4.Visible = False
            commitpass2.Visible = True
            topass2.Visible = True
            Label5.Visible = True
            Label6.Visible = True
            Label7.Visible = True
            Label8.Visible = False
            t1.Visible = True
            t2.Visible = True
            t3.Visible = True
            t4.Visible = True
            t5.Visible = True
            t6.Visible = True
            t7.Visible = True
            t8.Visible = True
            t9.Visible = True
            t10.Visible = True
            t11.Visible = True
            t12.Visible = True
            t13.Visible = True
            t14.Visible = True
            t15.Visible = True
            t16.Visible = True
            t17.Visible = True
            t18.Visible = True
            t19.Visible = True
            t20.Visible = True
            t21.Visible = True
            t22.Visible = True
        End If
    End Sub

    Private Sub commitname_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles commitname.Click
        sql = "SELECT * FROM codes"
        adaptor = New OleDb.OleDbDataAdapter(sql, cn)
        adaptor.Fill(ds, "getcodes")
        builder = New OleDb.OleDbCommandBuilder(adaptor)
        checkboxes(CInt(namenum.Text)).Text = changename2.Text
        ds.Tables("getcodes").Rows(CInt(namenum.Text) - 1).Item(1) = changename2.Text
        adaptor.Update(ds, "getcodes")
        MsgBox("Click OK to proceed", 0, "Done!")
    End Sub
End Class
