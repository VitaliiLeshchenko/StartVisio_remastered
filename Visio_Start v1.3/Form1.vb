Public Class Form1
    Public oraconn
    Public orarec
    Public sqlz
    Public ind
    Public a, b, c, d, e, f, g, h

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'відкриваю форму на увесь екран
        Me.WindowState = FormWindowState.Maximized
        'вставляю полоску прокрутки вертикальну
        Me.HorizontalScroll.Maximum = 0
        Me.AutoScroll = True
        initialize_connection()
        обновить_все()
    End Sub

    Private Sub initialize_connection()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open(My.Resources.db_path)
    End Sub


    Private Sub refresh_Неупакованные_лотки()
        DataGridView1.Rows.Clear()
        orarec.Open(My.Resources.sql01_неупаковані_лотки, oraconn)
        ind = 0
        Do Until orarec.EOF

            DataGridView1.Rows.Add(orarec.Fields("day2").Value,
                                   orarec.Fields("dep").Value,
                                   orarec.Fields("not_lotok").Value,
                                   orarec.Fields("lotok").Value,
                                   orarec.Fields("otobr_not_upak").Value,
                                   orarec.Fields("otobr_upak").Value,
                                   orarec.Fields("vsego").Value)
            If orarec.Fields("prosr").Value = 1 Then DataGridView1.Rows(ind).DefaultCellStyle.BackColor = Color.Yellow
            If orarec.Fields("dep").Value Is DBNull.Value Then DataGridView1.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
            ind = ind + 1
            orarec.MoveNext()
        Loop
        Me.DataGridView1.Refresh()
        DataGridView1.AllowUserToAddRows = False
        orarec.Close()
    End Sub
    Private Sub refresh_Зона_отбора_APL()
        'Зона отбора APL **************************************************
        DataGridView3.Rows.Clear()

        orarec.Open(My.Resources.sql02_Зона_отбора_APL, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF
            a = orarec.Fields("bl").Value
            b = orarec.Fields("zap").Value
            c = orarec.Fields("dep").Value
            DataGridView3.Rows.Add(a, b, c)
            d = d + a
            e = e + b
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView3.Rows.Add(d, e, "Всего") : DataGridView3.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView3.Refresh()
        DataGridView3.AllowUserToAddRows = False
        orarec.Close()

        '**********************************************************************
    End Sub
    Private Sub refresh_Зона_отбора_CON()
        DataGridView4.Rows.Clear()
        orarec.Open(My.Resources.sql03_Зона_отбора_CON, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF
            a = orarec.Fields("bl").Value
            b = orarec.Fields("zap").Value
            c = orarec.Fields("dep").Value
            DataGridView4.Rows.Add(a, b, c)
            d = d + a
            e = e + b
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView4.Rows.Add(d, e, "Всего") : DataGridView4.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView4.Refresh()
        DataGridView4.AllowUserToAddRows = False
        orarec.Close()
    End Sub
    Private Sub refresh_Зона_отбора_NKZ_TRZ()
        'Зона отбора NKZ, TRZ **************************************************
        TextBox5.Text = ""
        TextBox7.Text = ""
        orarec.Open(My.Resources.sql04_Зона_отбора_NKZ, oraconn)
        a = orarec.Fields("a").Value
        TextBox5.Text = a
        orarec.close()
        orarec.Open(My.Resources.sql05_Зона_отбора_TRZ, oraconn)
        a = orarec.Fields("a").Value
        TextBox7.Text = a
        orarec.Close()
        '**********************************************************************
    End Sub
    Private Sub refresh_Зона_отбора_IVS()
        'Зона отбора IVS **************************************************
        DataGridView2.Rows.Clear()
        orarec.Open(My.Resources.sql06_Зона_отбора_IVS, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF
            a = orarec.Fields("bl").Value
            b = orarec.Fields("zap").Value
            c = orarec.Fields("dep").Value
            DataGridView2.Rows.Add(a, b, c)
            d = d + a
            e = e + b
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView2.Rows.Add(d, e, "Всего") : DataGridView2.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView2.Refresh()
        DataGridView2.AllowUserToAddRows = False
        orarec.Close()
        '**********************************************************************
    End Sub
    Private Sub refresh_Зона_отбора_IV()
        'Зона отбора IV **************************************************
        DataGridView6.Rows.Clear()

        orarec.Open(My.Resources.sql07_Зона_отбора_IV, oraconn)
        Do Until orarec.EOF
            a = orarec.Fields("area").Value
            b = orarec.Fields("zad").Value
            c = orarec.Fields("dep").Value
            DataGridView6.Rows.Add(a, b, c)
            orarec.MoveNext()
        Loop
        Me.DataGridView6.Refresh()
        DataGridView6.AllowUserToAddRows = False
        orarec.Close()
        '**********************************************************************
    End Sub
    Private Sub refresh_Пополнение_под_ЗО()
        'Пополнение под ЗО **************************************************
        DataGridView5.Rows.Clear()
        orarec.Open(My.Resources.sql08_Пополнение_под_ЗО, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF
            a = orarec.Fields("bl").Value
            b = orarec.Fields("zap").Value
            c = orarec.Fields("dep").Value
            DataGridView5.Rows.Add(a, b, c)
            d = d + a
            e = e + b
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView5.Rows.Add(d, e, "Всего") : DataGridView5.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView5.Refresh()
        DataGridView5.AllowUserToAddRows = False
        orarec.Close()
        '**********************************************************************
    End Sub
    Private Sub refresh_Кол_во_лотков_на_станциях_мезонина()
        DataGridView7.Rows.Clear()
        orarec.Open(My.Resources.sql09_Кол_во_лотков_на_станциях_мезонина, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF
            a = orarec.Fields("st").Value
            b = orarec.Fields("lototb").Value
            c = orarec.Fields("lot_popoln").Value
            DataGridView7.Rows.Add(a, b, c)
            If a Like "*8" Then DataGridView7.Rows(ind).DefaultCellStyle.BackColor = Color.Yellow
            d = d + b
            e = e + c
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView7.Rows.Add("Итого:", d, e) : DataGridView7.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView7.Refresh()
        orarec.Close()
    End Sub
    Private Sub refresh_Задания_в_ПС()
        'Задания в ПС **************************************************
        DataGridView8.Rows.Clear()
        orarec.Open(My.Resources.sql10_Задания_в_ПС, oraconn)
        Do Until orarec.EOF
            a = orarec.Fields("pl").Value
            b = orarec.Fields("vsg").Value
            c = orarec.Fields("blk").Value
            d = orarec.Fields("vip").Value
            e = orarec.Fields("wrk").Value
            f = orarec.Fields("paus").Value
            g = orarec.Fields("st40").Value
            DataGridView8.Rows.Add(a, b, c, d, e, f, g)
            orarec.MoveNext()
        Loop
        Me.DataGridView8.Refresh()
        DataGridView8.AllowUserToAddRows = False
        orarec.Close()
        '**********************************************************************
    End Sub
    Private Sub refresh_Дистро()
        DataGridView9.Rows.Clear()
        orarec.Open(My.Resources.sql11_дистро, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF
            a = orarec.Fields("sku").Value
            b = orarec.Fields("in_di").Value
            c = orarec.Fields("br").Value
            DataGridView9.Rows.Add(a, b, c)
            d = d + a
            e = e + b
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView9.Rows.Add(d, e, "Всего") : DataGridView9.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView9.Refresh()
        orarec.Close()
    End Sub
    Private Sub refresh_Button_Незапущені_ЗО()
        DGV_Незапущені_ЗО.Rows.Clear()
        orarec.Open(My.Resources.sql14_незапущенные_ЗО, oraconn)
        Do Until orarec.EOF
            DGV_Незапущені_ЗО.Rows.Add(orarec.fields(0).value, orarec.fields(1).value)
            orarec.MoveNext()
        Loop
        orarec.Close()
    End Sub
    Private Sub refresh_Відвантажено_за_сьогодні()
        DGV_shipped_today.Rows.Clear()
        orarec.Open(My.Resources.sql15_відвантажено_за_сьогодні, oraconn)
        Do Until orarec.EOF
            DGV_shipped_today.Rows.Add(orarec.Fields("departament").Value,
                                                  orarec.Fields(1).value,
                                                  orarec.Fields(2).value,
                                                  orarec.Fields(3).value)
            orarec.MoveNext()
        Loop
        orarec.Close()
    End Sub


    Private Sub Button_Незапущені_ЗО_Click(sender As Object, e As EventArgs) Handles Button_Незапущені_ЗО.Click
        refresh_Button_Незапущені_ЗО()
    End Sub
    Private Sub Button_Відвантажено_за_сьогодні_Click(sender As Object, e As EventArgs) Handles Button_Відвантажено_за_сьогодні.Click
        refresh_Відвантажено_за_сьогодні()
    End Sub

    Public Sub обновить_все()
        Form1.ActiveForm.Refresh()
        Form1.ActiveForm.Text = "ОНОВЛЯЮ ДАНІ"
        refresh_Неупакованные_лотки()
        refresh_Зона_отбора_APL()
        refresh_Зона_отбора_CON()
        refresh_Зона_отбора_NKZ_TRZ()
        refresh_Зона_отбора_IVS()
        refresh_Зона_отбора_IV()
        refresh_Пополнение_под_ЗО()
        refresh_Кол_во_лотков_на_станциях_мезонина()
        refresh_Задания_в_ПС()
        refresh_Дистро()
        refresh_Button_Незапущені_ЗО()
        refresh_Відвантажено_за_сьогодні()
        Form1.ActiveForm.Text = "StartVisio | обновлено:" & Now
    End Sub

    Private Sub Button_refresh_all_Click(sender As Object, e As EventArgs) Handles Button_refresh_all.Click
        обновить_все()
    End Sub

    Private Sub Close_connection()
        oraconn.close()
    End Sub

    Private Sub Form1_Closing(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Close_connection()
        oraconn = Nothing
        orarec = Nothing
    End Sub
End Class