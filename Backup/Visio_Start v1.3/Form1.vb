Public Class Form1
    Public oraconn
    Public orarec
    Public sqlz
    Public ind
    Public a, b, c, d, e, f, g, h
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Location = New Point(Screen.PrimaryScreen.WorkingArea.Width - Me.Width, 0)
        'Me.WindowState = FormWindowState.Maximized
        sql()
        Timer1.Interval = 30000
    End Sub
    Public Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.Text = "Visio_Start v1.3   [Обновлено: " & Now() & "]"
        Sql()
    End Sub
    Public Sub sql()

        'Неупакованные лотки **************************************************
        DataGridView1.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select day2, dep, sum(not_lotok) not_lotok, sum(lotok) lotok, sum(otobr_upak) otobr_upak, sum(otobr_not_upak) otobr_not_upak, sum(not_lotok)+sum(lotok)+sum(otobr_upak)+sum(otobr_not_upak) vsego, prosr from( select day1, day2, dep, lot, count(locn) strok, sum(qty) qty, sum(ltr) ltr, not_lotok, lotok, otobr_upak, otobr_not_upak, prosr from (select distinct td.task_id lot, td.task_seq_nbr, to_char(th.mod_date_time,'YYYY-MM-DD') day1, to_char(th.mod_date_time,'DD.MM.YYYY') day2 "

        sqlz = sqlz + ",case when to_char(th.mod_date_time,'DD.MM.YYYY') <> to_char(sysdate,'DD.MM.YYYY') then 1 else 0 end prosr"

        sqlz = sqlz + ",case when im.cd_master_id = '3001' then 'MTI СБ' when im.cd_master_id = '19004' then 'DD' when im.cd_master_id = '3002' then '+ Маркет' when im.cd_master_id = '3003' then 'Альпина' when im.cd_master_id = '4003' then 'БНС Трейд' when im.cd_master_id = '4004' then 'БНС Кампани' when im.cd_master_id = '5003' then 'Мазеркер' when im.cd_master_id = '5004' then 'Форс' when im.cd_master_id = '6003' then case when im.sale_grp = 'P01' then 'LOR' when im.sale_grp is null and substr(im.sku_desc, 1, 3)= 'LOR' then 'LOR' "
        sqlz = sqlz + "when ph.shipto_addr_1 not like '%нтернет%' and ph.shipto_addr_1 not like '%ахтерск%' and ph.shipto_addr_1 not like '%АХТЕРСК%' and phi.total_nbr_of_units > 4 then 'PROT' Else 'PROT_IM' end when im.cd_master_id = '7003' then 'PWA' when im.cd_master_id = '8003' then 'БПИ' when im.cd_master_id = '9003' then 'Форс' when im.cd_master_id = '9004' then 'Силд Эйр' when im.cd_master_id = '9005' then 'Тека' when im.cd_master_id = '9006' then 'Легранд' when im.cd_master_id = '10003' then 'Сабриз' when im.cd_master_id = '10004' then 'Альфамарис' when im.cd_master_id = '11004' then 'Юнивест' when im.cd_master_id = '11005' then 'Калугин' when im.cd_master_id = '15004' then 'Смартмакс' when im.cd_master_id = '18005' then 'Орбико' when im.cd_master_id = '2001' then case when im.size_desc = '1214119' and ph.shipto_name like '%Плато%' then 'PLATO' when im.size_desc = 'UPAK-DOO' then 'DOO' when ph.vendor_nbr like '%W60%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' when ph.vendor_nbr like '%W63%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' "
        sqlz = sqlz + "when ph.vendor_nbr like '%W60%' then 'DD' when ph.vendor_nbr like '%W63%' then 'DD' when ph.vendor_nbr like '%W640%' then 'DOO' when ph.vendor_nbr like '%W650%' then 'DOO' when ph.vendor_nbr like '%W680%' then 'DOO' when ph.vendor_nbr like '%MITIN01%' then 'INET' when ph.vendor_nbr like '%MPLIN01%' then 'INET' when ph.vendor_nbr like '%MECIN01%' then 'INET' Else Case when 'P' in (select 'P' from pkt_dtl ad1, item_master im1 where im1.sku_id = ad1.sku_id and ph.pkt_ctrl_nbr = ad1.pkt_ctrl_nbr and substr(im1.sale_grp, 0, 1) = 'P') then 'P/D' else case when im.sale_grp = '      ' or substr(im.sale_grp,0,1)='T' or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) in ('CAM','CLR','GEX','LOB','VGB')) then 'DOO' when im.sale_grp = '   ' or (substr(im.sale_grp,0,1)='D' and im.sale_grp <> 'DRM') or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) not in ('CAM','CLR','GEX','LOB','VGB')) or im.sale_grp in ('P00','PPK','_PROT','PRO') then 'DD' when im.sale_grp in ('_DL','_TSOI','_BASHL') then 'DL' "
        sqlz = sqlz + "Else 'OTHER' end end end end dep ,im.size_desc, lf.dsp_locn locn, round(td.qty_pulld) qty, round(im.unit_vol*td.qty_pulld/1000,2) ltr ,case when td.stat_code < '90' and cu.cntr_nbr is null then 1 else 0 end not_lotok ,case when td.stat_code < '90' and cu.cntr_nbr is not null then 1 else 0 end lotok ,case when td.stat_code = '90' and ch.stat_code >= '20' then 1 else 0 end otobr_upak ,case when td.stat_code = '90' and ch.stat_code < '20' then 1 else 0 end otobr_not_upak from task_dtl td join item_master im on im.sku_id=td.sku_id join locn_hdr lf on lf.locn_id=td.pull_locn_id join task_hdr th on th.task_id = td.task_id join pkt_hdr ph on ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr join pkt_hdr_intrnl phi on phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr join c_umti_mhe_cntr cu on cu.task_id = th.task_id join carton_dtl cd on cd.carton_seq_nbr = td.carton_seq_nbr join carton_hdr ch on ch.carton_nbr=cd.carton_nbr and ch.pkt_ctrl_nbr = phi.pkt_ctrl_nbr "
        sqlz = sqlz + "join(select distinct aid50.alloc_invn_dtl_id,aid52.carton_seq_nbr,aid52.cntr_nbr as tote_nbr,aid52.carton_nbr from alloc_invn_dtl aid50 join alloc_invn_dtl aid52 on aid50.carton_nbr=aid52.carton_nbr)aid on aid.carton_seq_nbr=cd.carton_seq_nbr and aid.carton_nbr=cd.carton_nbr and aid.alloc_invn_dtl_id = td.alloc_invn_dtl_id where th.task_desc in ('Отбор с мезонина', 'Упаковка на мезонине','Pick pack MEZ') and substr(lf.dsp_locn,1,1) in ('K','L','M','N') and th.stat_code <> '99' and (to_char(th.mod_date_time,'DD-MM-YYYY')=to_char(sysdate,'DD-MM-YYYY') or th.stat_code = '10') )group by day1, day2, dep, lot, not_lotok, lotok, otobr_upak, otobr_not_upak, prosr )group by rollup (dep),day2,prosr order by 2,1"
        orarec.Open(sqlz, oraconn)
        ind = 0
        Do Until orarec.EOF
            a = orarec.Fields("day2").Value
            b = orarec.Fields("dep").Value
            c = orarec.Fields("not_lotok").Value
            d = orarec.Fields("lotok").Value
            e = orarec.Fields("otobr_not_upak").Value
            f = orarec.Fields("otobr_upak").Value
            g = orarec.Fields("vsego").Value
            h = orarec.Fields("prosr").Value
            DataGridView1.Rows.Add(a, b, c, d, e, f, g)
            If h = 1 Then DataGridView1.Rows(ind).DefaultCellStyle.BackColor = Color.Yellow
            If b Is DBNull.Value Then DataGridView1.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
            ind = ind + 1
            orarec.MoveNext()
        Loop
        Me.DataGridView1.Refresh()
        DataGridView1.AllowUserToAddRows = False
        orarec.Close()
        oraconn.close()
        '**********************************************************************



        'Зона отбора APL **************************************************
        DataGridView3.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select sum(bl) bl, sum(zap) zap, dep from (select distinct th.task_id "

        sqlz = sqlz + ",case when im.cd_master_id = '3001' then 'MTI СБ' when im.cd_master_id = '19004' then 'DD' when im.cd_master_id = '3002' then '+ Маркет' when im.cd_master_id = '3003' then 'Альпина' when im.cd_master_id = '4003' then 'БНС Трейд' when im.cd_master_id = '4004' then 'БНС Кампани' when im.cd_master_id = '5003' then 'Мазеркер' when im.cd_master_id = '5004' then 'Форс' when im.cd_master_id = '6003' then case when im.sale_grp = 'P01' then 'LOR' when im.sale_grp is null and substr(im.sku_desc, 1, 3)= 'LOR' then 'LOR' "
        sqlz = sqlz + "when ph.shipto_addr_1 not like '%нтернет%' and ph.shipto_addr_1 not like '%ахтерск%' and ph.shipto_addr_1 not like '%АХТЕРСК%' and phi.total_nbr_of_units > 4 then 'PROT' Else 'PROT_IM' end when im.cd_master_id = '7003' then 'PWA' when im.cd_master_id = '8003' then 'БПИ' when im.cd_master_id = '9003' then 'Форс' when im.cd_master_id = '9004' then 'Силд Эйр' when im.cd_master_id = '9005' then 'Тека' when im.cd_master_id = '9006' then 'Легранд' when im.cd_master_id = '10003' then 'Сабриз' when im.cd_master_id = '10004' then 'Альфамарис' when im.cd_master_id = '11004' then 'Юнивест' when im.cd_master_id = '11005' then 'Калугин' when im.cd_master_id = '2001' then case when im.size_desc = '1214119' and ph.shipto_name like '%Плато%' then 'PLATO' when im.size_desc = 'UPAK-DOO' then 'DOO' when ph.vendor_nbr like '%W60%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' when ph.vendor_nbr like '%W63%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' "
        sqlz = sqlz + "when ph.vendor_nbr like '%W60%' then 'DD' when ph.vendor_nbr like '%W63%' then 'DD' when ph.vendor_nbr like '%W640%' then 'DOO' when ph.vendor_nbr like '%W650%' then 'DOO' when ph.vendor_nbr like '%W680%' then 'DOO' when ph.vendor_nbr like '%MITIN01%' then 'INET' when ph.vendor_nbr like '%MPLIN01%' then 'INET' when ph.vendor_nbr like '%MECIN01%' then 'INET' Else Case when 'P' in (select 'P' from pkt_dtl ad1, item_master im1 where im1.sku_id = ad1.sku_id and ph.pkt_ctrl_nbr = ad1.pkt_ctrl_nbr and substr(im1.sale_grp, 0, 1) = 'P') then 'P/D' else case when im.sale_grp = '      ' or substr(im.sale_grp,0,1)='T' or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) in ('CAM','CLR','GEX','LOB','VGB')) then 'DOO' when im.sale_grp = '   ' or (substr(im.sale_grp,0,1)='D' and im.sale_grp <> 'DRM') or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) not in ('CAM','CLR','GEX','LOB','VGB')) or im.sale_grp in ('P00','PPK','_PROT','PRO') then 'DD' when im.sale_grp in ('_DL','_TSOI','_BASHL') then 'DL' "
        sqlz = sqlz + "Else 'OTHER' end end end end dep "

        sqlz = sqlz + ",case when th.stat_code = '5' then 1 else 0 end bl ,case when th.stat_code = '5' then 0 else 1 end zap  from task_hdr th join task_dtl td on td.task_id = th.task_id join locn_hdr lh on lh.locn_id = td.pull_locn_id join item_master im on im.sku_id = td.sku_id join pkt_hdr ph on ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr join pkt_hdr_intrnl phi on phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr where th.task_desc in ('Отбор с А-паллет') and substr(lh.area,1,1)='A' and th.stat_code < '90') group by dep order by 3"
        orarec.Open(sqlz, oraconn)
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
        oraconn.close()
        '**********************************************************************


        'Зона отбора CON **************************************************
        DataGridView4.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select sum(bl) bl, sum(zap) zap, dep from (select distinct th.task_id "

        sqlz = sqlz + ",case when im.cd_master_id = '3001' then 'MTI СБ' when im.cd_master_id = '19004' then 'DD' when im.cd_master_id = '3002' then '+ Маркет' when im.cd_master_id = '3003' then 'Альпина' when im.cd_master_id = '4003' then 'БНС Трейд' when im.cd_master_id = '4004' then 'БНС Кампани' when im.cd_master_id = '5003' then 'Мазеркер' when im.cd_master_id = '5004' then 'Форс' when im.cd_master_id = '6003' then case when im.sale_grp = 'P01' then 'LOR' when im.sale_grp is null and substr(im.sku_desc, 1, 3)= 'LOR' then 'LOR' "
        sqlz = sqlz + "when ph.shipto_addr_1 not like '%нтернет%' and ph.shipto_addr_1 not like '%ахтерск%' and ph.shipto_addr_1 not like '%АХТЕРСК%' and phi.total_nbr_of_units > 4 then 'PROT' Else 'PROT_IM' end when im.cd_master_id = '7003' then 'PWA' when im.cd_master_id = '8003' then 'БПИ' when im.cd_master_id = '9003' then 'Форс' when im.cd_master_id = '9004' then 'Силд Эйр' when im.cd_master_id = '9005' then 'Тека' when im.cd_master_id = '9006' then 'Легранд' when im.cd_master_id = '10003' then 'Сабриз' when im.cd_master_id = '10004' then 'Альфамарис' when im.cd_master_id = '11004' then 'Юнивест' when im.cd_master_id = '11005' then 'Калугин' when im.cd_master_id = '2001' then case when im.size_desc = '1214119' and ph.shipto_name like '%Плато%' then 'PLATO' when im.size_desc = 'UPAK-DOO' then 'DOO' when ph.vendor_nbr like '%W60%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' when ph.vendor_nbr like '%W63%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' "
        sqlz = sqlz + "when ph.vendor_nbr like '%W60%' then 'DD' when ph.vendor_nbr like '%W63%' then 'DD' when ph.vendor_nbr like '%W640%' then 'DOO' when ph.vendor_nbr like '%W650%' then 'DOO' when ph.vendor_nbr like '%W680%' then 'DOO' when ph.vendor_nbr like '%MITIN01%' then 'INET' when ph.vendor_nbr like '%MPLIN01%' then 'INET' when ph.vendor_nbr like '%MECIN01%' then 'INET' Else Case when 'P' in (select 'P' from pkt_dtl ad1, item_master im1 where im1.sku_id = ad1.sku_id and ph.pkt_ctrl_nbr = ad1.pkt_ctrl_nbr and substr(im1.sale_grp, 0, 1) = 'P') then 'P/D' else case when im.sale_grp = '      ' or substr(im.sale_grp,0,1)='T' or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) in ('CAM','CLR','GEX','LOB','VGB')) then 'DOO' when im.sale_grp = '   ' or (substr(im.sale_grp,0,1)='D' and im.sale_grp <> 'DRM') or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) not in ('CAM','CLR','GEX','LOB','VGB')) or im.sale_grp in ('P00','PPK','_PROT','PRO') then 'DD' when im.sale_grp in ('_DL','_TSOI','_BASHL') then 'DL' "
        sqlz = sqlz + "Else 'OTHER' end end end end dep "

        sqlz = sqlz + ",case when th.stat_code = '5' then 1 else 0 end bl ,case when th.stat_code = '5' then 0 else 1 end zap  from task_hdr th join task_dtl td on td.task_id = th.task_id join locn_hdr lh on lh.locn_id = td.pull_locn_id join item_master im on im.sku_id = td.sku_id and im.cd_master_id <> '10003' join pkt_hdr ph on ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr join pkt_hdr_intrnl phi on phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr where th.task_desc in ('Отбор из ПС','Отбор из ПС паллет','Отбор из ПС Ростовки','Отбор LPN из ПС') and substr(lh.area,1,1) in ('P','X') and th.stat_code < '90') group by dep order by 3"
        orarec.Open(sqlz, oraconn)
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
        oraconn.close()
        '**********************************************************************


        'Зона отбора NKZ, TRZ **************************************************
        TextBox5.Text = ""
        TextBox7.Text = ""
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select count(*) a from (select distinct th.task_id from task_dtl td join task_hdr th on th.task_id = td.task_id and th.stat_code < '90' join locn_hdr lh on lh.locn_id = td.pull_locn_id and lh.area = 'NKZ')"
        orarec.Open(sqlz, oraconn)
        a = orarec.Fields("a").Value
        TextBox5.Text = a
        orarec.close()
        sqlz = "select count(*) a from (select distinct th.task_id from task_dtl td join task_hdr th on th.task_id = td.task_id and th.stat_code < '90' join locn_hdr lh on lh.locn_id = td.pull_locn_id and lh.area = 'TRZ')"
        orarec.Open(sqlz, oraconn)
        a = orarec.Fields("a").Value
        TextBox7.Text = a
        orarec.Close()
        oraconn.close()
        '**********************************************************************


        'Зона отбора IVS **************************************************
        DataGridView2.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select sum(bl) bl, sum(zap) zap, dep from (select distinct th.task_id "

        sqlz = sqlz + ",case when im.cd_master_id = '3001' then 'MTI СБ' when im.cd_master_id = '19004' then 'DD' when im.cd_master_id = '3002' then '+ Маркет' when im.cd_master_id = '3003' then 'Альпина' when im.cd_master_id = '4003' then 'БНС Трейд' when im.cd_master_id = '4004' then 'БНС Кампани' when im.cd_master_id = '5003' then 'Мазеркер' when im.cd_master_id = '5004' then 'Форс' when im.cd_master_id = '6003' then case when im.sale_grp = 'P01' then 'LOR' when im.sale_grp is null and substr(im.sku_desc, 1, 3)= 'LOR' then 'LOR' "
        sqlz = sqlz + "when ph.shipto_addr_1 not like '%нтернет%' and ph.shipto_addr_1 not like '%ахтерск%' and ph.shipto_addr_1 not like '%АХТЕРСК%' and phi.total_nbr_of_units > 4 then 'PROT' Else 'PROT_IM' end when im.cd_master_id = '7003' then 'PWA' when im.cd_master_id = '8003' then 'БПИ' when im.cd_master_id = '9003' then 'Форс' when im.cd_master_id = '9004' then 'Силд Эйр' when im.cd_master_id = '9005' then 'Тека' when im.cd_master_id = '9006' then 'Легранд' when im.cd_master_id = '10003' then 'Сабриз' when im.cd_master_id = '10004' then 'Альфамарис' when im.cd_master_id = '11004' then 'Юнивест' when im.cd_master_id = '11005' then 'Калугин' when im.cd_master_id = '2001' then case when im.size_desc = '1214119' and ph.shipto_name like '%Плато%' then 'PLATO' when im.size_desc = 'UPAK-DOO' then 'DOO' when ph.vendor_nbr like '%W60%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' when ph.vendor_nbr like '%W63%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' "
        sqlz = sqlz + "when ph.vendor_nbr like '%W60%' then 'DD' when ph.vendor_nbr like '%W63%' then 'DD' when ph.vendor_nbr like '%W640%' then 'DOO' when ph.vendor_nbr like '%W650%' then 'DOO' when ph.vendor_nbr like '%W680%' then 'DOO' when ph.vendor_nbr like '%MITIN01%' then 'INET' when ph.vendor_nbr like '%MPLIN01%' then 'INET' when ph.vendor_nbr like '%MECIN01%' then 'INET' Else Case when 'P' in (select 'P' from pkt_dtl ad1, item_master im1 where im1.sku_id = ad1.sku_id and ph.pkt_ctrl_nbr = ad1.pkt_ctrl_nbr and substr(im1.sale_grp, 0, 1) = 'P') then 'P/D' else case when im.sale_grp = '      ' or substr(im.sale_grp,0,1)='T' or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) in ('CAM','CLR','GEX','LOB','VGB')) then 'DOO' when im.sale_grp = '   ' or (substr(im.sale_grp,0,1)='D' and im.sale_grp <> 'DRM') or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) not in ('CAM','CLR','GEX','LOB','VGB')) or im.sale_grp in ('P00','PPK','_PROT','PRO') then 'DD' when im.sale_grp in ('_DL','_TSOI','_BASHL') then 'DL' "
        sqlz = sqlz + "Else 'OTHER' end end end end dep "

        sqlz = sqlz + ",case when th.stat_code = '5' then 1 else 0 end bl ,case when th.stat_code = '5' then 0 else 1 end zap  from task_hdr th join task_dtl td on td.task_id = th.task_id join locn_hdr lh on lh.locn_id = td.pull_locn_id join item_master im on im.sku_id = td.sku_id and im.cd_master_id <> '10003' join pkt_hdr ph on ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr join pkt_hdr_intrnl phi on phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr where th.begin_area='IVS' and th.stat_code < '90') group by dep order by 3"
        orarec.Open(sqlz, oraconn)
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
        oraconn.close()
        '**********************************************************************


        'Зона отбора IV **************************************************
        DataGridView6.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select area, count(tsk) zad, dep from (select distinct th.task_id tsk, lh.area area "

        sqlz = sqlz + ",case when im.cd_master_id = '3001' then 'MTI СБ' when im.cd_master_id = '19004' then 'DD' when im.cd_master_id = '3002' then '+ Маркет' when im.cd_master_id = '3003' then 'Альпина' when im.cd_master_id = '4003' then 'БНС Трейд' when im.cd_master_id = '4004' then 'БНС Кампани' when im.cd_master_id = '5003' then 'Мазеркер' when im.cd_master_id = '5004' then 'Форс' when im.cd_master_id = '6003' then case when im.sale_grp = 'P01' then 'LOR' when im.sale_grp is null and substr(im.sku_desc, 1, 3)= 'LOR' then 'LOR' "
        sqlz = sqlz + "when ph.shipto_addr_1 not like '%нтернет%' and ph.shipto_addr_1 not like '%ахтерск%' and ph.shipto_addr_1 not like '%АХТЕРСК%' and phi.total_nbr_of_units > 4 then 'PROT' Else 'PROT_IM' end when im.cd_master_id = '7003' then 'PWA' when im.cd_master_id = '8003' then 'БПИ' when im.cd_master_id = '9003' then 'Форс' when im.cd_master_id = '9004' then 'Силд Эйр' when im.cd_master_id = '9005' then 'Тека' when im.cd_master_id = '9006' then 'Легранд' when im.cd_master_id = '10003' then 'Сабриз' when im.cd_master_id = '10004' then 'Альфамарис' when im.cd_master_id = '11004' then 'Юнивест' when im.cd_master_id = '11005' then 'Калугин' when im.cd_master_id = '2001' then case when im.size_desc = '1214119' and ph.shipto_name like '%Плато%' then 'PLATO' when im.size_desc = 'UPAK-DOO' then 'DOO' when ph.vendor_nbr like '%W60%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' when ph.vendor_nbr like '%W63%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' "
        sqlz = sqlz + "when ph.vendor_nbr like '%W60%' then 'DD' when ph.vendor_nbr like '%W63%' then 'DD' when ph.vendor_nbr like '%W640%' then 'DOO' when ph.vendor_nbr like '%W650%' then 'DOO' when ph.vendor_nbr like '%W680%' then 'DOO' when ph.vendor_nbr like '%MITIN01%' then 'INET' when ph.vendor_nbr like '%MPLIN01%' then 'INET' when ph.vendor_nbr like '%MECIN01%' then 'INET' Else Case when 'P' in (select 'P' from pkt_dtl ad1, item_master im1 where im1.sku_id = ad1.sku_id and ph.pkt_ctrl_nbr = ad1.pkt_ctrl_nbr and substr(im1.sale_grp, 0, 1) = 'P') then 'P/D' else case when im.sale_grp = '      ' or substr(im.sale_grp,0,1)='T' or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) in ('CAM','CLR','GEX','LOB','VGB')) then 'DOO' when im.sale_grp = '   ' or (substr(im.sale_grp,0,1)='D' and im.sale_grp <> 'DRM') or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) not in ('CAM','CLR','GEX','LOB','VGB')) or im.sale_grp in ('P00','PPK','_PROT','PRO') then 'DD' when im.sale_grp in ('_DL','_TSOI','_BASHL') then 'DL' "
        sqlz = sqlz + "Else 'OTHER' end end end end dep "

        sqlz = sqlz + "from task_hdr th join task_dtl td on td.task_id = th.task_id join locn_hdr lh on lh.locn_id = td.pull_locn_id join item_master im on im.sku_id = td.sku_id and im.cd_master_id <> '10003' join pkt_hdr ph on ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr join pkt_hdr_intrnl phi on phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr where substr(th.begin_area,1,2)='IV'  and th.begin_area <> 'IVS' and th.stat_code < '90') group by area, dep order by 1,3"
        orarec.Open(sqlz, oraconn)
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
        oraconn.close()
        '**********************************************************************






        'Пополнение под ЗО **************************************************
        DataGridView5.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select sum(bl) bl, sum(zap) zap, dep from (select distinct th.task_id "

        sqlz = sqlz + ",case when im.cd_master_id = '3001' then 'MTI СБ' when im.cd_master_id = '19004' then 'DD' when im.cd_master_id = '3002' then '+ Маркет' when im.cd_master_id = '3003' then 'Альпина' when im.cd_master_id = '4003' then 'БНС Трейд' when im.cd_master_id = '4004' then 'БНС Кампани' when im.cd_master_id = '5003' then 'Мазеркер' when im.cd_master_id = '5004' then 'Форс' when im.cd_master_id = '6003' then case when im.sale_grp = 'P01' then 'LOR' when im.sale_grp is null and substr(im.sku_desc, 1, 3)= 'LOR' then 'LOR' Else 'PROT_IM' end "
        sqlz = sqlz + "when im.cd_master_id = '7003' then 'PWA' when im.cd_master_id = '8003' then 'БПИ' when im.cd_master_id = '9003' then 'Форс' when im.cd_master_id = '9004' then 'Силд Эйр' when im.cd_master_id = '9005' then 'Тека' when im.cd_master_id = '9006' then 'Легранд' when im.cd_master_id = '10003' then 'Сабриз' when im.cd_master_id = '10004' then 'Альфамарис' when im.cd_master_id = '11004' then 'Юнивест' when im.cd_master_id = '11005' then 'Калугин' when im.cd_master_id = '2001' then case when im.size_desc = 'UPAK-DOO' then 'DOO'  "
        sqlz = sqlz + " Else case when im.sale_grp = '      ' or substr(im.sale_grp,0,1)='T' or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) in ('CAM','CLR','GEX','LOB','VGB')) then 'DOO' when im.sale_grp = '   ' or (substr(im.sale_grp,0,1)='D' and im.sale_grp <> 'DRM') or (im.sale_grp = 'DRM' and substr(im.sku_desc,0,3) not in ('CAM','CLR','GEX','LOB','VGB')) or im.sale_grp in ('P00','PPK','_PROT','PRO') then 'DD' when im.sale_grp in ('_DL','_TSOI','_BASHL') then 'DL' "
        sqlz = sqlz + "Else 'OTHER' end end end dep "

        sqlz = sqlz + ",case when th.start_curr_work_area <> 'MIS' then 1 else 0 end bl ,case when th.start_curr_work_area = 'MIS' then 1 else 0 end zap  from task_hdr th join task_dtl td on td.task_id = th.task_id join locn_hdr lh on lh.locn_id = td.pull_locn_id join item_master im on im.sku_id = td.sku_id and im.cd_master_id <> '10003' join case_hdr ch on ch.case_nbr = td.cntr_nbr and ch.plt_id is null where th.invn_need_type = 1 and th.end_curr_work_area <> 'APL' and th.end_curr_work_grp <> 'APL' and th.stat_code < '90') group by dep order by 3"
        orarec.Open(sqlz, oraconn)
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
        oraconn.close()
        '**********************************************************************



        ''Отбор из IV **************************************************
        'DataGridView6.Rows.Clear()
        'oraconn = CreateObject("ADODB.Connection")
        'orarec = CreateObject("ADODB.Recordset")
        'oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        'sqlz = "select area, count(zad) zad from (select distinct lh.area area, td.task_id zad from task_dtl td join locn_hdr lh on lh.locn_id = td.pull_locn_id where td.stat_code < '90' and substr(lh.area,1,2)='IV' and lh.area <> 'IVS') group by area order by 1"
        'orarec.Open(sqlz, oraconn)
        'Do Until orarec.EOF
        '    a = orarec.Fields("area").Value
        '    b = orarec.Fields("zad").Value
        '    DataGridView6.Rows.Add(a, b)
        '    orarec.MoveNext()
        'Loop
        'Me.DataGridView6.Refresh()
        'DataGridView6.AllowUserToAddRows = False
        'orarec.Close()
        'oraconn.close()
        ''**********************************************************************



        'Кол-во лотков на станциях мезонина **************************************************
        DataGridView7.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select st,max(lototb) lototb,max(lot_popoln) lot_popoln from (select st,count (*) lototb,0 lot_popoln from (select o.task_id, c.mod_date_time prisvoen, min (l.pick_detrm_zone) st from task_hdr o, task_dtl d, locn_hdr l, c_umti_mhe_cntr c where(o.task_id = d.task_id And d.pull_locn_id = l.locn_id And o.task_id = c.task_id And c.cntr_nbr Is Not null) and o.task_type='16' and d.stat_code<'90' group by o.task_id, c.mod_date_time ) a group by st union all select l.pick_detrm_zone st,0 lototb,count (*)lot_popoln from case_hdr ch, locn_hdr l where ch.dest_locn_id=l.locn_id and ch.stat_code<65 and ch.plt_id is not null and l.locn_class='A' group by l.pick_detrm_zone ) group by st order by 1"
        orarec.Open(sqlz, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF
            a = orarec.Fields("st").Value
            b = orarec.Fields("lototb").Value
            c = orarec.Fields("lot_popoln").Value
            DataGridView7.Rows.Add(a, b, c)
            d = d + b
            e = e + c
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView7.Rows.Add("Итого:", d, e) : DataGridView7.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView7.Refresh()
        DataGridView7.AllowUserToAddRows = False
        orarec.Close()
        oraconn.close()
        '**********************************************************************


        'Задания в ПС **************************************************
        DataGridView8.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select case when pl like 'A%L' then 'A_L' when pl like 'A%R' then 'A_R' else pl end pl,count(1) vsg,sum(blok) blk,sum(wait) vip,sum(process) wrk,max(prostoy) paus,sum(st40) st40 from(select th.task_id,th.start_curr_work_area PL,th.curr_task_prty pr,case when th.stat_code='5' then 1 else 0 end blok,case when th.stat_code>'5' and th.stat_code <20 then 1 else 0 end wait,case when th.stat_code>='20' then 1 else 0 end process,case when th.curr_task_prty = '40' then 1 else 0 end st40,round((sysdate-th.mod_date_time)*24) prostoy from task_hdr th join task_dtl td on td.task_id = th.task_id and td.cd_master_id <> '10003' where(th.stat_code < 90)and th.invn_need_type <>100 and th.task_type <> 09 and(th.start_curr_work_area like 'PL%' or (th.start_curr_work_area='XXL')or(th.start_curr_work_area like 'A%' and th.end_dest_work_area like 'A%')) )group by  case when pl like 'A%L' then 'A_L' when pl like 'A%R' then 'A_R' else pl end ORDER BY 1 DESC"
        orarec.Open(sqlz, oraconn)
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
        oraconn.close()
        '**********************************************************************


        'Дистро **************************************************
        DataGridView9.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select br, count(sd) sku, sum(in_di) in_di from (select im.size_desc sd,substr(im.sku_desc, 1, 3) br,sd.distro_nbr din,sm.store_nbr km,sm.name nm,a.city ci,round(sd.reqd_qty) in_di from store_distro sd join case_dtl cd on cd.sku_id = sd.sku_id and cd.prod_stat = '00' and sd.batch_nbr = nvl(cd.batch_nbr, '*') join case_hdr ch on ch.case_nbr = cd.case_nbr and ch.stat_code in ('10', '30', '45') and ch.locn_id not in (select l.locn_id from locn_hdr l where l.area = 'BLB') join item_master im on im.sku_id = sd.sku_id join store_master sm on sm.store_nbr = sd.store_nbr join address a on a.addr_id = sm.addr_id where sd.stat_code < '90' and sd.reqd_qty > 0 and (ch.rcvd_shpmt_nbr is null or ch.rcvd_shpmt_nbr in (select ah.shpmt_nbr from asn_hdr ah where ah.stat_code = '90')) group by im.size_desc,sd.distro_nbr,sd.reqd_qty,sm.store_nbr,sm.name,a.city,im.sku_desc Union all select im.size_desc sd,substr(im.sku_desc, 1, 3) br,sd.distro_nbr din,sm.store_nbr km,sm.name nm,a.city ci,round(sd.reqd_qty) in_di from store_distro sd join pick_locn_dtl pld on pld.sku_id = sd.sku_id join locn_hdr lh on lh.locn_id = pld.locn_id and lh.locn_class = 'A' join item_master im on im.sku_id = sd.sku_id join store_master sm on sm.store_nbr = sd.store_nbr join address a on a.addr_id = sm.addr_id join prod_trkg_tran ptt on ptt.sku_id = im.sku_id and ptt.tran_type = '100' where sd.stat_code < '90' and sd.reqd_qty > 0 group by im.size_desc,sd.distro_nbr,sd.reqd_qty,sm.store_nbr,sm.name,a.city,im.sku_desc )group by br order by 1"
        orarec.Open(sqlz, oraconn)
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
        DataGridView9.AllowUserToAddRows = False
        orarec.Close()
        oraconn.close()
        '**********************************************************************


        'Разбор ростовок **************************************************
        DataGridView10.Rows.Clear()
        DataGridView11.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select br, sum(sozd) sozd, sum(zap) zap from (select distinct woh.work_ord_nbr, case when woh.stat_code = '10' then 1 else 0 end sozd, case when woh.stat_code = '20' then 1 else 0 end zap,substr(im.sku_desc,1,3) br from work_ord_hdr woh join work_ord_dtl wod on wod.work_ord_nbr = woh.work_ord_nbr join item_master im on im.sku_id = wod.sku_id where woh.stat_code in ('10') )group by br order by 2 desc"
        orarec.Open(sqlz, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF Or ind = 23
            a = orarec.Fields("sozd").Value
            'b = orarec.Fields("zap").Value
            c = orarec.Fields("br").Value
            DataGridView10.Rows.Add(a, c)
            d = d + a
            'e = e + b
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView10.Rows.Add(d, "Всего") : DataGridView10.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView10.Refresh()
        DataGridView10.AllowUserToAddRows = False
        orarec.Close()
        sqlz = "select br, sum(sozd) sozd, sum(zap) zap from (select distinct woh.work_ord_nbr, case when woh.stat_code = '10' then 1 else 0 end sozd, case when woh.stat_code = '20' then 1 else 0 end zap,substr(im.sku_desc,1,3) br from work_ord_hdr woh join work_ord_dtl wod on wod.work_ord_nbr = woh.work_ord_nbr join item_master im on im.sku_id = wod.sku_id where woh.stat_code in ('20') )group by br order by 3 desc"
        orarec.Open(sqlz, oraconn)
        ind = 0
        d = 0
        e = 0
        Do Until orarec.EOF Or ind = 23
            a = orarec.Fields("zap").Value
            'b = orarec.Fields("zap").Value
            c = orarec.Fields("br").Value
            DataGridView11.Rows.Add(a, c)
            d = d + a
            'e = e + b
            ind = ind + 1
            orarec.MoveNext()
        Loop
        If d <> 0 Or e <> 0 Then DataGridView11.Rows.Add(d, "Всего") : DataGridView11.Rows(ind).DefaultCellStyle.BackColor = Color.Cyan
        Me.DataGridView11.Refresh()
        DataGridView11.AllowUserToAddRows = False
        orarec.Close()
        oraconn.close()
        '**********************************************************************


        'Зона отбора NK **************************************************
        DataGridView12.Rows.Clear()
        oraconn = CreateObject("ADODB.Connection")
        orarec = CreateObject("ADODB.Recordset")
        oraconn.Open("Provider=OraOLEDB.Oracle.1;User ID=gnatyk;Password=gnatyk;Data Source=WMOS_PROD;Persist Security Info=False")
        sqlz = "select area, count(tsk) zad, dep from (select distinct th.task_id tsk, lh.area area "

        sqlz = sqlz + ",case when im.cd_master_id = '19004' then 'DD' when im.cd_master_id = '2001' then case when im.size_desc = '1214119' and ph.shipto_name like '%Плато%' then 'PLATO' when im.size_desc = 'UPAK-DOO' then 'DOO' when ph.vendor_nbr like '%W60%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' when ph.vendor_nbr like '%W63%' and substr(ph.shipto_name, 1, 2) in ('П ', 'Н ') then 'DD_DROP' when ph.vendor_nbr like '%W60%' then 'DD' when ph.vendor_nbr like '%W63%' then 'DD' when ph.vendor_nbr like '%W640%' then 'DOO' when ph.vendor_nbr like '%W650%' then 'DOO' when ph.vendor_nbr like '%W680%' then 'DOO' when ph.vendor_nbr like '%MITIN01%' then 'INET' when ph.vendor_nbr like '%MPLIN01%' then 'INET' when ph.vendor_nbr like '%MECIN01%' then 'INET' Else Case when 'P' in (select 'P' from pkt_dtl ad1, item_master im1 where im1.sku_id = ad1.sku_id and ph.pkt_ctrl_nbr = ad1.pkt_ctrl_nbr and substr(im1.sale_grp, 0, 1) = 'P') then 'P/D' else case when im.sale_grp = '      ' or substr(im.sale_grp, 0, 1) = 'T' or (im.sale_grp = 'DRM' and substr(im.sku_desc, 0, 3) in ('CAM', 'CLR', 'GEX', 'LOB', 'VGB')) then 'DOO' "
        sqlz = sqlz + "when im.sale_grp = '   ' or (substr(im.sale_grp, 0, 1) = 'D' and  im.sale_grp <> 'DRM') or  (im.sale_grp = 'DRM' and substr(im.sku_desc, 0, 3) not in  ('CAM', 'CLR', 'GEX', 'LOB', 'VGB')) or im.sale_grp in ('P00', 'PPK', '_PROT', 'PRO') then 'DD' when im.sale_grp in ('_DL', '_TSOI', '_BASHL') then 'DL' Else 'OTHER' End End End else a.addr_line_1 end dep "

        sqlz = sqlz + "from task_hdr th join task_dtl td on td.task_id = th.task_id join locn_hdr lh on lh.locn_id = td.pull_locn_id join item_master im on im.sku_id = td.sku_id join pkt_hdr ph on ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr join pkt_hdr_intrnl phi on phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr left join cd_master cm on cm.cd_master_id = im.cd_master_id and cm.cd_master_id not in ('1','7') left join address a on substr(a.addr_key_1,7,5)=cm.div and a.addr_type = '01' where substr(lh.area,1,2)='NK' and th.stat_code < '90') group by area, dep order by 1,3"
        orarec.Open(sqlz, oraconn)
        Do Until orarec.EOF
            a = orarec.Fields("area").Value
            b = orarec.Fields("zad").Value
            c = orarec.Fields("dep").Value
            DataGridView12.Rows.Add(a, b, c)
            orarec.MoveNext()
        Loop
        Me.DataGridView12.Refresh()
        DataGridView12.AllowUserToAddRows = False
        orarec.Close()
        oraconn.close()
        '**********************************************************************

    End Sub
End Class
