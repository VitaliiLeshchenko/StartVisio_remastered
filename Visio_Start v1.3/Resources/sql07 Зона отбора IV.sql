SELECT area,
       COUNT(tsk) zad,
       dep
  FROM (SELECT DISTINCT th.task_id tsk,
                        lh.area area,
                        CASE
                          WHEN im.cd_master_id = '13004' THEN
                           'Коттон'
                          WHEN im.cd_master_id = '3001' AND im.sale_grp = 'TCK' THEN
                           'OXYGEN'
                          WHEN im.cd_master_id = '3001' AND im.sale_grp IN ('TCS', 'TCSR') THEN
                           'USB'
                          WHEN im.cd_master_id = '17004' THEN
                           'БиоЛайн'
                          WHEN im.cd_master_id = '18004' THEN
                           'Авто з/ч'
                          WHEN im.cd_master_id = '3001' THEN
                           'Л-Трейд'
                          WHEN im.cd_master_id = '19004' THEN
                           'DD'
                          WHEN im.cd_master_id = '3002' THEN
                           'MTI/00002'
                          WHEN im.cd_master_id = '3003' THEN
                           'Альпина'
                          WHEN im.cd_master_id = '4003' THEN
                           'БНС Трейд'
                          WHEN im.cd_master_id = '4004' THEN
                           'БНС Кампани'
                          WHEN im.cd_master_id = '5003' THEN
                           'Мазеркер'
                          WHEN im.cd_master_id = '5004' THEN
                           'Форс'
                          WHEN im.cd_master_id = '6003' THEN
                           CASE
                             WHEN im.sale_grp = 'P01' THEN
                              'LOR'
                             WHEN im.sale_grp IS NULL AND substr(im.sku_desc, 1, 3) = 'LOR' THEN
                              'LOR'
                             WHEN ph.shipto_addr_1 NOT LIKE '%нтернет%' AND ph.shipto_addr_1 NOT LIKE '%ахтерск%' AND
                                  ph.shipto_addr_1 NOT LIKE '%АХТЕРСК%' AND phi.total_nbr_of_units > 4 THEN
                              'PROT'
                             ELSE
                              'PROT_IM'
                           END
                          WHEN im.cd_master_id = '7003' THEN
                           'PWA'
                          WHEN im.cd_master_id = '8003' THEN
                           'БПИ'
                          WHEN im.cd_master_id = '9003' THEN
                           'Форс'
                          WHEN im.cd_master_id = '9004' THEN
                           'Силд Эйр'
                          WHEN im.cd_master_id = '9005' THEN
                           'Тека'
                          WHEN im.cd_master_id = '9006' THEN
                           'Легранд'
                          WHEN im.cd_master_id = '10003' THEN
                           'Сабриз'
                          WHEN im.cd_master_id = '10004' THEN
                           'Грейс Про'
                          WHEN im.cd_master_id = '11004' THEN
                           'Карма'
                          WHEN im.cd_master_id = '11005' THEN
                           'Калугин'
                          WHEN im.cd_master_id = '2001' THEN
                           CASE
                             WHEN im.size_desc = '1214119' AND ph.shipto_name LIKE '%Плато%' THEN
                              'PLATO'
                             WHEN im.size_desc = 'UPAK-DOO' THEN
                              'DOO'
                             WHEN ph.vendor_nbr LIKE '%W60%' AND substr(ph.shipto_name, 1, 2) IN ('П ', 'Н ') THEN
                              'DD_DROP'
                             WHEN ph.vendor_nbr LIKE '%W63%' AND substr(ph.shipto_name, 1, 2) IN ('П ', 'Н ') THEN
                              'DD_DROP'
                             WHEN ph.vendor_nbr LIKE '%W60%' THEN
                              'DD'
                             WHEN ph.vendor_nbr LIKE '%W63%' THEN
                              'DD'
                             WHEN ph.vendor_nbr LIKE '%W640%' THEN
                              'DOO'
                             WHEN ph.vendor_nbr LIKE '%W650%' THEN
                              'DOO'
                             WHEN ph.vendor_nbr LIKE '%W680%' THEN
                              'DOO'
                             WHEN ph.vendor_nbr LIKE '%MITIN01%' AND ph.ord_type <> ' ' THEN
                              'INET'
                             WHEN ph.vendor_nbr LIKE '%MPLIN01%' AND ph.ord_type <> ' ' THEN
                              'INET'
                             WHEN ph.vendor_nbr LIKE '%MECIN01%' AND ph.ord_type <> ' ' THEN
                              'INET'
                             WHEN substr(ph.shipto_name, 1, 6) = 'IN IT ' OR ph.pkt_type = 'D' OR ph.ord_type IN ('O', 'I') THEN
                              'INET'
                             ELSE
                              CASE
                                WHEN 'P' IN
                                     (SELECT 'P'
                                        FROM pkt_dtl     ad1,
                                             item_master im1
                                       WHERE im1.sku_id = ad1.sku_id AND ph.pkt_ctrl_nbr = ad1.pkt_ctrl_nbr AND substr(im1.sale_grp, 0, 1) = 'P') THEN
                                 'P/D'
                                ELSE
                                 CASE
                                   WHEN im.sale_grp = '      ' OR substr(im.sale_grp, 0, 1) = 'T' OR
                                        (im.sale_grp = 'DRM' AND substr(im.sku_desc, 0, 3) IN ('CAM', 'CLR', 'GEX', 'LOB', 'VGB')) THEN
                                    'DOO'
                                   WHEN im.sale_grp = '   ' OR (substr(im.sale_grp, 0, 1) = 'D' AND im.sale_grp <> 'DRM') OR
                                        (im.sale_grp = 'DRM' AND substr(im.sku_desc, 0, 3) NOT IN ('CAM', 'CLR', 'GEX', 'LOB', 'VGB')) OR
                                        im.sale_grp IN ('P00', 'PPK', '_PROT', 'PRO') THEN
                                    'DD'
                                   WHEN im.sale_grp IN ('_DL', '_TSOI', '_BASHL') THEN
                                    'DL'
                                   ELSE
                                    'DOO'
                                 END
                              END
                           END
                        END dep
          FROM task_hdr th
          JOIN task_dtl td ON td.task_id = th.task_id
          JOIN locn_hdr lh ON lh.locn_id = td.pull_locn_id
          JOIN item_master im ON im.sku_id = td.sku_id AND im.cd_master_id NOT IN ('9005', '9006', '11005', '18004', '10003')
          JOIN pkt_hdr ph ON ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr
          JOIN pkt_hdr_intrnl phi ON phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr
         WHERE substr(th.begin_area, 1, 2) = 'IV' AND th.begin_area <> 'IVS' AND th.stat_code < '90')
 GROUP BY area,
          dep
 ORDER BY 1,
          3
