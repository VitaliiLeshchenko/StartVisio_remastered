SELECT SUM(bl) bl,
       SUM(zap) zap,
       dep
  FROM (SELECT DISTINCT th.task_id,
                        CASE
                          WHEN im.cd_master_id = '18004' THEN
                           'Авто з/ч'
                          WHEN im.cd_master_id = '3001' AND im.sale_grp = 'TCK' THEN
                           'OXYGEN'
                          WHEN im.cd_master_id = '3001' AND im.sale_grp IN ('TCS', 'TCSR') THEN
                           'USB'
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
                             WHEN im.size_desc = 'UPAK-DOO' THEN
                              'DOO'
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
                        END dep,
                        CASE
                          WHEN th.start_curr_work_area <> 'MIS' THEN
                           1
                          ELSE
                           0
                        END bl,
                        CASE
                          WHEN th.start_curr_work_area = 'MIS' THEN
                           1
                          ELSE
                           0
                        END zap
          FROM task_hdr th
          JOIN task_dtl td ON td.task_id = th.task_id
          JOIN locn_hdr lh ON lh.locn_id = td.pull_locn_id
          JOIN item_master im ON im.sku_id = td.sku_id AND im.cd_master_id NOT IN ('9005', '9006', '11005', '18004', '10003')
          JOIN case_hdr ch ON ch.case_nbr = td.cntr_nbr AND ch.plt_id IS NULL
         WHERE th.invn_need_type = 1 AND th.end_curr_work_area <> 'APL' AND th.end_curr_work_grp <> 'APL' AND th.stat_code < '90')
 GROUP BY dep
 ORDER BY 3
