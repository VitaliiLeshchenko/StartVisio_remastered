SELECT SUM(bl) bl,
       SUM(zap) zap,
       dep
  FROM (SELECT DISTINCT th.task_id,
                        CASE
                          WHEN im.cd_master_id = '3001' AND im.sale_grp = 'TCK' THEN 'OXYGEN'
                          WHEN im.cd_master_id = '3001' AND im.sale_grp IN ('TCS', 'TCSR') THEN 'USB'
                          WHEN im.cd_master_id = '3001' THEN 'Л-Трейд'
                          WHEN im.cd_master_id = '6003' THEN
                           CASE
                             WHEN im.sale_grp = 'P01' THEN 'LOR'
                             WHEN im.sale_grp IS NULL AND substr(im.sku_desc, 1, 3) = 'LOR' THEN 'LOR'
                             WHEN ph.shipto_addr_1 NOT LIKE '%нтернет%' AND ph.shipto_addr_1 NOT LIKE '%ахтерск%' AND
                                  ph.shipto_addr_1 NOT LIKE '%АХТЕРСК%' AND phi.total_nbr_of_units > 4 THEN 'PROT'
                             ELSE 'PROT_IM'
                           END
                          WHEN im.cd_master_id = '2001' THEN
                           CASE
                             WHEN im.size_desc = '1214119' AND ph.shipto_name LIKE '%Плато%' THEN 'PLATO'
                             WHEN im.size_desc = 'UPAK-DOO' THEN a.name
                             WHEN ph.vendor_nbr LIKE '%W60%' AND substr(ph.shipto_name, 1, 2) IN ('П ', 'Н ') THEN 'DD_DROP'
                             WHEN ph.vendor_nbr LIKE '%W63%' AND substr(ph.shipto_name, 1, 2) IN ('П ', 'Н ') THEN 'DD_DROP'
                             WHEN substr(ph.shipto_name, 1, 6) = 'IN IT ' OR ph.pkt_type = 'D' OR ph.ord_type IN ('O', 'I') THEN 'INET'
                             ELSE a.name
                           END
                           ELSE a.name
                        END dep,
                        CASE
                          WHEN th.stat_code = '5' THEN 1
                          ELSE 0
                        END bl,
                        CASE
                          WHEN th.stat_code = '5' THEN 0
                          ELSE 1
                        END zap
          FROM task_hdr th
          JOIN task_dtl td ON td.task_id = th.task_id
          JOIN locn_hdr lh ON lh.locn_id = td.pull_locn_id
          JOIN item_master im ON im.sku_id = td.sku_id AND im.cd_master_id NOT IN ('9005', '9006', '11005', '18004')
          JOIN pkt_hdr ph ON ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr
          JOIN pkt_hdr_intrnl phi ON phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr
          JOIN wcd_master w ON w.cd_master_id = im.cd_master_id
          JOIN address a ON a.addr_id = w.pkt_addr_id
         WHERE th.task_desc IN ('Отбор с А-паллет') AND substr(lh.area, 1, 1) = 'A' AND th.stat_code < '90')
 GROUP BY dep
 ORDER BY 3
