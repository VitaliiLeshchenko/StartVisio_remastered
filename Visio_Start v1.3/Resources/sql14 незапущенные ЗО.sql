SELECT dep,
       COUNT(zo) zo
  FROM (SELECT DISTINCT                                 CASE
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
                        ph.pkt_ctrl_nbr zo
          FROM pkt_hdr_intrnl phi
          JOIN pkt_hdr ph ON ph.pkt_ctrl_nbr = phi.pkt_ctrl_nbr
          JOIN pkt_dtl pd ON pd.pkt_ctrl_nbr = ph.pkt_ctrl_nbr
          JOIN item_master im ON im.sku_id = pd.sku_id AND
                                 im.cd_master_id NOT IN ('3003', '9005', '9006', '10003', '11005', '14004', '18004', '20004')
          JOIN wcd_master w ON w.cd_master_id = im.cd_master_id
          JOIN address a ON a.addr_id = w.pkt_addr_id
         WHERE phi.stat_code = '10' AND
               (im.cd_master_id <> '2001' OR (substr(ph.shipto_name, 1, 6) = 'IN IT ' OR ph.pkt_type = 'D' OR ph.ord_type IN ('O', 'I'))) AND
               ph.pkt_ctrl_nbr NOT LIKE '%test%' AND ph.pkt_ctrl_nbr NOT LIKE '%TEST%' AND ph.pkt_ctrl_nbr NOT LIKE '%opol%')
 GROUP BY dep
