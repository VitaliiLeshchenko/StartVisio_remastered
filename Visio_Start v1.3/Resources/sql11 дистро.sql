SELECT br,
       COUNT(sd) sku,
       SUM(in_di) in_di
  FROM (SELECT im.size_desc sd,
               im.sku_desc sd1,
               substr(im.sku_desc, 1, 3) br,
               substr(im.sale_grp, 1, 3) sg,
               sd.distro_nbr din,
               sm.store_nbr km,
               sm.name nm,
               a.city ci,
               round(sd.reqd_qty) in_di,
               SUM(cd.actl_qty) na_skl,
               MAX(ch.create_date_time) posl_pr
          FROM store_distro sd
          JOIN case_dtl cd ON cd.sku_id = sd.sku_id AND cd.prod_stat = '00' AND nvl(cd.batch_nbr, '*') = sd.batch_nbr
          JOIN case_hdr ch ON ch.case_nbr = cd.case_nbr AND ch.stat_code IN ('10', '30', '45')
          JOIN item_master im ON im.sku_id = sd.sku_id
          JOIN store_master sm ON sm.store_nbr = sd.store_nbr
          JOIN address a ON a.addr_id = sm.addr_id
         WHERE sd.stat_code < '90' AND sd.reqd_qty > 0 AND
               (ch.rcvd_shpmt_nbr IS NULL OR ch.rcvd_shpmt_nbr IN (SELECT ah.shpmt_nbr
                                                                     FROM wmos.asn_hdr ah
                                                                    WHERE ah.stat_code = '90'
                                                                   UNION ALL
                                                                   SELECT ah.shpmt_nbr
                                                                     FROM archwmos.asn_hdr ah
                                                                    WHERE ah.stat_code = '90'))
         GROUP BY im.size_desc,
                  im.sku_desc,
                  im.sale_grp,
                  sd.distro_nbr,
                  sd.reqd_qty,
                  sm.store_nbr,
                  sm.name,
                  a.city
        HAVING SUM(cd.actl_qty) > 0
        UNION ALL
        SELECT im.size_desc sd,
               im.sku_desc sd1,
               substr(im.sku_desc, 1, 3) br,
               substr(im.sale_grp, 1, 3) sg,
               sd.distro_nbr din,
               sm.store_nbr km,
               sm.name nm,
               a.city ci,
               round(sd.reqd_qty) in_di,
               SUM(pld.actl_invn_qty) na_skl,
               MAX(ptt.create_date_time) posl_pr
          FROM store_distro sd
          JOIN pick_locn_dtl pld ON pld.sku_id = sd.sku_id AND nvl(pld.batch_nbr, '*') = sd.batch_nbr
          JOIN locn_hdr lh ON lh.locn_id = pld.locn_id AND lh.locn_class = 'A'
          JOIN item_master im ON im.sku_id = sd.sku_id
          JOIN store_master sm ON sm.store_nbr = sd.store_nbr
          JOIN address a ON a.addr_id = sm.addr_id
          JOIN prod_trkg_tran ptt ON ptt.sku_id = im.sku_id AND ptt.tran_type = '100'
          JOIN asn_hdr ah ON ah.shpmt_nbr = ptt.ref_field_1 AND ah.stat_code = '90'
         WHERE sd.stat_code < '90' AND sd.reqd_qty > 0
         GROUP BY im.size_desc,
                  im.sku_desc,
                  im.sale_grp,
                  sd.distro_nbr,
                  sd.reqd_qty,
                  sm.store_nbr,
                  sm.name,
                  a.city
        HAVING SUM(pld.actl_invn_qty) > 0)
 GROUP BY br
 ORDER BY 1
