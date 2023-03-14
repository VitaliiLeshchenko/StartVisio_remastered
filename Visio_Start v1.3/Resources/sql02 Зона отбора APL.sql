SELECT SUM(bl) bl,
       SUM(zap) zap,
       dep
  FROM (SELECT DISTINCT th.task_id,
                                CASE
                                   WHEN ph.cd_master_id = '2001' THEN
                                    CASE
                                      WHEN phi.rtl_pkt_flag = '1' THEN 'DOO SEZ'
                                      WHEN ph.ord_type IN ('I', 'A', 'T', 'UA', 'UI', 'JA', 'JI', 'O') THEN 'DOO INT'
                                       ELSE a.name
                                    END
                                   ELSE a.name
                                 END                         AS dep,
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
         WHERE th.task_desc IN ('Îעבמנ ס À-ןאככוע') AND substr(lh.area, 1, 1) = 'A' AND th.stat_code < '90')
 GROUP BY dep
 ORDER BY 3
