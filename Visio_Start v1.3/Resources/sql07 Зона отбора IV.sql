SELECT area,
       COUNT(tsk) zad,
       dep
  FROM (SELECT DISTINCT th.task_id tsk,
                        lh.area area,
               CASE
                WHEN ph.cd_master_id = '2001' THEN
                 CASE
                   WHEN phi.rtl_pkt_flag = '1' THEN 'DOO SEZ'
                   WHEN ph.ord_type IN ('I', 'A', 'T', 'UA', 'UI', 'JA', 'JI', 'O') THEN 'DOO INT'
                    ELSE a.name
                 END
               ELSE a.name
               END                         AS dep
          FROM task_hdr th
          JOIN task_dtl td ON td.task_id = th.task_id
          JOIN locn_hdr lh ON lh.locn_id = td.pull_locn_id
          JOIN item_master im ON im.sku_id = td.sku_id AND im.cd_master_id NOT IN ('9005', '9006', '11005', '18004', '10003')
          JOIN pkt_hdr ph ON ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr
          JOIN pkt_hdr_intrnl phi ON phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr
          JOIN wcd_master w ON w.cd_master_id = im.cd_master_id
          JOIN address a ON a.addr_id = w.pkt_addr_id            
         WHERE substr(th.begin_area, 1, 2) = 'IV' AND th.begin_area <> 'IVS' AND th.stat_code < '90')
 GROUP BY area,
          dep
 ORDER BY 1,
          3