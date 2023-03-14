SELECT SUM(bl) bl,
       SUM(zap) zap,
       dep
  FROM (SELECT DISTINCT th.task_id,
                        a.name AS dep,
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
          JOIN wcd_master w ON w.cd_master_id = im.cd_master_id
          JOIN address a ON a.addr_id = w.pkt_addr_id             
         WHERE th.invn_need_type = 1 AND th.end_curr_work_area <> 'APL' AND th.end_curr_work_grp <> 'APL' AND th.stat_code < '90')
 GROUP BY dep
 ORDER BY 3