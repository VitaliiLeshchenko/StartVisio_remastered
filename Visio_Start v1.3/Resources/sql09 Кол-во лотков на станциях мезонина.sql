SELECT st,
       MAX(lototb) lototb,
       MAX(lot_popoln) lot_popoln
  FROM (SELECT st,
               COUNT(*) lototb,
               0 lot_popoln
          FROM (SELECT o.task_id,
                       c.mod_date_time prisvoen,
                       MIN(l.pick_detrm_zone) st
                  FROM task_hdr        o,
                       task_dtl        d,
                       locn_hdr        l,
                       c_umti_mhe_cntr c
                 WHERE (o.task_id = d.task_id AND d.pull_locn_id = l.locn_id AND o.task_id = c.task_id AND c.cntr_nbr IS NOT NULL) AND
                       o.task_type = '16' AND d.stat_code < '90'
                 GROUP BY o.task_id,
                          c.mod_date_time) a
         GROUP BY st
        UNION ALL
        SELECT l.pick_detrm_zone st,
               0 lototb,
               COUNT(*) lot_popoln
          FROM case_hdr ch,
               locn_hdr l
         WHERE ch.dest_locn_id = l.locn_id AND ch.stat_code < 65 AND ch.plt_id IS NOT NULL AND l.locn_class = 'A' AND l.pick_detrm_zone <> 'SCS'
         GROUP BY l.pick_detrm_zone)
 GROUP BY st
 ORDER BY 1
