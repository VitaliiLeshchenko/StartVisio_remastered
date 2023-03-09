SELECT CASE
         WHEN pl LIKE 'A%L' THEN
          'A_L'
         WHEN pl LIKE 'A%R' THEN
          'A_R'
         ELSE
          pl
       END pl,
       COUNT(1) vsg,
       SUM(blok) blk,
       SUM(wait) vip,
       SUM(process) wrk,
       MAX(prostoy) paus,
       SUM(st40) st40
  FROM (SELECT th.task_id,
               th.start_curr_work_area pl,
               th.curr_task_prty pr,
               CASE
                 WHEN th.stat_code = '5' THEN
                  1
                 ELSE
                  0
               END blok,
               CASE
                 WHEN th.stat_code > '5' AND th.stat_code < 20 THEN
                  1
                 ELSE
                  0
               END wait,
               CASE
                 WHEN th.stat_code >= '20' THEN
                  1
                 ELSE
                  0
               END process,
               CASE
                 WHEN th.curr_task_prty = '40' THEN
                  1
                 ELSE
                  0
               END st40,
               round((SYSDATE - th.mod_date_time) * 24) prostoy
          FROM task_hdr th
          JOIN task_dtl td ON td.task_id = th.task_id AND td.cd_master_id NOT IN ('9005', '9006', '11005', '18004', '10003')
         WHERE (th.stat_code < 90) AND th.invn_need_type <> 100 AND th.task_type <> 09 AND
               (th.start_curr_work_area LIKE 'PL%' OR (th.start_curr_work_area = 'XXL') OR
               (th.start_curr_work_area LIKE 'A%' AND th.end_dest_work_area LIKE 'A%')))
 GROUP BY CASE
            WHEN pl LIKE 'A%L' THEN
             'A_L'
            WHEN pl LIKE 'A%R' THEN
             'A_R'
            ELSE
             pl
          END
 ORDER BY 1 DESC
