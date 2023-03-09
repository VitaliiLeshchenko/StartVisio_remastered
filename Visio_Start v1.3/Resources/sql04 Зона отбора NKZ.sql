SELECT COUNT(*) a
  FROM (SELECT DISTINCT th.task_id
          FROM task_dtl td
          JOIN task_hdr th ON th.task_id = td.task_id AND th.stat_code < '90'
          JOIN locn_hdr lh ON lh.locn_id = td.pull_locn_id AND lh.area = 'NKZ'
         WHERE td.cd_master_id NOT IN ('9005', '9006', '11005', '18004'))
