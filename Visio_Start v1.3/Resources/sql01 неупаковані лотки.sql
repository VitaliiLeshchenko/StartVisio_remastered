SELECT day2,
       dep,
       SUM(not_lotok) not_lotok,
       SUM(lotok) lotok,
       SUM(otobr_upak) otobr_upak,
       SUM(otobr_not_upak) otobr_not_upak,
       SUM(not_lotok) + SUM(lotok) + SUM(otobr_upak) + SUM(otobr_not_upak) vsego,
       prosr
  FROM (SELECT day1,
               day2,
               dep,
               lot,
               COUNT(locn) strok,
               SUM(qty) qty,
               SUM(ltr) ltr,
               not_lotok,
               lotok,
               otobr_upak,
               otobr_not_upak,
               prosr
          FROM (SELECT DISTINCT td.task_id lot,
                                td.task_seq_nbr,
                                to_char(th.mod_date_time, 'YYYY-MM-DD') day1,
                                to_char(th.mod_date_time, 'DD.MM.YYYY') day2,
                                CASE
                                  WHEN to_char(th.mod_date_time, 'DD.MM.YYYY') <> to_char(SYSDATE, 'DD.MM.YYYY') THEN
                                   1
                                  ELSE
                                   0
                                END prosr,
                                CASE
                                   WHEN ph.cd_master_id = '2001' THEN
                                    CASE
                                      WHEN phi.rtl_pkt_flag = '1' THEN 'DOO SEZ'
                                      WHEN ph.ord_type IN ('I', 'A', 'T', 'UA', 'UI', 'JA', 'JI', 'O') THEN 'DOO INT'
                                       ELSE a.name
                                    END
                                   ELSE a.name
                                 END                         AS dep,
                                im.size_desc,
                                lf.dsp_locn locn,
                                round(td.qty_pulld) qty,
                                round(im.unit_vol * td.qty_pulld / 1000, 2) ltr,
                                CASE
                                  WHEN td.stat_code < '90' AND cu.cntr_nbr IS NULL THEN 1
                                  ELSE 0
                                END not_lotok,
                                CASE
                                  WHEN td.stat_code < '90' AND cu.cntr_nbr IS NOT NULL THEN 1
                                  ELSE 0
                                END lotok,
                                CASE
                                  WHEN td.stat_code = '90' AND ch.stat_code >= '20' THEN 1
                                  ELSE 0
                                END otobr_upak,
                                CASE
                                  WHEN td.stat_code = '90' AND ch.stat_code < '20' THEN 1
                                  ELSE 0
                                END otobr_not_upak
                  FROM task_dtl td
                  JOIN item_master im ON im.sku_id = td.sku_id AND im.cd_master_id NOT IN ('9005', '9006', '11005', '18004')
                  JOIN locn_hdr lf ON lf.locn_id = td.pull_locn_id
                  JOIN task_hdr th ON th.task_id = td.task_id
                  JOIN pkt_hdr ph ON ph.pkt_ctrl_nbr = td.pkt_ctrl_nbr
                  JOIN pkt_hdr_intrnl phi ON phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr
                  JOIN c_umti_mhe_cntr cu ON cu.task_id = th.task_id
                  JOIN carton_dtl cd ON cd.carton_seq_nbr = td.carton_seq_nbr
                  JOIN carton_hdr ch ON ch.carton_nbr = cd.carton_nbr AND ch.pkt_ctrl_nbr = phi.pkt_ctrl_nbr
                  JOIN (SELECT DISTINCT aid50.alloc_invn_dtl_id,
                                       aid52.carton_seq_nbr,
                                       aid52.cntr_nbr AS tote_nbr,
                                       aid52.carton_nbr
                         FROM alloc_invn_dtl aid50
                         JOIN alloc_invn_dtl aid52 ON aid50.carton_nbr = aid52.carton_nbr) aid ON aid.carton_seq_nbr = cd.carton_seq_nbr AND
                                                                                                  aid.carton_nbr = cd.carton_nbr AND
                                                                                                  aid.alloc_invn_dtl_id = td.alloc_invn_dtl_id
                  JOIN wcd_master w ON w.cd_master_id = im.cd_master_id
                  JOIN address a ON a.addr_id = w.pkt_addr_id
                 WHERE th.task_desc IN ('Отбор с мезонина', 'Упаковка на мезонине', 'Pick pack MEZ') AND
                       substr(lf.dsp_locn, 1, 1) IN ('K', 'L', 'M', 'N') AND th.stat_code <> '99' AND
                       (to_char(th.mod_date_time, 'DD-MM-YYYY') = to_char(SYSDATE, 'DD-MM-YYYY') OR th.stat_code = '10'))
         GROUP BY day1,
                  day2,
                  dep,
                  lot,
                  not_lotok,
                  lotok,
                  otobr_upak,
                  otobr_not_upak,
                  prosr)
 GROUP BY ROLLUP(dep),
          day2,
          prosr
 ORDER BY 2, 1
