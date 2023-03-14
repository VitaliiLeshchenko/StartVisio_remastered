SELECT pkt.departament                         AS departament,
       COUNT(DISTINCT pkt.pkt_ctrl_nbr)        AS zo_qty,
       COUNT(*)                                AS row_qty,
       ROUND(SUM(pkt.row_volume) / 1000000, 2) AS volume_m3
  FROM (
SELECT CASE
         WHEN ph.cd_master_id = '2001' THEN
          CASE
            WHEN pi.rtl_pkt_flag = '1' THEN 'DOO SEZ'
            WHEN ph.ord_type IN ('I', 'A', 'T', 'UA', 'UI', 'JA', 'JI', 'O') THEN 'DOO INT'
             ELSE a.name
          END
         ELSE a.name
       END                         AS departament,
       ph.pkt_ctrl_nbr             AS pkt_ctrl_nbr,
       pd.units_pakd * im.unit_vol AS row_volume
  FROM pkt_hdr_intrnl pi
  JOIN pkt_hdr ph ON ph.pkt_ctrl_nbr = pi.pkt_ctrl_nbr
  JOIN pkt_dtl pd ON pd.pkt_ctrl_nbr = ph.pkt_ctrl_nbr
  JOIN item_master im ON im.sku_id = pd.sku_id
  JOIN wcd_master w ON w.cd_master_id = ph.cd_master_id
  JOIN address a ON a.addr_id = w.pkt_addr_id
 WHERE pi.stat_code = 90 AND to_char(pi.mod_date_time, 'YYYY-MM-DD') = to_char(SYSDATE, 'YYYY-MM-DD')
 ) pkt
 GROUP BY pkt.departament