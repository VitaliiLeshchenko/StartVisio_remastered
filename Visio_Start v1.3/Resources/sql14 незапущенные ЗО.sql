  SELECT dep,
   COUNT (zo) zo
    FROM (SELECT DISTINCT
                 CASE
                   WHEN im.cd_master_id = '2001' THEN
                     CASE
                       WHEN pi.rtl_pkt_flag = 1 THEN 'SEZ'
                       WHEN ph.ord_type IN ('I', 'A', 'T', 'UA', 'UI', 'JA', 'JI', 'O') THEN 'INT' END
                     ELSE a.name END dep,
                 ph.pkt_ctrl_nbr zo,
                 ph.vendor_nbr,
                 ph.shipto_name
            FROM pkt_hdr_intrnl pi
            JOIN pkt_hdr ph ON ph.pkt_ctrl_nbr = pi.pkt_ctrl_nbr
            JOIN pkt_dtl pd ON pd.pkt_ctrl_nbr = ph.pkt_ctrl_nbr
            JOIN item_master im ON im.sku_id = pd.sku_id
            JOIN wcd_master w ON w.cd_master_id = im.cd_master_id
            JOIN address a ON a.addr_id = w.pkt_addr_id
           WHERE pi.stat_code = '10' AND
                 (im.cd_master_id <> '2001' OR (substr(ph.shipto_name, 1, 6) = 'IN IT ' OR ph.pkt_type = 'D' OR ph.ord_type IN ('O', 'I'))) AND
                 ph.pkt_ctrl_nbr NOT LIKE '%test%' AND
                 ph.pkt_ctrl_nbr NOT LIKE '%TEST%' AND
                 ph.pkt_ctrl_nbr NOT LIKE '%opol%' AND
                 a.zip NOT IN ('×àéêà2', 'Ëüâ³â')
        )
GROUP BY dep