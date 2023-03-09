SELECT SUM(nev) nev,
       SUM(zap) zap
  FROM (SELECT DISTINCT ph.pkt_ctrl_nbr,
                        CASE
                          WHEN phi.stat_code < '20' THEN
                           1
                          ELSE
                           0
                        END nev,
                        CASE
                          WHEN phi.stat_code >= '20' THEN
                           1
                          ELSE
                           0
                        END zap
          FROM pkt_hdr ph
          JOIN pkt_hdr_intrnl phi ON phi.pkt_ctrl_nbr = ph.pkt_ctrl_nbr AND phi.stat_code < '90'
         WHERE ph.shipto_name IN ('Декомплектация ростовок', 'Д KYI Internet-Інтертоп', 'Декомплектация') OR ph.shipto_name LIKE 'Декомпл%')
