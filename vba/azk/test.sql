
##################################################

-- SPECIFIC TABLE - CREATION 
CREATE TABLE    IF NOT EXISTS default.vt_borrowlist_catlog (
id                            INT(11)                       NOT NULL                                                    COMMENT '',
typ                           TINYINT                       NOT NULL                      DEFAULT -1                    COMMENT '1 - INSERT/2 - CLOSED/3 - UNCHANGED/4 - CHANGED',
ver                           INT                           NOT NULL                      DEFAULT 1                     COMMENT 'VERSION NUMBER',
PRIMARY KEY (id,typ),
KEY idx(id)
 ) ENGINE=INNODB DEFAULT CHARSET=utf8 
 PARTITION BY HASH(typ) PARTITIONS 4;

CREATE TABLE IF NOT EXISTS default.vt_s_hyr_borrowlist_h LIKE default.s_hyr_borrowlist_h;

##################################################

-- Clear temporary table
TRUNCATE TABLE default.vt_borrowlist_catlog;
TRUNCATE TABLE default.vt_s_hyr_borrowlist_h;

-- step 01. Pick out records marked 'insert' ( type = 1 )
INSERT INTO default.vt_borrowlist_catlog 
SELECT
	  a.id
	, 1 -- 'insert'
	, 1 -- 'version 1 for insert-record'
FROM      default.borrowlist a
LEFT JOIN default.s_hyr_borrowlist_h b
  ON a.id = b.id
AND b.start_dt < '2016-01-01'
AND b.end_dt  >= '2016-01-01'
WHERE b.id IS NULL;

-- step 02. Pick out records marked 'closed'   ( type = 2 )
--                                  'unchanged'( type = 3 )
--                                  'changed'  ( type = 4 )
INSERT INTO default.vt_borrowlist_catlog 
SELECT
	  b.id
	, CASE WHEN a.id IS NULL                                
	       THEN 2 -- 'closed'                                    
	       WHEN                                                  
	                ((a.lendid                       = b.lendid                       ) OR (a.lendid                         IS NULL AND b.lendid                         IS NULL ))
	            AND ((a.crmlc_uid                    = b.crmlc_uid                    ) OR (a.crmlc_uid                      IS NULL AND b.crmlc_uid                      IS NULL ))
	            AND ((a.uid                          = b.uid                          ) OR (a.uid                            IS NULL AND b.uid                            IS NULL ))
	            AND ((COALESCE( a.usrcustid , '')    = COALESCE( b.usrcustid , '')    ) OR (a.usrcustid                      IS NULL AND b.usrcustid                      IS NULL ))
	            AND ((a.fromtype                     = b.fromtype                     ) OR (a.fromtype                       IS NULL AND b.fromtype                       IS NULL ))
	            AND ((a.ledger                       = b.ledger                       ) OR (a.ledger                         IS NULL AND b.ledger                         IS NULL ))
	            AND ((a.realrepay                    = b.realrepay                    ) OR (a.realrepay                      IS NULL AND b.realrepay                      IS NULL ))
	            AND ((a.orddatebid                   = b.orddatebid                   ) OR (a.orddatebid                     IS NULL AND b.orddatebid                     IS NULL ))
	            AND ((COALESCE( a.ordidbid , '')     = COALESCE( b.ordidbid , '')     ) OR (a.ordidbid                       IS NULL AND b.ordidbid                       IS NULL ))
	            AND ((a.loans_orddate                = b.loans_orddate                ) OR (a.loans_orddate                  IS NULL AND b.loans_orddate                  IS NULL ))
	            AND ((COALESCE( a.loans_ordid , '')  = COALESCE( b.loans_ordid , '')  ) OR (a.loans_ordid                    IS NULL AND b.loans_ordid                    IS NULL ))
	            AND ((a.repay_orddate                = b.repay_orddate                ) OR (a.repay_orddate                  IS NULL AND b.repay_orddate                  IS NULL ))
	            AND ((COALESCE( a.repay_ordid , '')  = COALESCE( b.repay_ordid , '')  ) OR (a.repay_ordid                    IS NULL AND b.repay_ordid                    IS NULL ))
	            AND ((a.lendtime                     = b.lendtime                     ) OR (a.lendtime                       IS NULL AND b.lendtime                       IS NULL ))
	            AND ((a.autobidtime                  = b.autobidtime                  ) OR (a.autobidtime                    IS NULL AND b.autobidtime                    IS NULL ))
	            AND ((a.statusrtime                  = b.statusrtime                  ) OR (a.statusrtime                    IS NULL AND b.statusrtime                    IS NULL ))
	            AND ((a.is_autobid                   = b.is_autobid                   ) OR (a.is_autobid                     IS NULL AND b.is_autobid                     IS NULL ))
	            AND ((a.lendstatus                   = b.lendstatus                   ) OR (a.lendstatus                     IS NULL AND b.lendstatus                     IS NULL ))
	            AND ((a.status                       = b.status                       ) OR (a.status                         IS NULL AND b.status                         IS NULL ))
	            AND ((a.is_transfer                  = b.is_transfer                  ) OR (a.is_transfer                    IS NULL AND b.is_transfer                    IS NULL ))
	            AND ((a.is_newcont                   = b.is_newcont                   ) OR (a.is_newcont                     IS NULL AND b.is_newcont                     IS NULL ))
	            AND ((a.newcontnum                   = b.newcontnum                   ) OR (a.newcontnum                     IS NULL AND b.newcontnum                     IS NULL ))
	            AND ((COALESCE( a.newcontid , '')    = COALESCE( b.newcontid , '')    ) OR (a.newcontid                      IS NULL AND b.newcontid                      IS NULL ))
	            AND ((COALESCE( a.newcrmid , '')     = COALESCE( b.newcrmid , '')     ) OR (a.newcrmid                       IS NULL AND b.newcrmid                       IS NULL ))
	            AND ((a.transfertime                 = b.transfertime                 ) OR (a.transfertime                   IS NULL AND b.transfertime                   IS NULL ))
	            AND ((a.transferstatus               = b.transferstatus               ) OR (a.transferstatus                 IS NULL AND b.transferstatus                 IS NULL ))
	            AND ((COALESCE( a.newconturl , '')   = COALESCE( b.newconturl , '')   ) OR (a.newconturl                     IS NULL AND b.newconturl                     IS NULL ))                                              
	       THEN     3 -- 'unchanged'                             
	       ELSE     4 -- 'changed'                               
	   END                                                       
	, b.version
FROM      default.s_hyr_borrowlist_h b
LEFT JOIN default.borrowlist a
  ON a.id = b.id
WHERE b.start_dt   <   '2016-01-01' -- fetch his-data !! support reload operation
  AND b.end_dt     >=  '2016-01-01' ; -- fetch his-data !! support reload operation

-- step 03. reduce result(new target table)       
INSERT INTO default.vt_s_hyr_borrowlist_h 
(
	  id
	, lendid
	, crmlc_uid
	, uid
	, usrcustid
	, borrowid
	, crmdk_uid
	, dkuid
	, dkusrcustid
	, loanmatcterm
	, crmlcmatcterm
	, fromtype
	, source_addr
	, matchrate
	, matchmoney
	, ledger
	, servicemoney
	, repay
	, realrepay
	, trxid
	, repaytype
	, orddatebid
	, ordidbid
	, loans_orddate
	, loans_ordid
	, repay_orddate
	, repay_ordid
	, frezzeordid
	, lendtime
	, autobidtime
	, statusrtime
	, is_autobid
	, lendstatus
	, status
	, is_transfer
	, is_newcont
	, newcontnum
	, newcontid
	, transfermoney
	, transferamt
	, service
	, newcrmid
	, transfertime
	, transferstatus
	, newconturl
	, addtime
	, reclaim_amount
 , version       
 , start_dt      
 , end_dt        
 , load_stamp    
)
--  insert (from src) 
SELECT
	  a.id                           as id
	, a.lendid                       as lendid
	, a.crmlc_uid                    as crmlc_uid
	, a.uid                          as uid
	, COALESCE( a.usrcustid , '')    as usrcustid
	, a.borrowid                     as borrowid
	, a.crmdk_uid                    as crmdk_uid
	, a.dkuid                        as dkuid
	, COALESCE( a.dkusrcustid , '')  as dkusrcustid
	, a.loanmatcterm                 as loanmatcterm
	, a.crmlcmatcterm                as crmlcmatcterm
	, a.fromtype                     as fromtype
	, a.source_addr                  as source_addr
	, a.matchrate                    as matchrate
	, a.matchmoney                   as matchmoney
	, a.ledger                       as ledger
	, a.servicemoney                 as servicemoney
	, a.repay                        as repay
	, a.realrepay                    as realrepay
	, COALESCE( a.trxid , '')        as trxid
	, a.repaytype                    as repaytype
	, a.orddatebid                   as orddatebid
	, COALESCE( a.ordidbid , '')     as ordidbid
	, a.loans_orddate                as loans_orddate
	, COALESCE( a.loans_ordid , '')  as loans_ordid
	, a.repay_orddate                as repay_orddate
	, COALESCE( a.repay_ordid , '')  as repay_ordid
	, COALESCE( a.frezzeordid , '')  as frezzeordid
	, a.lendtime                     as lendtime
	, a.autobidtime                  as autobidtime
	, a.statusrtime                  as statusrtime
	, a.is_autobid                   as is_autobid
	, a.lendstatus                   as lendstatus
	, a.status                       as status
	, a.is_transfer                  as is_transfer
	, a.is_newcont                   as is_newcont
	, a.newcontnum                   as newcontnum
	, COALESCE( a.newcontid , '')    as newcontid
	, a.transfermoney                as transfermoney
	, a.transferamt                  as transferamt
	, a.service                      as service
	, COALESCE( a.newcrmid , '')     as newcrmid
	, a.transfertime                 as transfertime
	, a.transferstatus               as transferstatus
	, COALESCE( a.newconturl , '')   as newconturl
	, UNIX_TIMESTAMP(a.addtime)      as addtime
	, COALESCE( a.reclaim_amount , '')         as reclaim_amount
  , 1                           AS version      
  , '2016-01-01'                    AS start_dt     
  , '3000-12-31'                   AS end_dt       
  , CURRENT_TIMESTAMP                           
FROM       default.borrowlist                        a
INNER JOIN default.vt_borrowlist_catlog                            b
  ON a.id = b.id
  AND b.`typ` = 1 ; -- 'insert'

-- # closed (tar) & changed (tar) & unchanged (src)
INSERT INTO default.vt_s_hyr_borrowlist_h
(
	  id
	, lendid
	, crmlc_uid
	, uid
	, usrcustid
	, borrowid
	, crmdk_uid
	, dkuid
	, dkusrcustid
	, loanmatcterm
	, crmlcmatcterm
	, fromtype
	, source_addr
	, matchrate
	, matchmoney
	, ledger
	, servicemoney
	, repay
	, realrepay
	, trxid
	, repaytype
	, orddatebid
	, ordidbid
	, loans_orddate
	, loans_ordid
	, repay_orddate
	, repay_ordid
	, frezzeordid
	, lendtime
	, autobidtime
	, statusrtime
	, is_autobid
	, lendstatus
	, status
	, is_transfer
	, is_newcont
	, newcontnum
	, newcontid
	, transfermoney
	, transferamt
	, service
	, newcrmid
	, transfertime
	, transferstatus
	, newconturl
	, addtime
	, reclaim_amount
   , version    
   , start_dt   
   , end_dt     
   , load_stamp 
)
SELECT
	  a.id                             AS id
	, CASE WHEN b.id IS NULL                                       THEN a.lendid
	       WHEN b.typ  = 3                                         THEN c.lendid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.lendid
	  END  AS lendid
	, CASE WHEN b.id IS NULL                                       THEN a.crmlc_uid
	       WHEN b.typ  = 3                                         THEN c.crmlc_uid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.crmlc_uid
	  END  AS crmlc_uid
	, CASE WHEN b.id IS NULL                                       THEN a.uid
	       WHEN b.typ  = 3                                         THEN c.uid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.uid
	  END  AS uid
	, CASE WHEN b.id IS NULL                                       THEN a.usrcustid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.usrcustid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.usrcustid
	  END  AS usrcustid
	, CASE WHEN b.id IS NULL                                       THEN a.borrowid
	       WHEN b.typ  = 3                                         THEN c.borrowid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.borrowid
	  END  AS borrowid
	, CASE WHEN b.id IS NULL                                       THEN a.crmdk_uid
	       WHEN b.typ  = 3                                         THEN c.crmdk_uid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.crmdk_uid
	  END  AS crmdk_uid
	, CASE WHEN b.id IS NULL                                       THEN a.dkuid
	       WHEN b.typ  = 3                                         THEN c.dkuid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.dkuid
	  END  AS dkuid
	, CASE WHEN b.id IS NULL                                       THEN a.dkusrcustid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.dkusrcustid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.dkusrcustid
	  END  AS dkusrcustid
	, CASE WHEN b.id IS NULL                                       THEN a.loanmatcterm
	       WHEN b.typ  = 3                                         THEN c.loanmatcterm
	       WHEN b.typ IN ( 2               , 4                )    THEN a.loanmatcterm
	  END  AS loanmatcterm
	, CASE WHEN b.id IS NULL                                       THEN a.crmlcmatcterm
	       WHEN b.typ  = 3                                         THEN c.crmlcmatcterm
	       WHEN b.typ IN ( 2               , 4                )    THEN a.crmlcmatcterm
	  END  AS crmlcmatcterm
	, CASE WHEN b.id IS NULL                                       THEN a.fromtype
	       WHEN b.typ  = 3                                         THEN c.fromtype
	       WHEN b.typ IN ( 2               , 4                )    THEN a.fromtype
	  END  AS fromtype
	, CASE WHEN b.id IS NULL                                       THEN a.source_addr
	       WHEN b.typ  = 3                                         THEN c.source_addr
	       WHEN b.typ IN ( 2               , 4                )    THEN a.source_addr
	  END  AS source_addr
	, CASE WHEN b.id IS NULL                                       THEN a.matchrate
	       WHEN b.typ  = 3                                         THEN c.matchrate
	       WHEN b.typ IN ( 2               , 4                )    THEN a.matchrate
	  END  AS matchrate
	, CASE WHEN b.id IS NULL                                       THEN a.matchmoney
	       WHEN b.typ  = 3                                         THEN c.matchmoney
	       WHEN b.typ IN ( 2               , 4                )    THEN a.matchmoney
	  END  AS matchmoney
	, CASE WHEN b.id IS NULL                                       THEN a.ledger
	       WHEN b.typ  = 3                                         THEN c.ledger
	       WHEN b.typ IN ( 2               , 4                )    THEN a.ledger
	  END  AS ledger
	, CASE WHEN b.id IS NULL                                       THEN a.servicemoney
	       WHEN b.typ  = 3                                         THEN c.servicemoney
	       WHEN b.typ IN ( 2               , 4                )    THEN a.servicemoney
	  END  AS servicemoney
	, CASE WHEN b.id IS NULL                                       THEN a.repay
	       WHEN b.typ  = 3                                         THEN c.repay
	       WHEN b.typ IN ( 2               , 4                )    THEN a.repay
	  END  AS repay
	, CASE WHEN b.id IS NULL                                       THEN a.realrepay
	       WHEN b.typ  = 3                                         THEN c.realrepay
	       WHEN b.typ IN ( 2               , 4                )    THEN a.realrepay
	  END  AS realrepay
	, CASE WHEN b.id IS NULL                                       THEN a.trxid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.trxid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.trxid
	  END  AS trxid
	, CASE WHEN b.id IS NULL                                       THEN a.repaytype
	       WHEN b.typ  = 3                                         THEN c.repaytype
	       WHEN b.typ IN ( 2               , 4                )    THEN a.repaytype
	  END  AS repaytype
	, CASE WHEN b.id IS NULL                                       THEN a.orddatebid
	       WHEN b.typ  = 3                                         THEN c.orddatebid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.orddatebid
	  END  AS orddatebid
	, CASE WHEN b.id IS NULL                                       THEN a.ordidbid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.ordidbid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.ordidbid
	  END  AS ordidbid
	, CASE WHEN b.id IS NULL                                       THEN a.loans_orddate
	       WHEN b.typ  = 3                                         THEN c.loans_orddate
	       WHEN b.typ IN ( 2               , 4                )    THEN a.loans_orddate
	  END  AS loans_orddate
	, CASE WHEN b.id IS NULL                                       THEN a.loans_ordid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.loans_ordid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.loans_ordid
	  END  AS loans_ordid
	, CASE WHEN b.id IS NULL                                       THEN a.repay_orddate
	       WHEN b.typ  = 3                                         THEN c.repay_orddate
	       WHEN b.typ IN ( 2               , 4                )    THEN a.repay_orddate
	  END  AS repay_orddate
	, CASE WHEN b.id IS NULL                                       THEN a.repay_ordid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.repay_ordid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.repay_ordid
	  END  AS repay_ordid
	, CASE WHEN b.id IS NULL                                       THEN a.frezzeordid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.frezzeordid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.frezzeordid
	  END  AS frezzeordid
	, CASE WHEN b.id IS NULL                                       THEN a.lendtime
	       WHEN b.typ  = 3                                         THEN c.lendtime
	       WHEN b.typ IN ( 2               , 4                )    THEN a.lendtime
	  END  AS lendtime
	, CASE WHEN b.id IS NULL                                       THEN a.autobidtime
	       WHEN b.typ  = 3                                         THEN c.autobidtime
	       WHEN b.typ IN ( 2               , 4                )    THEN a.autobidtime
	  END  AS autobidtime
	, CASE WHEN b.id IS NULL                                       THEN a.statusrtime
	       WHEN b.typ  = 3                                         THEN c.statusrtime
	       WHEN b.typ IN ( 2               , 4                )    THEN a.statusrtime
	  END  AS statusrtime
	, CASE WHEN b.id IS NULL                                       THEN a.is_autobid
	       WHEN b.typ  = 3                                         THEN c.is_autobid
	       WHEN b.typ IN ( 2               , 4                )    THEN a.is_autobid
	  END  AS is_autobid
	, CASE WHEN b.id IS NULL                                       THEN a.lendstatus
	       WHEN b.typ  = 3                                         THEN c.lendstatus
	       WHEN b.typ IN ( 2               , 4                )    THEN a.lendstatus
	  END  AS lendstatus
	, CASE WHEN b.id IS NULL                                       THEN a.status
	       WHEN b.typ  = 3                                         THEN c.status
	       WHEN b.typ IN ( 2               , 4                )    THEN a.status
	  END  AS status
	, CASE WHEN b.id IS NULL                                       THEN a.is_transfer
	       WHEN b.typ  = 3                                         THEN c.is_transfer
	       WHEN b.typ IN ( 2               , 4                )    THEN a.is_transfer
	  END  AS is_transfer
	, CASE WHEN b.id IS NULL                                       THEN a.is_newcont
	       WHEN b.typ  = 3                                         THEN c.is_newcont
	       WHEN b.typ IN ( 2               , 4                )    THEN a.is_newcont
	  END  AS is_newcont
	, CASE WHEN b.id IS NULL                                       THEN a.newcontnum
	       WHEN b.typ  = 3                                         THEN c.newcontnum
	       WHEN b.typ IN ( 2               , 4                )    THEN a.newcontnum
	  END  AS newcontnum
	, CASE WHEN b.id IS NULL                                       THEN a.newcontid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.newcontid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.newcontid
	  END  AS newcontid
	, CASE WHEN b.id IS NULL                                       THEN a.transfermoney
	       WHEN b.typ  = 3                                         THEN c.transfermoney
	       WHEN b.typ IN ( 2               , 4                )    THEN a.transfermoney
	  END  AS transfermoney
	, CASE WHEN b.id IS NULL                                       THEN a.transferamt
	       WHEN b.typ  = 3                                         THEN c.transferamt
	       WHEN b.typ IN ( 2               , 4                )    THEN a.transferamt
	  END  AS transferamt
	, CASE WHEN b.id IS NULL                                       THEN a.service
	       WHEN b.typ  = 3                                         THEN c.service
	       WHEN b.typ IN ( 2               , 4                )    THEN a.service
	  END  AS service
	, CASE WHEN b.id IS NULL                                       THEN a.newcrmid
	       WHEN b.typ  = 3                                         THEN COALESCE( c.newcrmid , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.newcrmid
	  END  AS newcrmid
	, CASE WHEN b.id IS NULL                                       THEN a.transfertime
	       WHEN b.typ  = 3                                         THEN c.transfertime
	       WHEN b.typ IN ( 2               , 4                )    THEN a.transfertime
	  END  AS transfertime
	, CASE WHEN b.id IS NULL                                       THEN a.transferstatus
	       WHEN b.typ  = 3                                         THEN c.transferstatus
	       WHEN b.typ IN ( 2               , 4                )    THEN a.transferstatus
	  END  AS transferstatus
	, CASE WHEN b.id IS NULL                                       THEN a.newconturl
	       WHEN b.typ  = 3                                         THEN COALESCE( c.newconturl , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.newconturl
	  END  AS newconturl
	, CASE WHEN b.id IS NULL                                       THEN a.addtime
	       WHEN b.typ  = 3                                         THEN UNIX_TIMESTAMP(c.addtime)
	       WHEN b.typ IN ( 2               , 4                )    THEN a.addtime
	  END  AS addtime
	, CASE WHEN b.id IS NULL                                       THEN a.reclaim_amount
	       WHEN b.typ  = 3                                         THEN COALESCE( c.reclaim_amount , '')
	       WHEN b.typ IN ( 2               , 4                )    THEN a.reclaim_amount
	  END  AS reclaim_amount
	, a.ver       AS version  
	, a.start_dt  AS start_dt 
	, CASE WHEN b.id IS NULL                                       THEN a.end_dt    -- his             
	       WHEN b.typ = 3                                          THEN '3000-12-31'                    
	       WHEN b.typ IN ( 2                , 4                )   THEN '2016-01-01'                     
	       ELSE '0001-01-01' -- never used                                                                
	   END    As end_dt                                                                                    
  , a.load_stamp                                                                                        
 FROM       default.s_hyr_borrowlist_h   a          
 LEFT JOIN  default.vt_borrowlist_catlog       b          
  ON a.id = b.id
   -- AND b.`typ` IN ('closed', 'changed', 'unchanged')                                                 
  AND a.start_dt < '2016-01-01' -- support reload operation                                               
  AND a.end_dt  >= '2016-01-01' -- support reload operation                                               
 LEFT JOIN  default.borrowlist                                           c         
  ON b.id = c.id
 WHERE a.start_dt < '2016-01-01'  -- contains his-data && support reload operation 
;

 -- changed(from src)
INSERT INTO default.vt_s_hyr_borrowlist_h 
(
	  id
	, lendid
	, crmlc_uid
	, uid
	, usrcustid
	, borrowid
	, crmdk_uid
	, dkuid
	, dkusrcustid
	, loanmatcterm
	, crmlcmatcterm
	, fromtype
	, source_addr
	, matchrate
	, matchmoney
	, ledger
	, servicemoney
	, repay
	, realrepay
	, trxid
	, repaytype
	, orddatebid
	, ordidbid
	, loans_orddate
	, loans_ordid
	, repay_orddate
	, repay_ordid
	, frezzeordid
	, lendtime
	, autobidtime
	, statusrtime
	, is_autobid
	, lendstatus
	, status
	, is_transfer
	, is_newcont
	, newcontnum
	, newcontid
	, transfermoney
	, transferamt
	, service
	, newcrmid
	, transfertime
	, transferstatus
	, newconturl
	, addtime
	, reclaim_amount
	, version    
	, start_dt   
	, end_dt     
	, load_stamp 
)
SELECT 
	  a.id                           as id
	, a.lendid                       as lendid
	, a.crmlc_uid                    as crmlc_uid
	, a.uid                          as uid
	, COALESCE( a.usrcustid , '')    as usrcustid
	, a.borrowid                     as borrowid
	, a.crmdk_uid                    as crmdk_uid
	, a.dkuid                        as dkuid
	, COALESCE( a.dkusrcustid , '')  as dkusrcustid
	, a.loanmatcterm                 as loanmatcterm
	, a.crmlcmatcterm                as crmlcmatcterm
	, a.fromtype                     as fromtype
	, a.source_addr                  as source_addr
	, a.matchrate                    as matchrate
	, a.matchmoney                   as matchmoney
	, a.ledger                       as ledger
	, a.servicemoney                 as servicemoney
	, a.repay                        as repay
	, a.realrepay                    as realrepay
	, COALESCE( a.trxid , '')        as trxid
	, a.repaytype                    as repaytype
	, a.orddatebid                   as orddatebid
	, COALESCE( a.ordidbid , '')     as ordidbid
	, a.loans_orddate                as loans_orddate
	, COALESCE( a.loans_ordid , '')  as loans_ordid
	, a.repay_orddate                as repay_orddate
	, COALESCE( a.repay_ordid , '')  as repay_ordid
	, COALESCE( a.frezzeordid , '')  as frezzeordid
	, a.lendtime                     as lendtime
	, a.autobidtime                  as autobidtime
	, a.statusrtime                  as statusrtime
	, a.is_autobid                   as is_autobid
	, a.lendstatus                   as lendstatus
	, a.status                       as status
	, a.is_transfer                  as is_transfer
	, a.is_newcont                   as is_newcont
	, a.newcontnum                   as newcontnum
	, COALESCE( a.newcontid , '')    as newcontid
	, a.transfermoney                as transfermoney
	, a.transferamt                  as transferamt
	, a.service                      as service
	, COALESCE( a.newcrmid , '')     as newcrmid
	, a.transfertime                 as transfertime
	, a.transferstatus               as transferstatus
	, COALESCE( a.newconturl , '')   as newconturl
	, UNIX_TIMESTAMP(a.addtime)      as addtime
	, COALESCE( a.reclaim_amount , '')         as reclaim_amount
	, b.ver + 1         AS version   
	, '2016-01-01'          AS start_dt  
	, '3000-12-31'         AS end_dt    
	, CURRENT_TIMESTAMP              
FROM       default.borrowlist                 a
INNER JOIN default.vt_borrowlist_catlog                     b
  ON a.id = b.id
  AND b.`typ` = 4 /* 'changed' */       
 ;

-- Just for test - bakup
${MYSQL_COMMENT} CREATE TABLE IF NOT EXISTS default.s_hyr_borrowlist_h${MY_DT} AS SELECT * FROM default.s_hyr_borrowlist_h;
TRUNCATE TABLE default.s_hyr_borrowlist_h;
INSERT INTO default.s_hyr_borrowlist_h SELECT * FROM default.vt_s_hyr_borrowlist_h;
