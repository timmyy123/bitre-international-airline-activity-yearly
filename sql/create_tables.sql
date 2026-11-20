-- =====================================================================
-- BITRE International Airline Activity — yearly aggregation
-- MySQL 8.0+ schema for the sheets in
--   output/International_airline_activity_yearly_ALL.xlsx
--
-- One table per data sheet; column names follow the sheet headers
-- (lower-cased). Exception: Table 4's "In"/"Out" are reserved words in
-- MySQL, so they are named inbound / outbound here.
--
-- Values:
--   NULL  = '..' in BITRE's tables (not available, e.g. passengers on
--           freight-only airlines)
--   0     = '-' (nil)
--   *_t   = tonnes
--   *_pct = percentage points (82.17 means 82.17%), EXCEPT in
--           monthly_series, where the sheet stores fractions (0.8217).
--
-- Tables 1–3 and 5 mix detail rows with subtotal / total rows
-- (row_type). Filter row_type = 'detail' before SUM()/GROUP BY,
-- or you will double-count.
-- =====================================================================

CREATE DATABASE IF NOT EXISTS airline_activity
  CHARACTER SET utf8mb4 COLLATE utf8mb4_0900_ai_ci;
USE airline_activity;

-- ---------------------------------------------------------------------
-- Summary: headline totals by year
-- ---------------------------------------------------------------------
CREATE TABLE summary_by_year (
  year                               SMALLINT UNSIGNED NOT NULL,
  months_included                    TINYINT  UNSIGNED NOT NULL COMMENT '12 = full year; 2017 = 6 (Jul–Dec), 2026 = 7 (Jan–Jul)',
  passengers_in                      INT UNSIGNED      NOT NULL,
  passengers_out                     INT UNSIGNED      NOT NULL,
  passengers_total                   INT UNSIGNED      NOT NULL,
  passengers_same_months_prev_year   INT UNSIGNED      NULL,
  passenger_change_pct               DECIMAL(12,4)     NULL COMMENT 'percent, vs same months of previous year',
  freight_t_total                    DECIMAL(14,3)     NULL,
  mail_t_total                       DECIMAL(14,3)     NULL,
  flights_total                      INT UNSIGNED      NULL,
  seats_total                        INT UNSIGNED      NULL,
  seat_util_pct                      DECIMAL(7,4)      NULL COMMENT 'percent',
  operators_with_traffic             SMALLINT UNSIGNED NULL,
  PRIMARY KEY (year)
) ENGINE=InnoDB COMMENT='Sheet: Summary';

-- ---------------------------------------------------------------------
-- Monthly_series: one row per month (Jul 2017 – Jul 2026)
-- ---------------------------------------------------------------------
CREATE TABLE monthly_series (
  year                               SMALLINT UNSIGNED NOT NULL,
  month                              CHAR(3)           NOT NULL COMMENT 'Jan … Dec, as in the sheet',
  month_num                          TINYINT UNSIGNED  GENERATED ALWAYS AS
                                       (FIELD(month,'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec')) STORED,
  period_start                       DATE              GENERATED ALWAYS AS
                                       (MAKEDATE(year, 1) + INTERVAL (FIELD(month,'Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec') - 1) MONTH) STORED,
  passengers_in                      INT UNSIGNED      NOT NULL,
  passengers_out                     INT UNSIGNED      NOT NULL,
  passengers_total                   INT UNSIGNED      NOT NULL,
  passengers_same_month_prev_year    INT UNSIGNED      NULL,
  passenger_change_pct               DECIMAL(12,6)     NULL COMMENT 'FRACTION (0.05 = 5%)',
  freight_t                          DECIMAL(14,3)     NULL,
  mail_t                             DECIMAL(14,3)     NULL,
  available_seats                    INT UNSIGNED      NULL,
  flights                            INT UNSIGNED      NULL,
  aircraft_movements                 INT UNSIGNED      NULL,
  seat_util_pct                      DECIMAL(9,6)      NULL COMMENT 'FRACTION (0.82 = 82%)',
  PRIMARY KEY (year, month),
  UNIQUE KEY uq_period (period_start)
) ENGINE=InnoDB COMMENT='Sheet: Monthly_series';

-- ---------------------------------------------------------------------
-- Table 1: passengers, freight and mail by operator and country
-- ---------------------------------------------------------------------
CREATE TABLE t1_traffic_by_operator (
  id                 INT UNSIGNED      NOT NULL AUTO_INCREMENT,
  year               SMALLINT UNSIGNED NOT NULL,
  months_included    TINYINT  UNSIGNED NOT NULL,
  operator           VARCHAR(100)      NOT NULL COMMENT '"TOTAL" on the grand-total row',
  country            VARCHAR(100)      NULL     COMMENT '"ALL SERVICES" on operator subtotals; NULL on the grand total',
  in_passengers      INT UNSIGNED      NULL,
  in_freight_t       DECIMAL(14,3)     NULL,
  in_mail_t          DECIMAL(14,3)     NULL,
  out_passengers     INT UNSIGNED      NULL,
  out_freight_t      DECIMAL(14,3)     NULL,
  out_mail_t         DECIMAL(14,3)     NULL,
  total_passengers   INT UNSIGNED      NULL,
  total_freight_t    DECIMAL(14,3)     NULL,
  total_mail_t       DECIMAL(14,3)     NULL,
  row_type           ENUM('detail','subtotal','total') NOT NULL,
  PRIMARY KEY (id),
  UNIQUE KEY uq_t1 (year, operator, country),
  KEY ix_t1_type (row_type, year),
  KEY ix_t1_country (country, year)
) ENGINE=InnoDB COMMENT='Sheet: T1_Traffic_by_operator';

-- ---------------------------------------------------------------------
-- Table 2: operator market shares and growth
-- ---------------------------------------------------------------------
CREATE TABLE t2_market_share (
  id                   INT UNSIGNED      NOT NULL AUTO_INCREMENT,
  year                 SMALLINT UNSIGNED NOT NULL,
  months_included      TINYINT  UNSIGNED NOT NULL,
  operator             VARCHAR(100)      NOT NULL,
  country              VARCHAR(100)      NULL,
  pax_prev             INT UNSIGNED      NULL COMMENT 'same months of previous year',
  pax                  INT UNSIGNED      NULL,
  pax_share_pct        DECIMAL(7,4)      NULL COMMENT 'percent of TOTAL',
  pax_change_pct       DECIMAL(14,4)     NULL COMMENT 'percent',
  freight_prev         DECIMAL(14,3)     NULL,
  freight              DECIMAL(14,3)     NULL,
  freight_share_pct    DECIMAL(7,4)      NULL,
  freight_change_pct   DECIMAL(14,4)     NULL,
  mail_prev            DECIMAL(14,3)     NULL,
  mail                 DECIMAL(14,3)     NULL,
  mail_share_pct       DECIMAL(7,4)      NULL,
  mail_change_pct      DECIMAL(14,4)     NULL,
  row_type             ENUM('detail','subtotal','total') NOT NULL,
  PRIMARY KEY (id),
  UNIQUE KEY uq_t2 (year, operator, country),
  KEY ix_t2_type (row_type, year)
) ENGINE=InnoDB COMMENT='Sheet: T2_Market_share';

-- ---------------------------------------------------------------------
-- Table 3: passenger capacity and seat utilisation by operator
-- ---------------------------------------------------------------------
CREATE TABLE t3_capacity_utilisation (
  id                   INT UNSIGNED      NOT NULL AUTO_INCREMENT,
  year                 SMALLINT UNSIGNED NOT NULL,
  months_included      TINYINT  UNSIGNED NOT NULL,
  operator             VARCHAR(100)      NOT NULL,
  service              VARCHAR(100)      NULL COMMENT 'service to/from; "ALL SERVICES" on subtotals',
  in_flights           INT UNSIGNED      NULL,
  in_pax               INT UNSIGNED      NULL,
  in_seats             INT UNSIGNED      NULL,
  in_seat_util_pct     DECIMAL(7,4)      NULL COMMENT 'percent',
  out_flights          INT UNSIGNED      NULL,
  out_pax              INT UNSIGNED      NULL,
  out_seats            INT UNSIGNED      NULL,
  out_seat_util_pct    DECIMAL(7,4)      NULL COMMENT 'percent',
  row_type             ENUM('detail','subtotal','total') NOT NULL,
  PRIMARY KEY (id),
  UNIQUE KEY uq_t3 (year, operator, service),
  KEY ix_t3_type (row_type, year)
) ENGINE=InnoDB COMMENT='Sheet: T3_Capacity_utilisation';

-- ---------------------------------------------------------------------
-- Table 4: airport traffic and aircraft movements
-- (the sheet has no Row_type: the national total is airport = 'TOTAL')
-- ---------------------------------------------------------------------
CREATE TABLE t4_airports (
  id                   INT UNSIGNED      NOT NULL AUTO_INCREMENT,
  year                 SMALLINT UNSIGNED NOT NULL,
  months_included      TINYINT  UNSIGNED NOT NULL,
  airport              VARCHAR(100)      NOT NULL COMMENT '"TOTAL" = all airports',
  measure              VARCHAR(40)       NOT NULL COMMENT 'Passengers | Freight (Tonnes) | Aircraft Movements',
  inbound_prev         DECIMAL(16,3)     NULL COMMENT 'sheet column In_prev',
  inbound              DECIMAL(16,3)     NULL COMMENT 'sheet column In',
  inbound_change_pct   DECIMAL(14,4)     NULL COMMENT 'sheet column In_change_pct (percent)',
  outbound_prev        DECIMAL(16,3)     NULL COMMENT 'sheet column Out_prev',
  outbound             DECIMAL(16,3)     NULL COMMENT 'sheet column Out',
  outbound_change_pct  DECIMAL(14,4)     NULL COMMENT 'sheet column Out_change_pct (percent)',
  total_prev           DECIMAL(16,3)     NULL,
  total                DECIMAL(16,3)     NULL,
  total_change_pct     DECIMAL(14,4)     NULL COMMENT 'percent',
  share_of_total_pct   DECIMAL(7,4)      NULL COMMENT 'percent of the TOTAL row for the same measure',
  is_total             TINYINT(1)        GENERATED ALWAYS AS (airport = 'TOTAL') STORED,
  PRIMARY KEY (id),
  UNIQUE KEY uq_t4 (year, airport, measure),
  KEY ix_t4_measure (measure, year)
) ENGINE=InnoDB COMMENT='Sheet: T4_Airports';

-- ---------------------------------------------------------------------
-- Table 5: traffic by city pair (foreign port x Australian port)
--   row_type 'subtotal'             : foreign_port = 'Total' for one Australian port
--   row_type 'all Australian ports' : australian_port = 'Australia', one row per foreign port
--   row_type 'total'                : australian_port = 'Australia', foreign_port = 'Total'
-- ---------------------------------------------------------------------
CREATE TABLE t5_city_pairs (
  id                   INT UNSIGNED      NOT NULL AUTO_INCREMENT,
  year                 SMALLINT UNSIGNED NOT NULL,
  months_included      TINYINT  UNSIGNED NOT NULL,
  australian_port      VARCHAR(100)      NOT NULL,
  foreign_port         VARCHAR(100)      NOT NULL,
  pax_in_prev          INT UNSIGNED      NULL COMMENT 'same months of previous year',
  pax_out_prev         INT UNSIGNED      NULL,
  pax_total_prev       INT UNSIGNED      NULL,
  pax_in               INT UNSIGNED      NULL,
  pax_out              INT UNSIGNED      NULL,
  pax_total            INT UNSIGNED      NULL,
  frt_in_prev          DECIMAL(14,3)     NULL COMMENT 'freight, tonnes',
  frt_out_prev         DECIMAL(14,3)     NULL,
  frt_total_prev       DECIMAL(14,3)     NULL,
  frt_in               DECIMAL(14,3)     NULL,
  frt_out              DECIMAL(14,3)     NULL,
  frt_total            DECIMAL(14,3)     NULL,
  row_type             ENUM('detail','subtotal','all Australian ports','total') NOT NULL,
  PRIMARY KEY (id),
  UNIQUE KEY uq_t5 (year, australian_port, foreign_port),
  KEY ix_t5_foreign (foreign_port, year),
  KEY ix_t5_type (row_type, year)
) ENGINE=InnoDB COMMENT='Sheet: T5_City_pairs';
