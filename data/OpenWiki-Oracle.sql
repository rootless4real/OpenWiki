CREATE TABLE openwiki_attachments (
    att_wrv_name     VARCHAR2(128) NOT NULL,
    att_wrv_revision NUMBER(9)     NOT NULL,
    att_name         VARCHAR2(255) NOT NULL,
    att_revision     NUMBER(9)     NOT NULL,
    att_hidden       NUMBER(9)     NOT NULL,
    att_deprecated   NUMBER(9)     NOT NULL,
    att_filename     VARCHAR2(255) NOT NULL,
    att_timestamp    DATE          NOT NULL,
    att_filesize     NUMBER(9)     NOT NULL,
    att_host         VARCHAR2(128),
    att_agent        VARCHAR2(255),
    att_by           VARCHAR2(128),
    att_byalias      VARCHAR2(128),
    att_comment      LONG
)
;

CREATE TABLE openwiki_attachments_log (
    ath_wrv_name     VARCHAR2(128) NOT NULL,
    ath_wrv_revision NUMBER(9)     NOT NULL,
    ath_name         VARCHAR2(255) NOT NULL,
    ath_revision     NUMBER(9)     NOT NULL,
    ath_timestamp    DATE          NOT NULL,
    ath_action       VARCHAR2(20)  NOT NULL,
    ath_agent        VARCHAR2(255),
    ath_by           VARCHAR2(128),
    ath_byalias      VARCHAR2(128)
)
;

CREATE TABLE openwiki_cache (
    chc_name         VARCHAR2(128) NOT NULL,
    chc_hash         NUMBER(9)     NOT NULL,
    chc_xmlisland    LONG          NOT NULL
)
;

CREATE TABLE openwiki_interwikis (
    wik_name         VARCHAR2(128) NOT NULL,
    wik_url          VARCHAR2(255) NOT NULL
)
;

CREATE TABLE openwiki_pages (
    wpg_name         VARCHAR2(128) NOT NULL,
    wpg_lastmajor    NUMBER(9)     NOT NULL,
    wpg_lastminor    NUMBER(9)     NOT NULL,
    wpg_changes      NUMBER(9)     NOT NULL
)
;

CREATE TABLE openwiki_revisions (
    wrv_name         VARCHAR2(128) NOT NULL,
    wrv_revision     NUMBER(9)     NOT NULL,
    wrv_current      NUMBER(9)     NOT NULL,
    wrv_status       NUMBER(9)     NOT NULL,
    wrv_timestamp    DATE          NOT NULL,
    wrv_minoredit    NUMBER(9)     NOT NULL,
    wrv_host         VARCHAR2(128),
    wrv_agent        VARCHAR2(255),
    wrv_by           VARCHAR2(128),
    wrv_byalias      VARCHAR2(128),
    wrv_comment      VARCHAR2(1000),
    wrv_text         LONG
)
;

CREATE TABLE openwiki_rss (
    rss_url          VARCHAR2(255) NOT NULL,
    rss_last         DATE          NOT NULL,
    rss_next         DATE          NOT NULL,
    rss_refreshrate  NUMBER(9)     NOT NULL,
    rss_cache        LONG          NOT NULL
)
;

CREATE TABLE openwiki_rss_aggregations (
    agr_feed         VARCHAR2(200) NOT NULL,
    agr_resource     VARCHAR2(200) NOT NULL,
    agr_rsslink      VARCHAR2(200) NOT NULL,
    agr_timestamp    DATE          NOT NULL,
    agr_dcdate       VARCHAR2(25),
    agr_xmlisland    LONG          NOT NULL
)
;


ALTER TABLE openwiki_attachments ADD PRIMARY KEY (att_wrv_name, att_name, att_revision)
;
ALTER TABLE openwiki_cache ADD PRIMARY KEY (chc_name, chc_hash)
;
ALTER TABLE openwiki_pages ADD PRIMARY KEY (wpg_name)
;
ALTER TABLE openwiki_revisions ADD PRIMARY KEY (wrv_name, wrv_revision)
;
ALTER TABLE openwiki_rss ADD PRIMARY KEY (rss_url)
;
ALTER TABLE openwiki_rss_aggregations ADD PRIMARY KEY (agr_feed, agr_resource)
;
ALTER TABLE openwiki_attachments ADD CONSTRAINT FK_att_arv FOREIGN KEY (att_wrv_name, att_wrv_revision) REFERENCES openwiki_revisions (wrv_name, wrv_revision)
;
ALTER TABLE openwiki_attachments_log ADD CONSTRAINT FK_ath_wrv FOREIGN KEY (ath_wrv_name, ath_wrv_revision) REFERENCES openwiki_revisions (wrv_name, wrv_revision)
;
ALTER TABLE openwiki_revisions ADD CONSTRAINT FK_wrv_wpg FOREIGN KEY (wrv_name) REFERENCES openwiki_pages (wpg_name)
;
