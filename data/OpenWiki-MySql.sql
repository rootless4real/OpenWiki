CREATE TABLE openwiki_attachments (
    att_wrv_name     VARCHAR(128) NOT NULL,
    att_wrv_revision INT          NOT NULL,
    att_name         VARCHAR(255) NOT NULL,
    att_revision     INT          NOT NULL,
    att_hidden       INT          NOT NULL,
    att_deprecated   INT          NOT NULL,
    att_filename     VARCHAR(255) NOT NULL,
    att_timestamp    DATETIME     NOT NULL,
    att_filesize     INT          NOT NULL,
    att_host         VARCHAR(128),
    att_agent        VARCHAR(255),
    att_by           VARCHAR(128),
    att_byalias      VARCHAR(128),
    att_comment      TEXT
)
;

CREATE TABLE openwiki_attachments_log (
    ath_wrv_name     VARCHAR(128) NOT NULL,
    ath_wrv_revision INT          NOT NULL,
    ath_name         VARCHAR(255) NOT NULL,
    ath_revision     INT          NOT NULL,
    ath_timestamp    DATETIME     NOT NULL,
    ath_action       VARCHAR(20)  NOT NULL,
    ath_agent        VARCHAR(255),
    ath_by           VARCHAR(128),
    ath_byalias      VARCHAR(128)
)
;

CREATE TABLE openwiki_cache (
    chc_name         VARCHAR(128) NOT NULL,
    chc_hash         INT          NOT NULL,
    chc_xmlisland    TEXT         NOT NULL
)
;

CREATE TABLE openwiki_interwikis (
    wik_name         VARCHAR(128) NOT NULL,
    wik_url          VARCHAR(255) NOT NULL
)
;

CREATE TABLE openwiki_pages (
    wpg_name         VARCHAR(128) NOT NULL,
    wpg_lastmajor    INT          NOT NULL,
    wpg_lastminor    INT          NOT NULL,
    wpg_changes      INT          NOT NULL
)
;

CREATE TABLE openwiki_revisions (
    wrv_name         VARCHAR(128) NOT NULL,
    wrv_revision     INT          NOT NULL,
    wrv_current      INT          NOT NULL,
    wrv_status       INT          NOT NULL,
    wrv_timestamp    DATETIME     NOT NULL,
    wrv_minoredit    INT          NOT NULL,
    wrv_host         VARCHAR(128),
    wrv_agent        VARCHAR(255),
    wrv_by           VARCHAR(128),
    wrv_byalias      VARCHAR(128),
    wrv_comment      TEXT,
    wrv_text         TEXT
)
;

CREATE TABLE openwiki_rss (
    rss_url          VARCHAR(255) NOT NULL,
    rss_last         DATETIME     NOT NULL,
    rss_next         DATETIME     NOT NULL,
    rss_refreshrate  INT          NOT NULL,
    rss_cache        TEXT         NOT NULL
)
;

CREATE TABLE openwiki_rss_aggregations (
    agr_feed         VARCHAR(200) NOT NULL,
    agr_resource     VARCHAR(200) NOT NULL,
    agr_rsslink      VARCHAR(200) NOT NULL,
    agr_timestamp    DATETIME     NOT NULL,
    agr_dcdate       VARCHAR(25),
    agr_xmlisland    TEXT         NOT NULL
)
;


# ALTER TABLE openwiki_attachments ADD PRIMARY KEY (att_wrv_name, att_name, att_revision)
# ;
ALTER TABLE openwiki_cache ADD PRIMARY KEY (chc_name, chc_hash)
;
ALTER TABLE openwiki_pages ADD PRIMARY KEY (wpg_name)
;
ALTER TABLE openwiki_revisions ADD PRIMARY KEY (wrv_name, wrv_revision)
;
ALTER TABLE openwiki_rss ADD PRIMARY KEY (rss_url)
;
#ALTER TABLE openwiki_rss_aggregations ADD PRIMARY KEY (agr_feed, agr_resource)
#;
ALTER TABLE openwiki_attachments ADD CONSTRAINT FK_att_wrv FOREIGN KEY (att_wrv_name, att_wrv_revision) REFERENCES openwiki_revisions (wrv_name, wrv_revision)
;
ALTER TABLE openwiki_attachments_log ADD CONSTRAINT FK_ath_wrv FOREIGN KEY (ath_wrv_name, ath_wrv_revision) REFERENCES openwiki_revisions (wrv_name, wrv_revision)
;
ALTER TABLE openwiki_revisions ADD CONSTRAINT FK_wrv_wpg FOREIGN KEY (wrv_name) REFERENCES openwiki_pages (wpg_name)
;
