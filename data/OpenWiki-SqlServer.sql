CREATE TABLE [openwiki_attachments] (
    [att_wrv_name] [nvarchar] (128) NOT NULL ,
    [att_wrv_revision] [int] NOT NULL ,
    [att_name] [nvarchar] (255) NOT NULL ,
    [att_revision] [int] NOT NULL ,
    [att_hidden] [int] NOT NULL ,
    [att_deprecated] [int] NOT NULL ,
    [att_filename] [nvarchar] (255) NOT NULL ,
    [att_timestamp] [datetime] NOT NULL ,
    [att_filesize] [int] NOT NULL ,
    [att_host] [nvarchar] (128) NULL ,
    [att_agent] [nvarchar] (255) NULL ,
    [att_by] [nvarchar] (128) NULL ,
    [att_byalias] [nvarchar] (128) NULL ,
    [att_comment] [text] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [openwiki_attachments_log] (
    [ath_wrv_name] [nvarchar] (128) NOT NULL ,
    [ath_wrv_revision] [int] NOT NULL ,
    [ath_name] [nvarchar] (255) NOT NULL ,
    [ath_revision] [int] NOT NULL ,
    [ath_timestamp] [datetime] NOT NULL ,
    [ath_agent] [nvarchar] (255) NULL ,
    [ath_by] [nvarchar] (128) NULL ,
    [ath_byalias] [nvarchar] (128) NULL ,
    [ath_action] [nvarchar] (20) NOT NULL
) ON [PRIMARY]
GO

CREATE TABLE [openwiki_cache] (
    [chc_name] [nvarchar] (128) NOT NULL ,
    [chc_hash] [int] NOT NULL ,
    [chc_xmlisland] [text] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [openwiki_interwikis] (
    [wik_name] [nvarchar] (128) NOT NULL ,
    [wik_url] [nvarchar] (255) NOT NULL
) ON [PRIMARY]
GO

CREATE TABLE [openwiki_pages] (
    [wpg_name] [nvarchar] (128) NOT NULL ,
    [wpg_lastmajor] [int] NOT NULL ,
    [wpg_lastminor] [int] NOT NULL ,
    [wpg_changes] [int] NOT NULL
) ON [PRIMARY]
GO

CREATE TABLE [openwiki_revisions] (
    [wrv_name] [nvarchar] (128) NOT NULL ,
    [wrv_revision] [int] NOT NULL ,
    [wrv_current] [int] NOT NULL ,
    [wrv_status] [int] NOT NULL ,
    [wrv_timestamp] [datetime] NOT NULL ,
    [wrv_minoredit] [int] NOT NULL ,
    [wrv_host] [nvarchar] (128) NULL ,
    [wrv_agent] [nvarchar] (255) NULL ,
    [wrv_by] [nvarchar] (128) NULL ,
    [wrv_byalias] [nvarchar] (128) NULL ,
    [wrv_comment] [nvarchar] (1024) NULL ,
    [wrv_text] [ntext] NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [openwiki_rss] (
    [rss_url] [nvarchar] (256) NOT NULL ,
    [rss_last] [datetime] NOT NULL ,
    [rss_next] [datetime] NOT NULL ,
    [rss_refreshrate] [int] NOT NULL ,
    [rss_cache] [ntext] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

CREATE TABLE [openwiki_rss_aggregations] (
    [agr_feed] [nvarchar] (200) NOT NULL ,
    [agr_resource] [nvarchar] (200) NOT NULL ,
    [agr_rsslink] [nvarchar] (200) NULL ,
    [agr_timestamp] [datetime] NOT NULL ,
    [agr_dcdate] [nvarchar] (25) NULL ,
    [agr_xmlisland] [ntext] NOT NULL
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]
GO

ALTER TABLE [openwiki_pages] WITH NOCHECK ADD
    CONSTRAINT [PK_openwiki_pages] PRIMARY KEY  CLUSTERED
    (
        [wpg_name]
    ) WITH  FILLFACTOR = 90  ON [PRIMARY]
GO

ALTER TABLE [openwiki_revisions] WITH NOCHECK ADD
    CONSTRAINT [PK_openwiki_revisions] PRIMARY KEY  CLUSTERED
    (
        [wrv_name],
        [wrv_revision]
    ) WITH  FILLFACTOR = 90  ON [PRIMARY]
GO

ALTER TABLE [openwiki_attachments] WITH NOCHECK ADD
    CONSTRAINT [PK_openwiki_attachments] PRIMARY KEY  NONCLUSTERED
    (
        [att_wrv_name],
        [att_name],
        [att_revision]
    ) WITH  FILLFACTOR = 90  ON [PRIMARY]
GO

ALTER TABLE [openwiki_cache] WITH NOCHECK ADD
    CONSTRAINT [PK_openwiki_cache] PRIMARY KEY  NONCLUSTERED
    (
        [chc_name],
        [chc_hash]
    ) WITH  FILLFACTOR = 90  ON [PRIMARY]
GO

ALTER TABLE [openwiki_interwikis] WITH NOCHECK ADD
    CONSTRAINT [PK_openwiki_interwikis] PRIMARY KEY  NONCLUSTERED
    (
        [wik_name]
    ) WITH  FILLFACTOR = 90  ON [PRIMARY]
GO

ALTER TABLE [openwiki_rss] WITH NOCHECK ADD
    CONSTRAINT [PK_openwiki_rss] PRIMARY KEY  NONCLUSTERED
    (
        [rss_url]
    ) WITH  FILLFACTOR = 90  ON [PRIMARY]
GO

ALTER TABLE [openwiki_rss_aggregations] WITH NOCHECK ADD
    CONSTRAINT [PK_openwiki_rss_aggregations] PRIMARY KEY  NONCLUSTERED
    (
        [agr_feed],
        [agr_resource]
    ) WITH  FILLFACTOR = 90  ON [PRIMARY]
GO

ALTER TABLE [openwiki_attachments] ADD
    CONSTRAINT [FK_openwiki_attachments_openwiki_revisions] FOREIGN KEY
    (
        [att_wrv_name],
        [att_wrv_revision]
    ) REFERENCES [openwiki_revisions] (
        [wrv_name],
        [wrv_revision]
    )
GO

ALTER TABLE [openwiki_attachments_log] ADD
    CONSTRAINT [FK_openwiki_attachments_log_openwiki_revisions] FOREIGN KEY
    (
        [ath_wrv_name],
        [ath_wrv_revision]
    ) REFERENCES [openwiki_revisions] (
        [wrv_name],
        [wrv_revision]
    )
GO

ALTER TABLE [openwiki_revisions] ADD
    CONSTRAINT [FK_openwiki_revisions_openwiki_pages] FOREIGN KEY
    (
        [wrv_name]
    ) REFERENCES [openwiki_pages] (
        [wpg_name]
    )
GO

