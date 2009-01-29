-- index to improve performance of message retrieval in auto import for Oracle Server databases for patch 3.0.77
CREATE INDEX IDX_MES_DIRREC ON MESSAGE (MESSAGEDIRECTION, MESSAGERECEIVED);