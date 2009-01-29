-- index to improve performance of message retrieval in data transfer for Oracle Server databases
CREATE INDEX IDX_MES_SITERECDIR ON MESSAGE ( TRIALSITE, MESSAGERECEIVED, MESSAGEDIRECTION );