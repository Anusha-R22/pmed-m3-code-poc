-- index to improve performance of mimessage data transfer (proposed for) patch 3.0.77
CREATE INDEX IDX_MIM_MIMSITESOURCESENT ON MIMESSAGE (MIMESSAGESITE, MIMESSAGESOURCE, MIMESSAGESENT);
