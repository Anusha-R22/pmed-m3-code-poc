-- creates a correct seeded messageid sequence for upgrading to patch 3.0.70b and 3.0.73
BEGIN
  DECLARE seed NUMBER(16);
  BEGIN
    SELECT MAX(MESSAGEID)+1 INTO seed FROM MESSAGE;
    EXECUTE IMMEDIATE 'CREATE SEQUENCE SEQ_MESSAGEID INCREMENT BY 1 START WITH '|| TO_CHAR(seed);
  END;
END;
/