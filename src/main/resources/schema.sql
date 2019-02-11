DROP TABLE IF EXISTS TEMPLATE_MASTER_TEST;
DROP TABLE IF EXISTS SHEET_MASTER_TEST;
DROP TABLE IF EXISTS HEADER_MASTER_TEST;

CREATE TABLE IF NOT EXISTS TEMPLATE_MASTER_TEST (
	TEMPLATE_ID INTEGER PRIMARY KEY,
	TEMPLATE_NAME VARCHAR(250)
);

CREATE TABLE  IF NOT EXISTS SHEET_MASTER_TEST (
	SHEET_ID INTEGER PRIMARY KEY,
	TEMPLATE_ID INTEGER,
	SHEET_NAME VARCHAR(250),
	FOREIGN KEY fk_cat(TEMPLATE_ID)
    REFERENCES TEMPLATE_MASTER_TEST(TEMPLATE_ID)
    ON UPDATE CASCADE
    ON DELETE CASCADE
);

CREATE TABLE  IF NOT EXISTS SHEET_HEADER_MASTER_TEST (
	HEADER_ID INTEGER PRIMARY KEY,
	SHEET_ID INTEGER,
	SHEET_INDEX INTEGER,
	ROW_FROM INTEGER,
	ROW_TO INTEGER,
	COL_FROM INTEGER,
	COL_TO INTEGER,
	HEADER_TEXT VARCHAR(250),
	FOREIGN KEY fk_cat(SHEET_ID)
    REFERENCES SHEET_MASTER_TEST(SHEET_ID)
    ON UPDATE CASCADE
    ON DELETE CASCADE
);