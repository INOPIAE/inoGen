ALTER TABLE tblPerson ADD COLUMN active YESNO;

UPDATE tblPerson SET active = 1;


ALTER TABLE tblFamilie ADD COLUMN active YESNO;

UPDATE tblFamilie SET active = 1;


ALTER TABLE tblEreignis ADD COLUMN active YESNO;

UPDATE tblEreignis SET active = 1;


UPDATE tblVersion SET Version = 10;
