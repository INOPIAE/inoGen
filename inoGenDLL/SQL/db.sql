DROP TABLE tblVersion;

CREATE TABLE tblVersion (Version INTEGER);
INSERT INTO tblVersion (Version) VALUES (0);

DROP TABLE tblKonfession;

CREATE TABLE tblKonfession (
    tblKonfessionID COUNTER,
    Konfessionkurz VARCHAR(10), 
    Konfession VARCHAR(255),   
	CONSTRAINT PrimaryKey PRIMARY KEY (tblKonfessionID));

DROP TABLE tblKreis;

CREATE TABLE tblKreis (
    tblKreisID COUNTER,
    Kreis VARCHAR(255), 
    Info MEMO,   
	CONSTRAINT PrimaryKey PRIMARY KEY (tblKreisID));

DROP TABLE tblOrt;

CREATE TABLE tblOrt(
    tblOrtID COUNTER,
    Ort VARCHAR(255),
    tblKreisID INTEGER,
    Info MEMO,
    Breite DOUBLE,
    Laenge DOUBLE,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblOrtID));

DROP TABLE tblEreignisArt;

CREATE TABLE tblEreignisArt(
    tblEreignisArtID COUNTER,
    EreignisArt VARCHAR(255),
    Zeichen VARCHAR(10),
    Reihenfolge INTEGER,
    PersonenEreignis YESNO DEFAULT TRUE,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblEreignisArtID));

DROP TABLE tblPerson;

CREATE TABLE tblPerson(
    tblPersonID COUNTER,
    PS VARCHAR(12),
    Sex VARCHAR(1),
    FSID VARCHAR(10),
    tblFamilieID INTEGER,
    tblNachnameID INTEGER,
    tblKonfessionID INTEGER,
    Vorname VARCHAR(255),
    Info MEMO,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblPersonID));

DROP TABLE tblVorname;

CREATE TABLE tblVorname(
    tblVornameID COUNTER,
    Vorname VARCHAR(255),
	CONSTRAINT PrimaryKey PRIMARY KEY (tblVornameID));

DROP TABLE tblPVorname;

CREATE TABLE tblPVorname(
    tblPVornameID COUNTER,
    tblPersonID INTEGER,
    tblVornameID INTEGER,
    Zeichen VARCHAR(1),
    Reihenfolge INTEGER,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblPVornameID));


DROP TABLE tblNachname;

CREATE TABLE tblNachname(
    tblNachnameID COUNTER,
    Nachname VARCHAR(255),
	CONSTRAINT PrimaryKey PRIMARY KEY (tblNachnameID));

DROP TABLE tblFamilie;

CREATE TABLE tblFamilie(
    tblFamilieID COUNTER,
    FS VARCHAR(12),
    tblPersonIDV INTEGER,
    tblPersonIDM INTEGER,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblFamilieID));

DROP TABLE tblEreignis;

CREATE TABLE tblEreignis(
    tblEreignisID COUNTER,
    tblEreignisArtID INTEGER,
    tblPersonID INTEGER,
    tblFamilieID INTEGER,
    Datum DATE,
    DatumText VARCHAR(20),
    BisDatum DATE,
    BisDatumText VARCHAR(20),
    tblOrtID INTEGER,
    tblKonfessionID INTEGER,
    Zusatz VARCHAR(255),
    Referenz VARCHAR(255),
    FSID VARCHAR(255),
    Info MEMO,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblEreignisID));

DROP TABLE tblEreignisDokument;

CREATE TABLE tblEreignisDokument(
    tblEreignisDokumentID COUNTER,
    tblEreignisID INTEGER,
    Speicherort VARCHAR(255),
    Referenz VARCHAR(255),
    FSID VARCHAR(255),
    Info MEMO,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblEreignisDokumentID));

DROP TABLE tblEreignisPersonArt;

CREATE TABLE tblEreignisPersonArt(
    tblEreignisPersonArtID COUNTER,
    EreignisPersonArt VARCHAR(255),
    Zeichen VARCHAR(10),
    Reihenfolge INTEGER,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblEreignisPersonArtID));


DROP TABLE tblEreignisPerson;

CREATE TABLE tblEreignisPerson(
    tblEreignisPersonID COUNTER,
    tblEreignisID INTEGER,
    tblEreignisPersonArtID INTEGER,
    tblPersonID INTEGER,
    Info MEMO,
	CONSTRAINT PrimaryKey PRIMARY KEY (tblEreignisPersonID));

DROP TABLE tblZusatz;

CREATE TABLE tblZusatz(
    tblZusatzID COUNTER,
    tblEreignisArtID INTEGER,
    Zusatz VARCHAR(255),
	CONSTRAINT PrimaryKey PRIMARY KEY (tblZusatzID));

DROP TABLE tblPVorname;

DROP TABLE tblVKH;

CREATE TABLE tblVKH(
    tblVKHID COUNTER,
    BUCH_H	VARCHAR(255),
    SEITE_H	INTEGER,
    NR_H VARCHAR(10),
    HDatum DATE,
    DimDatum DATE,
    VN_BR VARCHAR(255),
    FN_BR VARCHAR(255),
    W_BR VARCHAR(255),
    H_BR VARCHAR(255),
    Z_BR VARCHAR(255),
    VN_VBR VARCHAR(255),
    FN_VBR VARCHAR(255),
    Z_VBR VARCHAR(255),
    VN_MBR VARCHAR(255),
    FN_MBR VARCHAR(255),
    Z_MBR VARCHAR(255),
    W_EBR VARCHAR(255),
    VN_BT VARCHAR(255),
    FN_BT VARCHAR(255),
    W_BT VARCHAR(255),
    H_BT VARCHAR(255),
    Z_BT VARCHAR(255),
    VN_VBT VARCHAR(255),
    FN_VBT VARCHAR(255),
    Z_VBT VARCHAR(255),
    VN_MBT VARCHAR(255),
    FN_MBT VARCHAR(255),
    Z_MBT VARCHAR(255),
    W_EBT VARCHAR(255),
    ANM_H VARCHAR(255),
    VN_HZ1 VARCHAR(255),
    FN_HZ1 VARCHAR(255),
    G_HZ1 VARCHAR(255),
    Z_HZ1 VARCHAR(255),
    VN_HZ2 VARCHAR(255),
    FN_HZ2 VARCHAR(255),
    G_HZ2 VARCHAR(255),
    Z_HZ2 VARCHAR(255),
    VN_HZ3 VARCHAR(255),
    FN_HZ3 VARCHAR(255),
    G_HZ3 VARCHAR(255),
    Z_HZ3 VARCHAR(255),
    VN_HZ4 VARCHAR(255),
    FN_HZ4 VARCHAR(255),
    G_HZ4 VARCHAR(255),
    Z_HZ4 VARCHAR(255),
    CheckNeeded YESNO DEFAULT 0,
    CONSTRAINT PrimaryKey PRIMARY KEY (tblVKHID));

DROP PROCEDURE qryPerson;

CREATE PROCEDURE qryPerson AS
    SELECT
        tblPerson.tblPersonID,
        tblPerson.PS,
        tblNachname.Nachname,
        tblPerson.Vorname
    FROM
        tblPerson
        INNER JOIN tblNachname ON tblPerson.tblNachnameID = tblNachname.tblNachnameID;

INSERT INTO tblKonfession (Konfessionkurz, Konfession) VALUES ('', '');
INSERT INTO tblKonfession (Konfessionkurz, Konfession) VALUES ('ev', 'evangelisch');
INSERT INTO tblKonfession (Konfessionkurz, Konfession) VALUES ('luth', 'lutherisch');
INSERT INTO tblKonfession (Konfessionkurz, Konfession) VALUES ('rk', 'römisch katholisch');

INSERT INTO tblKreis (Kreis) VALUES ('');

INSERT INTO tblOrt (Ort, tblKreisID) VALUES ('', 1);

INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Geburt', '*', 1, True);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Tauf', '~', 2, True);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Heirat', 'oo', 3, False);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Heirat K', 'ook', 4, False);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Scheidung', '', 5, False);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Sterbe', '+', 6, True);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Begräbnis', '', 7, True);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Verlobung', '', 8, False);
INSERT INTO tblEreignisArt (EreignisArt, Zeichen, Reihenfolge, PersonenEreignis) VALUES ('Beruf', '', 10, True);

INSERT INTO tblEreignisPersonArt (EreignisPersonArt) VALUES ('Taufpate');
INSERT INTO tblEreignisPersonArt (EreignisPersonArt) VALUES ('Trauzeuge');

UPDATE tblVersion SET Version = 4;
