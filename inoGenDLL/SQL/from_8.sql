DROP TABLE tblQuelle;

CREATE TABLE tblQuelle(
    tblQuelleID COUNTER,
    Quelle VARCHAR(255),
    QuelleKurz VARCHAR(255),
    QuelleBeschreibung MEMO,
    active YESNO DEFAULT 1,
    CONSTRAINT PrimaryKey PRIMARY KEY (tblQuelleID));

DROP TABLE tblQuellZitat;

CREATE TABLE tblQuellZitat(
    tblQuellZitatID COUNTER,
    tblQuelleID INTEGER,
    tblEreignisArtID INTEGER,
    Jahr INTEGER,
    Seite VARCHAR(20),
    Bd VARCHAR(20),
    Nummer VARCHAR(20),
    Datum DATE,
    InternetAdresse VARCHAR(255),
    URLBeschreibung MEMO,
    ZitatBeschreibung MEMO,
    active YESNO DEFAULT 1,
    CONSTRAINT PrimaryKey PRIMARY KEY (tblQuellZitatID));

DROP TABLE tblEventTag;

CREATE TABLE tblEventTag(
    Tag VARCHAR(255),
    TagD VARCHAR(255),
    CONSTRAINT PrimaryKey PRIMARY KEY (Tag));

DROP TABLE tblEreignisZitat;

CREATE TABLE tblEreignisZitat(
    tblEreignisZitatID COUNTER,
    tblQuellZitatID INTEGER,
    tblEreignisID INTEGER,
    tblPersonID INTEGER,
    EventTag VARCHAR(255),
    active YESNO DEFAULT 1,
    CONSTRAINT PrimaryKey PRIMARY KEY (tblEreignisZitatID));



INSERT INTO tblEventTag (Tag, TagD) VALUES ('OFFICIATOR', 'Amtsperson');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('WIFE', 'Ehefrau');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('SPOU', 'Ehegatte');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('HUSB', 'Ehemann');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('PARENT', 'Elternteil');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('FRIEND', 'Freund');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('CHIL', 'Kind');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('MULTIPLE', 'Mehrling');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('MOTH', 'Mutter');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('NGHBR', 'Nachbar');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('GODP', 'Pate');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('OTHER', 'Sonstige');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('FATH', 'Vater');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('WITN', 'Zeuge');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('CLERGY', 'religiöser Amtsträger');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_PROB', 'Proband');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_ADOP_FATH', 'Adoptivvater');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_ADOP_MOTH', 'Adoptivmutter');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_HUSB_FATH', 'Vater des Ehemannes');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_HUSB_MOTH', 'Mutter des Ehemannes');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_HUSB_PARENT', 'Eltern des Ehemannes');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_WIFE_FATH', 'Vater der Ehefrau');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_WIFE_MOTH', 'Mutter der Ehefrau');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_WIFE_PARENT', 'Eltern der Ehefrau');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_FATH_FATH', 'Vater des Vaters');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_FATH_MOTH', 'Mutter des Vaters');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_MOTH_FATH', 'Vater der Mutter');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_MOTH_MOTH', 'Mutter der Mutter');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_BRIDEGROOM', 'Bräutigam');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_BRIDE', 'Braut');
INSERT INTO tblEventTag (Tag, TagD) VALUES ('_TWIN', 'Zwilling');

UPDATE tblVersion SET Version = 9;
