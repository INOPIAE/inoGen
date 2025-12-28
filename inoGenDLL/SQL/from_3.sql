ALTER TABLE tblVKH ADD COLUMN CheckNeeded YESNO;

UPDATE tblVKH SET CheckNeeded = 0 WHERE CheckNeeded IS NULL;

UPDATE tblVersion SET Version = 4;