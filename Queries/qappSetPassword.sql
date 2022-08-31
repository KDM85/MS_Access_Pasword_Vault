PARAMETERS [hash] Text ( 255 ), [user] Text ( 255 ), [service] Text ( 255 );
INSERT INTO tblVault ( Service, Username, [Password] )
VALUES ([service], [user], [hash]);
