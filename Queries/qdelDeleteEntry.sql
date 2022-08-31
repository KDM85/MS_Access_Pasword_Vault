PARAMETERS [service] Text ( 255 );
DELETE *
FROM tblVault
WHERE Service = [service];
