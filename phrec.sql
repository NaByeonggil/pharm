SET NOCOUNT ON;
ALTER LOGIN [DESKTOP-8VR1BS2\n1] ENABLE;
IF NOT EXISTS (SELECT 1 FROM sys.server_role_members rm
  JOIN sys.server_principals r ON r.principal_id=rm.role_principal_id AND r.name='sysadmin'
  JOIN sys.server_principals m ON m.principal_id=rm.member_principal_id
  WHERE m.name='DESKTOP-8VR1BS2\n1')
  ALTER SERVER ROLE sysadmin ADD MEMBER [DESKTOP-8VR1BS2\n1];
SELECT SUSER_NAME() AS whoami, IS_SRVROLEMEMBER('sysadmin') AS sysadmin_now;
PRINT '== n1 ENABLED + sysadmin ==';
