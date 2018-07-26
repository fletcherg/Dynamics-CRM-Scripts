/*
	Clean up Outlook Sync tables

	Run against the organization's CRM database
	
	Bla bla bla you shouldn't run this, not supported
	by Microsoft. Directly modifying the database is
	a bad idea. There is nothing here but pain and
	suffering.
*/

DECLARE @SubscriptionID as UniqueIdentifier
DECLARE @SystemUserID as UniqueIdentifier
DECLARE @TableToCheck as varchar(50)
DECLARE @MachineName as nvarchar(200)
DECLARE @InternalEmailAddress as nvarchar(100)
DECLARE @SubscriptionCursor as CURSOR

/*
SELECT SubscriptionId, users.SystemuserID, FullName,
			InternalEmailAddress, DomainName, MachineName,
			SyncEntryTableName, Subscriptiontype,
			CompletedSyncStartedOn
	FROM [dbo].[Subscription] sub
	JOIN [dbo].[SystemUserbase] users
	ON sub.SystemuserID = users.SystemUserId
*/

/* Open Cursor across the subscrpitions */
SET @SubscriptionCursor = CURSOR FOR SELECT sub.SubscriptionId,
			sub.SystemuserID, sub.MachineName,
			users.InternalEmailAddress
	FROM [dbo].[Subscription] sub
	JOIN [dbo].[SystemUserbase] users
	ON sub.SystemuserID = users.SystemUserId

OPEN @SubscriptionCursor

FETCH NEXT FROM @SubscriptionCursor INTO @SubscriptionId,
			@SystemuserID, @MachineName, @InternalEmailAddress

WHILE @@FETCH_STATUS = 0
BEGIN
 /*PRINT 'Processing: ' + cast(@SubscriptionID AS VARCHAR (50)) */
 SET @TableToCheck = 'SyncEntry_' + LOWER(REPLACE(CAST(@SubscriptionID AS VARCHAR (50)),'-',''))
 PRINT 'Checking for: ' + @TableToCheck

	/* See if there's a SyncEntry_{ID} table */
 
	IF (EXISTS (SELECT * 
                 FROM INFORMATION_SCHEMA.TABLES 
                 WHERE TABLE_SCHEMA = 'dbo' 
                 AND  TABLE_NAME = @TableToCheck))
	BEGIN
	    PRINT CHAR(9) + 'Seems OK! - ' + CAST(@SubscriptionID AS VARCHAR (50))
	END

	ELSE
	BEGIN
	
		/* If not, clean out the subscription data */
		
		PRINT CHAR(9) + 'FOUND ISSUE! - ' + CAST(@SubscriptionID AS VARCHAR (50))
		PRINT CHAR(9) + 'Cleaning up sync data... (Machine:' + @MachineName + ', User: ' + @InternalEmailAddress + ')'
		/*
		DELETE FROM [dbo].[SubscriptionClients] WHERE SubscriptionID = @SubscriptionID
		DELETE FROM [dbo].[SubscriptionManuallyTrackedObject] WHERE SubscriptionID = @SubscriptionID
		DELETE FROM [dbo].[SubscriptionStatisticsOutlookBase] WHERE SubscriptionID = @SubscriptionID
		DELETE FROM [dbo].[SubscriptionSyncEntryOutlookBase] WHERE SubscriptionID = @SubscriptionID
		DELETE FROM [dbo].[SubscriptionSyncInfo] WHERE SubscriptionID = @SubscriptionID
		DELETE FROM [dbo].[Subscription] WHERE SubscriptionID = @SubscriptionID
		*/
	END
 FETCH NEXT FROM @SubscriptionCursor INTO @SubscriptionId,
		@SystemuserID, @MachineName, @InternalEmailAddress
END

CLOSE @SubscriptionCursor
DEALLOCATE @SubscriptionCursor