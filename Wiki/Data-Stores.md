The app uses the following data stores:

  

1. Azure Storage Account  

* The [TeamConfiguration]  table stores details of team’s information where bot is installed.
* The [AdminDetail] table stores details of all the R&R Admin details.
* The [AwardDetail]  table stores various awards created by R&R admin.
* The [RewardCycleDetail]  table stores various rewards created by R&R admin.
* The [NominationDetail]  table stores reward nomination, nominated by team members.
* The [EndorseDetail]  table stores Endorse nomination, endorsed by team members.

  

2. Azure Search service list item index

* Search service to query data related to nominations from Azure storage.

  

All these resources are created in your Azure subscription. None are hosted directly by Microsoft.

  

## Storage account

### TeamConfiguration table

|Attribute | Comment|
|-----------|---------
|PartitionKey|This represents the partition key of the azure storage table [TeamId]|
|RowKey|Represents the unique id of each row|
|TeamId|Represents team id|
|BotInstalledOn|The date time when bot was installed|
|ServiceUrl|Bot activity service url|
|ServiceUrl|String	Represents the service url.|


  


### AdminDetail table
 
|Attribute | Comment|
|-----------|---------
|Partition Key|This represents the partition key of the azure storage table [TeamId]|
|Row Key|Represents the unique id of each row|
|AdminName|The name of the newly created Admin|
|TeamId|String	Represent team id where bot installed.|
|AdminPrincipalName|he email address of the newly created Admin|
|AdminObjectId|The AAD Object Id of newly created Admin|
|NoteForTeam|Represent remarks about selecting new Admin|
|CreatedOn|Contains the date time of Admin creation|
|CreatedByUserPrincipalName|The email address of the end user who created Admin|
|CreatedByObjectId|The AAD Object Id of user who created Admin

  

### AwardDetail table

|Attribute | Comment|
|-----------|---------
|Partition Key|This represents the partition key of the azure storage table [TeamId]|
|Row Key|Represents the unique id of each row|
|TeamId|String	Represent team id where bot installed.|
|AwardName|Represent award unique name|
|AwardDescription|Represent description about new award|
|AwardLink|Represent image URL of award|
|CreatedBy|The AAD Object Id of Admin who created the award|
|CreatedOn|Contains the date time of award creation
|ModifiedBy|The AAD Object Id of Admin who modified the award.|
|AwardId|Unique id of each award|

  

### RewardCycleDetail table

| Attribute |  Comment |
| ----------- | ---------
|Partition Key| This represents the partition key of the azure storage table [TeamId]|
|RowKey|Represents the unique id of each cycle|
|RewardCycleStartDate|Represents start date of reward cycle|
|RewardCycleEndDate|Represents end date of reward cycle|
|NumberOfOccurrences|Represents number of occurrences of each reward cycle|
|IsRecurring|Represents the state of recurring. Integer value. 0 = No / 1 = Yes|
|RangeOfOccurrence|Represents the state of occurrence. Integer value. 0 = NoEndDate / 1 = End On Date / 2= End After Number of Occurrences|
|RangeOfOccurrenceEndDate|Represents end date of occurrence|
|TeamId|String	Represent team id where bot installed.|
|RewardCycleState|Represents current state of reward cycle. Integer value. 0 = Inactive / 1 =Active|
|CreatedByObjectId|The AAD Object Id of Admin who created the award cycle.|
|CreatedOn|Represents the date time of award cycle creation.|
|ResultPublished|Represents the state of reward publish for current reward cycle. 0  = False / 1 = true|



### NominationDetail table

| Attribute | Comment |
|------------|---------|
|PartitionKey|This represents the partition key of the azure storage table [TeamId]|
|RowKey|Represents the unique id of each|
|AwardName|Represents the award name|
|AwardId|Represents the unique id of award - RowKey [Award - table]|
|NominatedToName|Represents name of the user to whom nomination is sent|
|NominatedToPrincipalName|Represents mail address user to whom nomination is sent|
|NominatedToObjectId|Represents AAD Object Id of the user to whom nomination is sent|
|NominatedOn|Represents the date time of award nomination|
|ReasonForNomination|Represents reason/Justification for award nomination|
|RewardCycleId|Represents the unique id of reward cycle - RowKey [RewardCycle - table]|
|NominatedByName|The name of the team member who nominated the award|
|NominatedByPrincipalName|The email address of the team member who nominated the award|
|NominatedByObjectId|The AAD Object Id of the team member who nominated the award|
|IsGroupNomination|Represents the state of group nomination. 0 = Yes / 1 = No|
|GroupName|Represents the unique id of each group nomination|
|AwardGranted|Represents the state if reward is granted by R&R admin. 0 = false / 1 = true|
|AwardPublishedOn|The datetime of award published|



### EndorseDetail table

| Attribute | Comment |
|------------|---------|
|PartitionKey| This represents the partition key of the azure storage table [TeamId]|
|RowKey| Represents the unique id of each row|
|EndorseForAward| Represents the award name for which user is endorsed|
|EndorseForAwardId| Represents award Id|
|AwardCycle| Represents the award cycle when user in endorsed|
|EndorsedToPrincipalName| Represents the email address of the nominee|
|EndorsedToObjectId|Represents AAD Object Id of the nominee|
|EndorsedByPrincipalName|Represents the email address of the person endorsed by|
|EndorsedByObjectId| Represents AAD Object Id of the person endorsed by|
|EndorsedOn|Represents the date time of endorsement|