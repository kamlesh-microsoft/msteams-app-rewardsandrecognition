[[/Images/architecture_overview.png|Rewards and Recognition architecture diagram]]

  

The **Rewards and Recognition** application has the following main components:

  

## Rewards and Recognition Bot

The bot is built using the [Bot Framework SDK v4 for .NET](https://docs.microsoft.com/en-us/azure/bot-service/bot-service-overview-introduction?view=azure-bot-service-4.0) and [ASP.NET Core 2.1](https://docs.microsoft.com/en-us/aspnet/core/?view=aspnetcore-2.1). The bot has a conversational interface in team scope for users. 

Rewards and Recognition Bot provides all end users (internal users seeking help from a central team) and easy interface (bot) to:

* Configure the R&R admin from list of members in team.

* Nominate team members once the reward cycle is active.

* Endorse other team members for the rewards. Endorsement can be done only once for given nomination by each user.

## Rewards and Recognition Tab

Manage Rewards tab will show nominations for all rewards in current cycle. Only R&R admin will have access to this tab. This tab will be used to  publish results. If the user is not an admin, he/she can see the winners of the previously published reward cycle.

Tab will have three  action buttons:

-   Publish Results: Publish results will reward selected nominations and publish winners in team channel. A confirmation will appear on click of Publish Results button. Confirmation message will show count of winners for each category.
    
-   Manage Reward Cycle: Manage new cycle will open a task module. Task module will have two options:  *Manage Awards* & *Set Reward Cycle*

- Configure R&R Admin: Existing Admin can change the R&R admin, in case he/she chooses to leave the team or any other reasons.
  

## Rewards and Recognition Messaging Extension

A messaging extension with [query commands](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/messaging-extensions/search-extensions), which the team can use to search for nominations of current reward cycle in the team. It also implements messaging action that user can use to nominate other team members for the reward.