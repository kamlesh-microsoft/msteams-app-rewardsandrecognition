### Generic possible issues

There are certain issues that can arise that are common to many of the app templates. Please check [here](https://github.com/OfficeDev/microsoft-teams-stickers-app/wiki/Troubleshooting) for reference to these.
  
### **Problems deploying to Azure**

### **1. Error when attempting to reuse a Microsoft Azure AD application ID for the bot registration**

#### Description

Bot is not valid.

```
Errors: MsaAppId is already in use.
```

- Creating the resource of type Microsoft.BotService/botServices failed with status "BadRequest"

This happens when the Microsoft Azure application ID entered during the setup of the deployment has already been used and registered for a bot.

#### Fix

Either register a new Microsoft Azure AD application or delete the bot registration that is currently using the attempted Microsoft Azure application ID.

### **2. Error while deploying the ARM Template**

#### Description

This happens when the resources are already created or due to some conflicts.
```

Errors: The resource operation completed with terminal provisioning state 'Failed'

```
#### Fix

In case of such a scenario, user needs to navigate to deployment center section of failed/conflict resources through the azure portal and check the error logs to get the actual errors and can fix it accordingly.

Redeploy it after fixing the issue/conflict.

**Didn't find your problem here?**

Please, report the issue [here](https://github.com/OfficeDev/microsoft-teams-apps-incidentreport/issues/new)