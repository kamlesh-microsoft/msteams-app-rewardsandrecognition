To begin, you will need:

* An Azure subscription where you can create the following kind of resources:

* App service

* App service plan

* Bot channels registration

* Azure storage account

* Azure search

* Application Insights

* A copy of the Awardster app GitHub [repo](https://github.com/OfficeDev/microsoft-teams-apps-rewardsandrecognition)

  

### Step 1: Register Azure AD applications

Register one Azure AD applications in your tenant's directory: for the bot and tab app authentication.

  

1. Log in to the Azure Portal for your subscription, and go to the "App registrations" blade [here](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps).

  

2. Click on "New registration", and create an Azure AD application.

1.  **Name**: The name of your Teams app - if you are following the template for a default deployment, we recommend "Awardster".

2.  **Supported account types**: Select "Accounts in any organizational directory"

3. Leave the "Redirect URL" field blank.

  

[[/Images/multitenant_app_creation.png|Multi tenant selection in app registration portal]]

  

3. Click on the "Register" button.

  

4. When the app is registered, you'll be taken to the app's "Overview" page. Copy the **Application (client) ID**; we will need it later. Verify that the "Supported account types" is set to **Multiple organizations**.

  

[[/Images/multitenant_app_overview.png| App registration overview page]]

  

5. On the side rail in the Manage section, navigate to the "Certificates & secrets" section. In the Client secrets section, click on "+ New client secret". Add a description for the secret and select an expiry time. Click "Add".

  

[[Images/multitenant_app_secret.png| App registration secret overview page]]

  

6. Once the client secret is created, copy its **Value**; we will need it later.

   **Name**: The name of your tab app. We advise appending “Tab” to the name of this app; for example, “Reward and Recognition Tab”.


At this point you have 4 unique values:

 
* Application (client) ID for the bot and tab

* Client secret for the bot

* Directory (tenant) ID, which is the same for both apps


  
We recommend that you copy these values into a text file, using an application like Notepad. We will need these values later.

  

### Step 2: Deploy to your Azure subscription

  

1. Click on the "Deploy to Azure" button below.

[![Deploy to Azure](https://azuredeploy.net/deploybutton.png)](https://portal.azure.com/#create/Microsoft.Template/uri/https%3A%2F%2Fraw.githubusercontent.com%2FOfficeDev%2Fmicrosoft-teams-apps-incidentreport%2Fmaster%2FDeployment%2Fazuredeploy.json)

  

2. When prompted, log in to your Azure subscription.

  

3. Azure will create a "Custom deployment" based on the ARM template and ask you to fill in the template parameters.

  

4. Select a subscription and resource group.

* We recommend creating a new resource group.

* The resource group location MUST be in a data center that supports: Application Insights; and Azure Search. For an up-to-date list, click [here](https://azure.microsoft.com/en-us/global-infrastructure/services/?products=logic-apps,cognitive-services,search,monitor), and select a region where the following services are available:

* Application Insights

* Azure Search

  

5. Enter a "Base Resource Name", which the template uses to generate names for the other resources.

* The app service names `[Base Resource Name]`, must be available. For example, if you select `contosoawardster` as the base name, the names `contosoawardster` must be available (not taken); otherwise, the deployment will fail with a Conflict error.

* Remember the base resource name that you selected. We will need it later.

  

6. Fill in the various IDs in the template:

  

1.  **Bot Client ID**: The application (client) ID of the Microsoft Teams Bot app

  

2.  **Bot Client Secret**: The client secret of the Microsoft Teams Bot app

  

3.  **Tenant Id**: The tenant ID above

  
Make sure that the values are copied as-is, with no extra spaces. The template checks that GUIDs are exactly 36 characters.

  

7. If you wish to change the app name, description, and icon from the defaults, modify the corresponding template parameters.

  

8. Add the team link in team link input field. Get the link to the team with your experts from the Teams client. To do so, open Microsoft Teams, and navigate to the team. Click on the "..." next to the team name, then select "Get link to team".

  

9. Agree to the Azure terms and conditions by clicking on the check box "I agree to the terms and conditions stated above" located at the bottom of the page.

  

10. Click on "Purchase" to start the deployment.

  

11. Wait for the deployment to finish. You can check the progress of the deployment from the "Notifications" pane of the Azure Portal. It can take more than 10 minutes for the deployment to finish.

  

12. Once the deployment has finished, you would be directed to a page that has the following fields:

* botId - This is the Microsoft Application ID for the Incident Reporter bot.

* appDomain - This is the base domain for the Incident Reporter Bot.

* configurationAppUrl - This is the URL for the configuration web application.

### Step 3: Set up authentication for the app

1. Go back to the "App Registrations" page [here]([https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps](https://portal.azure.com/#blade/Microsoft_AAD_IAM/ActiveDirectoryMenuBlade/RegisteredApps)) 

2.  On the overview page of app created in Step 1, copy and save the  **Bot(client) ID**. You’ll need it later when updating your Teams application manifest.

3.  Select  **Expose an API**  under  **Manage**. Select the  **Set**  link to generate the Application ID URI in the form of  `api://{BotID}`. Insert your fully qualified domain name (with a forward slash "/" appended to the end) between the double forward slashes and the GUID. The entire ID should have the form of:  `api://fully-qualified-domain-name.com/{BotID}`
    -   ex:  `api://subdomain.example.com:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7`.


5.  Select the  **Add a scope**  button. In the panel that opens, enter  `access_as_user`  as the  **Scope name**.

6.  Set Who can consent? to Admins and users

7.  Fill in the fields for configuring the admin and user consent prompts with values that are appropriate for the  `access_as_user`  scope. Suggestions:
    -   **Admin consent title:**  Teams can access the user’s profile
    -   **Admin consent description**: Allows Teams to call the app’s web APIs as the current user.
    -   **User consent title**: Teams can access your user profile and make requests on your behalf
    -   **User consent description:**  Enable Teams to call this app’s APIs with the same rights that you have
8.  Ensure that  **State**  is set to  **Enabled**
9.  Select  **Add scope**
    -   Note: The domain part of the  **Scope name**  displayed just below the text field should automatically match the  **Application ID**  URI set in the previous step, with  `/access_as_user`  appended to the end; for example:
        -   `api://subdomain.example.com:6789/c6c1f32b-5e55-4997-881a-753cc1d563b7/access_as_user`

10.  In the  **Authorized client applications**  section, you identify the applications that you want to authorize to your app’s web application. Each of the following IDs needs to be entered:
   -   1fec8e78-bce4-4aaf-ab1b-5451cc387264`  (Teams mobile/desktop application)
   -   5e3ce6c0-2b1f-4285-8d4b-75ee78787346`  (Teams web application)

11.  Navigate to  **API Permissions**, and make sure to add the follow permissions:
   -   User.Read (enabled by default)
   -   email
   -   offline_access
   -   openid
   -   profile

**Note:** The detailed guidelines for registering an application for SSO Microsoft Teams tab can be found [here]([https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso))



### Step 4: Create the Teams app packages

  

Create two Teams app packages: one for end-users to install personally, and one to be installed to the experts team.

  

1. Open the `Manifest\manifest.json` file in a text editor.

  

2. Change the placeholder fields in the manifest to values appropriate for your organization.

  

*  `developer.name` ([What's this?](https://docs.microsoft.com/en-us/microsoftteams/platform/resources/schema/manifest-schema#developer))

  

*  `developer.websiteUrl`

  

*  `developer.privacyUrl`

  

*  `developer.termsOfUseUrl`

  

3. Change the `<<botId>>` placeholder to your Azure AD application's ID from above. This is the same GUID that you entered in the template under "Bot Client ID".

4. In the "validDomains" section, replace the `<<appDomain>>` with your Bot App Service's domain. This will be `[BaseResourceName].fdwebsites.net`. For example if you chose "contosoincidentreport" as the base name, change the placeholder to `contosorewardandrecognition.fdwebsites.net`.

5. In the "webApplicationInfo" section, replace the   `<<application_GUID>>` with client ID of the tab app created in Step 3(2). Also replace `<<web_api resource>>` with following Application ID URI set in Step 3(3). This will be as follows `api://[BaseResourceName].fdwebsites.net/<<application_GUID>>`. 

7. Create a ZIP package with the `manifest.json`,`color.png`, and `outline.png`. The two image files are the icons for your app in Teams.

* Make sure that the 3 files are the _top level_ of the ZIP package, with no nested folders.

  

[[/Images/file-explorer.png|Manifest file in windows explorer]]

  

### Step 5: Run the apps in Microsoft Teams

1. If your tenant has side loading apps enabled, you can install your app by following the instructions
[here](https://docs.microsoft.com/en-us/microsoftteams/platform/concepts/apps/apps-upload#load-your-package-into-teams)

  2. You can also upload it to your tenant's app catalog, so that it can be available for everyone in your tenant to install. See [here](https://docs.microsoft.com/en-us/microsoftteams/tenant-apps-catalog-teams)

* We recommend using [app permission policies](https://docs.microsoft.com/en-us/microsoftteams/teams-app-permission-policies) to restrict access to this app to the members of the experts team.

  
 3. Install the app (the `awardster.zip` package) to your team.

  

### Troubleshooting

  

Please see our [Troubleshooting](Troubleshooting) page.