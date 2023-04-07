## Integrating ChatGPT with Microsoft Teams
- Wouldn't it be great if you and your teammates could easily send questions to ChatGPT in Microsoft Teams group chats or meetings? This guide will show you how to integrate Microsoft Teams with the Azure OpenAI platform in just a few steps. By taking advantage of a Microsoft 365 developer platform membership and an Azure free subscription, you can create a sandbox environment specifically for this proof of concept.

![intro](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-overview-small.png)

## What is TeamsGPT?
- Teamsgpt is a custom Teams app that is integrated with Azure and OpenAI's ChatGPT model. It can be added into a chat/group/channel as an AI assistant to help answer questions and provide assistance. It leverages Azure Bot Service, App Service, and Cognitive Service OpenAI to provide its functionality.
- This diagram shows the fewest cloud resources necessary to create a POC/POV development environment. A different architecture will be discussed later for production workloads.

![devarch](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/devarch.png)
## Prerequisites for Custom Teams App development
1. Microsoft 365 Admin access (Production) or Join the Microsoft 365 Developer Program (Dev)
2. Azure subscription owner access (Production) or Enroll a Azure free subscription (Dev)
3. Azure OpenAI Registration for your Azure subscription
4. Visual Studio / Visual Studio Code
5. Teams Toolkit
6. C# or Javascript programming capabilities
[More prerequisites in details](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/tools-prerequisites)

## Microsoft 365 Developer Program
- You have to have Microsoft Teams Admin permission for publishing your custom app to Teams. If you don't, try to apply a M365 sandbox by joining a membership from [Microsoft 365 Developer Program](https://developer.microsoft.com/en-us/microsoft-365/dev-program)
- You will be allocated a free M365 sandbox tenant with an E5 subscription for 90 days, which can be renewed. If the email you registered with has a Visual Studio Pro or Enterprise subscription, there will be no expiration date.
- After the sandbox tenant created, click [go to myDashboard](https://developer.microsoft.com/en-us/microsoft-365/profile), you will find your profile information there. Click [go to subscription](https://www.office.com) will redirect you to office website.
- Click the ADMIN icon to access [M365 admin portal](https://admin.microsoft.com/Adminportal/Home?source=applauncher#/homepage).
![m365admin](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-join365devpng1.png)
- Sooner we will come back here to approve your pending custom app by the left blade menu, Teams Apps > Manage apps 
  
## Azure Free Subscription
- Get started with 12 months of free services, 40+ services that are always free, and USD200 in credit. Create your free account today with Microsoft Azure. [Start free](https://azure.microsoft.com/en-us/free/)
- If you want to run this tutorial smoothly, avoid spend a lot of time on troubleshooting, [free account is recommended](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/tools-prerequisites#azure-account)

## Azure OpenAI (AOAI) Registration
- To access AOAI's ChatGPT model endpoint and deploy a AOAI resource for your custom app, your Azure subscription must be approved.
- Azure OpenAI requires registration and is currently only available to approved **enterprise customers and partners**. Customers who wish to use Azure OpenAI are required to submit a [registration form](https://aka.ms/oai/access)
- [Azure OpenAI Registration process](https://learn.microsoft.com/en-us/legal/cognitive-services/openai/limited-access#registration-process)

*Hints: You can accelerate the application process by using the Azure subscription under enterprise agreement. The App Service deployed with your free Azure subscription can send AOAI API to the granted AOAI resources that deployed with your enterprise Azure subscription.*

## TeamsGPT Components
![components](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-components.png)

## Install Teams Toolkit
- You can either use your favorite Dev IDE like Visual Studio Code or Visual Studio. Both of them support Teams Toolkit. If you prefer Javascipt programming, go for VSC whereas C# for VS. Check it out over here for corresponding installation guides
  - [Visual Studio Code / Typescript](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode&pivots=visual-studio-code#install-teams-toolkit-for-visual-studio)
  - [Visual Studio / C#](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/install-teams-toolkit?tabs=vscode&pivots=visual-studio#install-teams-toolkit-for-visual-studio)
- this repo\bot-sso-v5 uses Teams Toolkit PreRelease Version 4.99.2 (aka v5.0)
- If you installed Toolkit Release version 4.2.4, please use another folder repo\bot-sso and refer to another [README.md](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/README.md) instead.

*Hints: Javascript is preferred as AOAI API library only available with Python (Preview) and with Node.js ([from Github community](https://github.com/1openwindow/azure-openai-node)) at this moment (~March 2023)*

## Create a new Teams Project and start debugging in local environment
1. In VSC, click the Teams Toolkit icon on the left blade menu, select "Create a new app" > "Start from a sample"
2. Pick **"Bot App with SSO enabled"** as your template project and assign a new folder path
![ssosample](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-dev-createnewapp.png)
3. Sign in both M365 and Azure accounts
![singinkit](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/signin_teamstoolkit.png)
- If no Azure subscription has been found after signing succeed, try to edit teamsapp.yml, fill in your *AZURE_SUBSCRIPTION_ID* and targeted *AZURE_RESOURCE_GROUP_NAME* and then try to sign in and sign off back and fore for a few times, or restart VSC.
##### {projectfolder}\bot-sso-v5\teamsapp.yml
```javascript
    {
        with:
            subscriptionId: ${{AZURE_SUBSCRIPTION_ID}} # The AZURE_SUBSCRIPTION_ID is a built-in environment variable. TeamsFx will ask you select one subscription if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select subscription if it's empty in this case.
            resourceGroupName: ${{AZURE_RESOURCE_GROUP_NAME}} # The AZURE_RESOURCE_GROUP_NAME is a built-in environment variable. TeamsFx will ask you to select or create one resource group if its value is empty. You're free to reference other environment varialbe here, but TeamsFx will not ask you to select or create resource grouop if it's empty in this case.
    }
```
4. After the project created, press F5 to debug start this custom app without any code changes.
5. The debug pipeline will invoke your browser (Edge/Chrome), and you need to sign in with your sandbox user account like "xxxx@yyyyyyy.onmicrosoft.com". 
6. A modal screen will be prompted to ask you to install your custom app like "xxxxxxxxxx-local-debug"
![debugmode](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-dev-01.png)
7. In the teams chat, you can try to send "welcome", "learn" to bot. This will invoke the pre-built logic of AdaptiveCards.
- Check it out over here for more details, [Designing Adaptive Cards for your Microsoft Teams app](https://learn.microsoft.com/en-us/microsoftteams/platform/task-modules-and-cards/cards/design-effective-cards?tabs=design)
![acard](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/adaptivecards.jpg)
8. Right now, if everything on the right track, please stop the debugger to avoid any disruption made during app deployment.

## Understanding Teams Toolkit local environment debugging lifecycle
- For understanding more about the debugging lifecycle, please refer to [Teams Toolkit Visual Studio Code v5.0 Prerelease Guide](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5.0-Prerelease-Guide)

![tk5](https://user-images.githubusercontent.com/103554011/218403711-68502280-228d-43e5-afa1-e6d0c01712e5.png)

## Deploy to Azure Cloud
- Provisioning and deploying your custom app to Azure is very easy things. Teams toolkit will do all the complicated things for you, what you need to do 

![tkmenu](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-teamstoolkit.png)
1. Click the Teams toolkit menu> Deployment > Provision in the cloud, and then be patient. A prompt will be pop-up asking for Azure Region and ResourceGroup name. Please stick with [AOAI supported regions](https://learn.microsoft.com/en-us/azure/cognitive-services/openai/overview#features-overview) like EastUS.
2. Click the Teams toolkit menu> Deployment > Deploy to the cloud, and then be patient. 
3. Click the Teams toolkit menu> Deployment > Publish to Teams, a prompt will be pop-up asking you either "send your app to your Teams administrator." or "send it to your Teams administrator manually". We will pick "Install for your organization" for convenience.
![askpublish](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-askpublish.png)

## Approving the pending custom app
1. Go to [Teams admin center](https://admin.teams.microsoft.com/policies/manage-apps)
2. Left blade menu > Teams apps > Manage apps
3. In the grid listing, put your app name like "SSOBotSample" in the "Search by name" textbox.
4. Click into your app master page, approve the pending app request.
![publish1st](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-publish.png)

## Testing: You and TeamsGPT bot
- Try to interact with your added custom app TeamsGPT bot.
1. Teams > Apps > Built for your org > click your custom app.
![addapp1](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-builtforyourorg.png)
2. In the App master page, click the "Add".
![addapp2](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-addapptome.png)
3. You also can add your custom app in the left blade menu.
![addapp3](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-selectcustomapp.png)
4. Any message in this 1:1 chat window will be delivered to your bot without @mention, try to send a welcome.

*Hints: Add apps" is not available for Guest user role, [check it out](https://support.microsoft.com/en-us/office/team-owner-member-and-guest-capabilities-in-teams-d03fdf5b-1a6e-48e4-8e07-b13e1350ec7b)*

## Testing: Tester, you and TeamsGPT bot
- We will setup a TeamsGPT bot in the conversation chat window with others.
1. Go to [M365 admin center](https://admin.microsoft.com/Adminportal/Home?#/users) > Users > Active Users > 
2. Pick one predefined user as tester, reset the password, spin off another Teams (browser), sign in with this tester account.
3. Make a chat by saying hello with your tester to initiate a new chat window.
4. Teams > Apps > Built for your org > click your custom app
![addtochat2](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-addtowhichchat.png)
5. In the App master page, click the down arrow to expose more options. Try "Add to a chat", select the chat with your tester
6. Use mention @{yourcustomappname} to chat with your bot
![addtochat3](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-dev-addbotinchat.png)


## Review provisioned Azure resources
- Teams toolkit help us to CI/CD our custom app to Azure cloud. Let's check what resources have been deployed.
1. Sign in [Azure Portal](https://portal.azure.com/) with your free account.
2. Locate the resource group you have named.
![azureres](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/v5resourcegroup.png)
3. 3 resources have been provisioned to support your custom app workload.
   - Azure Bot service: This is the underlying infrastructure that used to support the Bot framework integrated with your custom app.
   - App Service plan: This is the underlying infrastructure that used to support your App Service.
   - App Service: This is logical app instance to support your custom app logic.
4. Later, we will go back to App Service to add some application parameters for AOAI.

## Deploy AOAI ChatGPT resources in Azure Cognitive Services
1. Please follow this [guideline](https://learn.microsoft.com/en-us/azure/cognitive-services/openai/how-to/create-resource?pivots=web-portal) to deploy a GPT-35-Turbo resource. Keep the deployment region same with the deployed App Service can save the cost on cross-region traffic.
2. After the instance created, follow this [guideline](https://learn.microsoft.com/en-us/azure/cognitive-services/openai/chatgpt-quickstart?tabs=command-line&pivots=programming-language-python#retrieve-key-and-endpoint) to copy
   - KEY
   - Endpoint
   - the Model name you typed (not the model name displayed for selection)
3. Try out the [chat playground](https://oai.azure.com/portal/chat) to get the concept of how to interact with the ChatGPT model.

## Modify your custom app to send AOAI API to ChatGPT resources
- First of all, Python is the one and only one supported AOAI library with PREVIEW agreement at this moment (~March 2023). As your custom app is bootstrapped by Teams toolkit project template that is written in Javascript. You can leverage the node.js version of AOAI API library that forked by community over here [azure-openai-node](https://github.com/1openwindow/azure-openai-node)
1. Install the library in this path *{projectfolder}\bot-sso\bot*
```javascript
npm install azure-openai -save
```
2. Modify your custom app version is very important because the new version of custom app take times (~ few hours) to be reflected to Teams user in Teams App publishing workflow. An incremental versioning practice is recommended to make things clear. Try increment the "version":"1.0.0" to 1.0.1,
##### {projectfolder}\bot-sso-v5\appPackage\manifest.json
```javascript
    "version": "1.0.1",
```
*Hints: If you want to change the custom app name*
```javascript
    "name": {
        "short": "sso-bot-${{TEAMSFX_ENV}}",
        "full": "full name for sso-bot"
    },
    "description": {
        "short": "short description for sso-bot",
        "full": "full description for sso-bot"
    },
```
3. Replace your teamsBot.ts by this [teamsBot.ts](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/bot-sso/bot/teamsBot.ts) in this repo:main. This revised teamsBot.ts added a async API call to AOAI resource within OnMessage handler.
##### {projectfolder}\bot-sso-v5\bot\teamsBot.ts 
```javascript
import { Configuration, OpenAIApi, ChatCompletionRequestMessageRoleEnum} from "azure-openai"; 
```
```javascript

    this.openAiApi = new OpenAIApi(
      new Configuration({
         apiKey: process.env.AOAI_APIKEY,
         // add azure info into configuration
         azure: {
            apiKey: process.env.AOAI_APIKEY,
            endpoint: process.env.AOAI_ENDPOINT,
            // deploymentName is optional, if you donot set it, you need to set it in the request parameter
            deploymentName: process.env.AOAI_MODEL,
         }
      }),
    );

  
    this.onMessage(async (context, next) => {
      console.log("Running with Message Activity.");

      let txt = context.activity.text;
      // remove the mention of this bot
      const removedMentionText = TurnContext.removeRecipientMention(
        context.activity
      );
      if (removedMentionText) {
        // Remove the line break
        txt = removedMentionText.toLowerCase().replace(/\n|\r/g, "").trim();
      }

      let revisedprompt = [{role:ChatCompletionRequestMessageRoleEnum.System,content:process.env.CHATGPT_SYSTEMCONTENT},{role:ChatCompletionRequestMessageRoleEnum.User, content:txt}];
      console.log("createChatCompletion request: " + JSON.stringify(revisedprompt));
      try {
        const completion = await this.openAiApi.createChatCompletion({
          model: process.env.AOAI_MODEL,
          messages: revisedprompt,
          temperature: parseInt(process.env.CHATFPT_TEMPERATURE),
          max_tokens:parseInt(process.env.CHATGPT_MAXTOKEN),
          top_p: parseInt(process.env.CHATGPT_TOPP),
          stop: process.env.CHATGPT_STOPSEQ
        });
        console.log("createChatCompletion response: " + completion.data.choices[0].message.content);
        await context.sendActivity(completion.data.choices[0].message.content);
      } catch (error) {
        if (error.response) {
          console.log(error.response.status);
          console.log(error.response.data);
        } else {
          console.log(error.message);
        }
      }
         
```
4. Add the following environmental parameters in 
##### {projectfolder}\bot-sso-v5\.localSettings
```javascript
AOAI_APIKEY={your AOAI KEY}
AOAI_ENDPOINT={your AOAI endpoint}
AOAI_MODEL={your AOAI deployed model name}
CHATGPT_SYSTEMCONTENT="You are an AI assistant that helps people find information."
CHATGPT_MAXTOKEN=1000
CHATGPT_TOPP=0.90
CHATGPT_STOPSEQ=
CHATFPT_TEMPERATURE=0.7
```
5. Save all, press F5 to start debugging your custom app. Right now, you will discover your bot became smarter to respond more than just *'learn', 'welcome' and 'show'*.
6. Feel free to alter any *CHATGPT_XXXXXXX* environmental parameters to try out the differences of how the model reacts.

## Redeploy the updated custom app to cloud
- Using Teams Toolkit, deploying your updated custom app to the cloud is easy. Alternatively, you may choose to containerize the app and deploy it with other PaaS container services such as Azure Kubernetes Service, Azure Container Instance, and Azure Container App, or with DevOps services like Azure DevOps.
1. Revise the version number in *manifest.template.json*, save it.
2. Add environmental parameters mentioned above in App Service > Configuration > Application settings > New application setting, save it.
![appparameter](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/appserviceparam.png)
3. VSC: click the Teams toolkit menu> Deployment > Deploy to the cloud, and then be patient. 
4. VSC: click the Teams toolkit menu> Deployment > Publish to Teams > "Install for your organization".
5. Go back to perform the workflow above. **Approve the pending custom app**
![publishupdate1](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-publishupdate.png)
6. Your custom app also can be added into any Team's channel
![publishupdate2](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/teamsgpt-addtochannel.png)

## Referenced Architecture for Production Platform
![refarch](https://github.com/denlai-mshk/teamsgpt-azureopenai/blob/main/screenshots/refarch.png)
- If your Proof of Concept/Point of View runs smoothly, it may be a good idea to set up a production platform for service deployment. Here is the reference architecture.
  - **Public Endpoint Protection:** The Teams client sends messages over the public internet to your App Service. Utilizing FrontDoor or Application Gateway with Web Application Firewall can provide a high SLA to your service, protecting it from DDoS attacks. [DDoS Network Protection with PaaS web application architecture](https://learn.microsoft.com/en-us/azure/ddos-protection/ddos-protection-reference-architectures#ddos-network-protection-with-paas-web-application-architecture)
  - **Throttle Control:** Imagine what would happen if your users became addicted to this service. To protect your backend resources from spamming activities, API Management offers a [Rate limits and quotas policy](https://learn.microsoft.com/en-us/azure/api-management/api-management-sample-flexible-throttling#rate-limits-and-quotas).
  - **Chat Session:** Every single message sent from Teams client is stateless as well as the ChatGPT model API. Bot framework SDK offers waterfall dialogs to support a stateful conversation flow. Alternatively, using Azure Redis Cache for implementation brings more feasibility and control points to your custom app logic.
  - **Key and Secret Management:** The AOAI resource's KEY, endpoint and model name should not be stored in environmental variables, as this is not the best security practice. Instead, they should be stored in Azure Key Vault for better management of the application secret.
  - **Sensitive Query Censorship:** The custom app workload pane can be used to perform sensitive data filtering, alert and monitoring and blocking. For example, App Service can query Microsoft Purview data catalog to get the definition of sensitive data, e,g. RegEx to determine if a user query is compliant with enterprise IT policy or not. 
  - **Networking Security:** By using VNET injection, App Gateway, APIM, App Service, and Redis can be protected within the confines of a VNET. Meanwhile, Purview, Bot service and AOAI can take advantage of VNET integration (privatelink) to ensure greater security compliance.

## Custom app logic To-Do List
- AOAI offers a range of GPT models beyond ChatGPT. With these models, we can do more than just chat with a bot - such as summarizing documents, generating product descriptions, which leveraging text-davinci-003 and text-embedding-ada-002
- To further enhance custom app capabilities, here are some to-do recommendations.
  - Using Adaptive Cards (Teams Toolkit) to add a button for regenerating the user's question with different model parameters.
  - Besides of using Redis to implement a stateful conversation, you can consider to adopt the out of box capability provided by [Bot Framework SDK](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-basics?view=azure-bot-service-4.0): [Waterfall dialogs](https://learn.microsoft.com/en-us/azure/bot-service/bot-builder-concept-waterfall-dialogs?view=azure-bot-service-4.0#waterfall-dialogs), which supports stateful complex conversation flow.
![waterfall](https://learn.microsoft.com/en-us/azure/bot-service/v4sdk/media/bot-builder-dialog-concept.png)  
  - Add command set to let power user
    - send query with specified model parameters like: max. token, temperature
    - summarizing document or text generation (*text-davinci-003*)
    - classify and translate the text (*text-embedding-ada-002*)
   
## References
- [Publish Teams apps using Teams Toolkit](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/publish)
- [Get started for building and deploying customized apps for Microsoft Teams](https://learn.microsoft.com/en-us/microsoftteams/platform/toolkit/teams-toolkit-fundamentals?pivots=visual-studio-code)
- [Teams app capabilities and development tools](https://learn.microsoft.com/en-us/microsoftteams/platform/get-started/get-started-overview#app-capabilities-and-development-tools)
- [Working with the ChatGPT and GPT-4 models (preview)](https://learn.microsoft.com/en-us/azure/cognitive-services/openai/how-to/chatgpt?pivots=programming-language-chat-completions)
- [Teams Toolkit Visual Studio Code v5.0 Prerelease Guide](https://github.com/OfficeDev/TeamsFx/wiki/Teams-Toolkit-Visual-Studio-Code-v5.0-Prerelease-Guide)