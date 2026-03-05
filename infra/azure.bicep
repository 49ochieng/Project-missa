@maxLength(20)
@minLength(4)
@description('Used to generate names for all resources in this file')
param resourceBaseName string

param webAppSKU string

@maxLength(42)
param botDisplayName string

param AOAI_ENDPOINT string
param AOAI_API_KEY string
param AOAI_MODEL string

@secure()
param meetingMediaBotSecret string = ''

// Meeting-media-bot service params
param meetingBotTenantId string = ''
param meetingBotClientId string = ''
@secure()
param meetingBotClientSecret string = ''
param meetingBotAppId string = ''
@secure()
param azureSpeechKey string = ''
param azureSpeechRegion string = 'southcentralus'

param sqlServer string
param sqlDatabase string
param sqlUsername string
@secure()
param sqlPassword string

param serverfarmsName string = resourceBaseName
param webAppName string = resourceBaseName
param meetingBotAppName string = 'meetingbot${resourceBaseName}'
param identityName string = resourceBaseName
param location string = resourceGroup().location

resource identity 'Microsoft.ManagedIdentity/userAssignedIdentities@2023-01-31' = {
  location: location
  name: identityName
}

// Compute resources for your Web App
resource serverfarm 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'app'
  location: location
  name: serverfarmsName
  sku: {
    name: webAppSKU
  }
}

// Web App that hosts your agent
resource webApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app'
  location: location
  name: webAppName
  properties: {
    serverFarmId: serverfarm.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      appSettings: [
        {
          name: 'WEBSITE_RUN_FROM_PACKAGE'
          value: '1' // Run Azure App Service from a package file
        }
        {
          name: 'WEBSITE_NODE_DEFAULT_VERSION'
          value: '~20' // Set NodeJS version to 20.x for your site
        }
        {
          name: 'RUNNING_ON_AZURE'
          value: '1'
        }
        {
          name: 'CLIENT_ID'
          value: identity.properties.clientId
        }
        {
          name: 'TENANT_ID'
          value: identity.properties.tenantId
        }
        {
          name: 'BOT_TYPE'
          value: 'UserAssignedMsi'
        }
        {
          name: 'AOAI_ENDPOINT'
          value: AOAI_ENDPOINT
        }
        {
          name: 'AOAI_API_KEY'
          value: AOAI_API_KEY
        }
        {
          name: 'AOAI_MODEL'
          value: AOAI_MODEL
        }
        {
          name: 'SQL_CONNECTION_STRING'
          value: 'Server=tcp:${sqlServer},1433;Initial Catalog=${sqlDatabase};Persist Security Info=False;User ID=${sqlUsername};Password=${sqlPassword};MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;'
        }
        {
          name: 'SQL_SERVER'
          value: sqlServer
        }
        {
          name: 'SQL_DATABASE'
          value: sqlDatabase
        }
        {
          name: 'SQL_USERNAME'
          value: sqlUsername
        }
        {
          name: 'SQL_PASSWORD'
          value: sqlPassword
        }
        {
          name: 'MEETING_MEDIA_BOT_URL'
          value: 'https://${meetingBotApp.properties.defaultHostName}'
        }
        {
          name: 'MEETING_MEDIA_BOT_SHARED_SECRET'
          value: meetingMediaBotSecret
        }
        {
          name: 'INTERNAL_API_PORT'
          value: '3980'
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
  identity: {
    type: 'UserAssigned'
    userAssignedIdentities: {
      '${identity.id}': {}
    }
  }
}

// Linux App Service Plan for meeting-media-bot (Linux + Windows can't share a plan)
resource meetingBotPlan 'Microsoft.Web/serverfarms@2021-02-01' = {
  kind: 'linux'
  location: location
  name: '${serverfarmsName}-linux'
  sku: {
    name: webAppSKU
  }
  properties: {
    reserved: true // Required for Linux
  }
}

// Meeting-media-bot App Service (joins Teams calls, transcribes meetings) — Linux
resource meetingBotApp 'Microsoft.Web/sites@2021-02-01' = {
  kind: 'app,linux'
  location: location
  name: meetingBotAppName
  properties: {
    serverFarmId: meetingBotPlan.id
    httpsOnly: true
    siteConfig: {
      alwaysOn: true
      linuxFxVersion: 'NODE|20-lts'
      appCommandLine: 'node dist/index.js'
      appSettings: [
        {
          name: 'AZURE_TENANT_ID'
          value: meetingBotTenantId
        }
        {
          name: 'AZURE_CLIENT_ID'
          value: meetingBotClientId
        }
        {
          name: 'AZURE_CLIENT_SECRET'
          value: meetingBotClientSecret
        }
        {
          name: 'BOT_APP_ID'
          value: meetingBotAppId
        }
        {
          name: 'BOT_ENDPOINT'
          value: 'https://${meetingBotAppName}.azurewebsites.net'
        }
        {
          name: 'AZURE_SPEECH_KEY'
          value: azureSpeechKey
        }
        {
          name: 'AZURE_SPEECH_REGION'
          value: azureSpeechRegion
        }
        {
          name: 'SHARED_SECRET'
          value: meetingMediaBotSecret
        }
        {
          name: 'PROJECT_MISSA_URL'
          value: 'https://${webAppName}.azurewebsites.net'
        }
        {
          name: 'PORT'
          value: '8080'
        }
      ]
      ftpsState: 'FtpsOnly'
    }
  }
}

// Register your web service as a bot with the Bot Framework
module azureBotRegistration './botRegistration/azurebot.bicep' = {
  name: 'Azure-Bot-registration'
  params: {
    resourceBaseName: resourceBaseName
    identityClientId: identity.properties.clientId
    identityResourceId: identity.id
    identityTenantId: identity.properties.tenantId
    botAppDomain: webApp.properties.defaultHostName
    botDisplayName: botDisplayName
  }
}

// The output will be persisted in .env.{envName}. Visit https://aka.ms/teamsfx-actions/arm-deploy for more details.
output BOT_AZURE_APP_SERVICE_RESOURCE_ID string = webApp.id
output BOT_DOMAIN string = webApp.properties.defaultHostName
output BOT_ID string = identity.properties.clientId
output BOT_TENANT_ID string = identity.properties.tenantId
output SQL_SERVER_FQDN string = sqlServer
output SQL_DATABASE_NAME string = sqlDatabase
output MEETING_BOT_AZURE_APP_SERVICE_RESOURCE_ID string = meetingBotApp.id
output MEETING_BOT_DOMAIN string = meetingBotApp.properties.defaultHostName
