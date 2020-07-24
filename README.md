# DEMO - MS Teams application calling external APIs (SSO)

This demo application shows authentication in a MS Teams application using single sign-on (SSO). The application is able to call external APIs:

- Directly with the SSO token:
  - Web API which is hosted in the same App Service as the front-end
- Using the on-behalf-of (OBO) flow:
  - Graph API
  - SharePoint REST API

## Settings and variables

When deploying the demo, make sure to update the settings/variables denoted by {{...}} or <...> in:

- TeamsSSO.WebApp/appsettings.json
- TeamsSSO.WebApp/ClientApp/src/App.tsx

## Custom domain

In order to enable SSO in your MS Teams you will need an application which is available on a custom domain name.

Currently the `*.azurewebsites.net` domain is not supported.

## Azure AD Application

When creating the Azure AD application, make sure to correctly format the Application ID URI like this:

- `api://<domain_name>/<client_id>`
  - e.g.: api://my.domain.com/00000000-0000-0000-0000-000000000000

The exposed scope should be named `access_as_user`.

## Useful resources

- [Single Sign-On (A)](https://docs.microsoft.com/en-us/microsoftteams/platform/tabs/how-to/authentication/auth-aad-sso) and [Single Sign-On (B)](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-on-behalf-of-flow#first-case-access-token-request-with-a-shared-secret)
- [OAuth2 On-Behalf-Of flow (OBO)](https://docs.microsoft.com/en-us/azure/active-directory/azuread-dev/v1-oauth2-on-behalf-of-flow)
- [ASP.NET Core React project template](https://docs.microsoft.com/en-us/aspnet/core/client-side/spa/react?view=aspnetcore-3.1&tabs=visual-studio)
