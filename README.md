# SharePoint User Permissions Report

This tool (console application) allows you to fetch and save a report (`.xlsx`) with the permissions of a specific user against one or many SharePoint site collections.

## SharePoint connection

There are two ways to make the tool connect to SharePoint:

### SharePoint App Add-in

You can put a SharePoint Add-in Client ID & Secret in the `appSettings` section of the `App.config` file. Available settings:

```
...
<appSettings>
  <add key="tenantAdminUrl" value="https://{tenant}-admin.sharepoint.com" />
  <add key="clientId" value="{Add-in Client ID}" />
  <add key="clientSecret" value="{Add-in Client Secret}" />
</appSettings>
...
```

**Note:** the Add-in should have `full-control` permissions on the tenant. For more information, visit [this page](https://docs.microsoft.com/en-us/sharepoint/dev/solution-guidance/security-apponly-azureacs).

### Tenant administrator account

If you have an account with administrative privileges on the tenant, you can use web login. Just leave the `appSettings` empty (except maybe for the `tenantAdminUrl`) and the tool will take you to Microsoft's login page.

## How to use

Run the app on the `dist` folder or debug the source code in Visual Studio. It will ask you for the user, site url (you can use wilcards to filter the sites you want to target `e.g. */sites/my-site*`) and local path with file name (`*.xlsx`) to save the results.