# CleanUnifiedGroup

CleanUnifiedGroup is a simple .netcore application that will delete all conversations older than some time period (default 1 day) in a Unified Group (O365 Group).

## Dependencies

.NET Core 3.1 SDK

##### The following nuget modules are required, they should be automatically restored when building/running application using dotnet commands

* Microsoft.Exchange.WebServices `v2.2.0`
* Microsoft.Identity.Client `v4.19.0`
* System.Configuration.ConfigurationManager `v4.7.0`
* System.DirectoryService `v4.7.0`

*Listed versions are validated with this application, other's may work but have not been tested*

## Configuration

### Azure AD Configuration

To utilize this application you must configure an enterprise application registration with the following settings:
1. Application must be configured with http://localhost in the redirect URIs. Also, redirect URI's should be set as "Mobile and desktop applications" and not the default of "web" for public client authentication ![AzurePortalScreenshot](https://i.imgur.com/dXFb08o.png)
1. Application must be set to "treat application as a public client![AzurePortalScreenshot](https://i.imgur.com/ToN6RIT.png)

### Configuration File

Rename the configuration file from Example.App.config to App.config and fill in the following values
```
<add key="appId" value="{YOUR_APP_ID}" />
<add key="tenantId" value="{YOUR_AAD_TENANT_ID}" />
<add key="smtpAddressOfUnifiedGroup" value="{SMTP_ADDRESS_OF_UNIFIED_GROUP}"/>
```

## Usage

To run this application run the following command
```
dotnet run
```

## Contributing

Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

Please make sure to update tests as appropriate.

## License
GNU General Public License v3.0
