# CleanUnifiedGroup

CleanUnifiedGroup is a simple .netcore application that will delete all conversations older than some time period (default 1 day) in a Unified Group (O365 Group).

## Configuration

### Azure AD Configuration

To utilize this application you must configure an enterprise application registration with the following settings:
1. Application must be configured with http://localhost in the redirect URIs. Also, redirect URI's should be set as "Mobile and desktop applications" and not the default of "web" for public client authentication ![AzurePortalScreenshot](https://i.imgur.com/dXFb08o.png)
1. Application must be set to "treat application as a public client![AzurePortalScreenshot](https://i.imgur.com/ToN6RIT.png)

## Usage

## Contributing

## License
GNU General Public License v3.0
