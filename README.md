# calling-cards

## Summary

Calling Cards can be used to create contact/general information tiles for agency/directorate leadership members. It gives the option for their name, email, multiple phone numbers, a link to their biography (if needed), and uploading an image. 


## Used SharePoint Framework Version

This uses SPfx version 1.14.0

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

> Any special pre-requisites?

Required dependencies:
"@pnp/sp": "^3.2.0",
"@pnp/spfx-controls-react": "^3.7.2",
"@pnp/spfx-property-controls": "^3.6.0",

## Solution

CallingCards|Jordan Wallingsford
------------|---------
contactCards| DPAA Software Dev

## Version history

Version|Date|Comments
-------|----|--------
1.0 | Initial release 5/24/2022

## Disclaimer

**THIS CODE IS PROVIDED *AS IS* WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- in the command-line run:
  - **npm install**
  - **gulp serve**

Change serve.json "initial page" to your agencies tenant site. 

## Features

Calling cards are used to display contact information for key personnel within the agency, or can be used for whatever you feel like... They display email (hyperlinked to open default mail application), Name (hyperlinked to open a link to their personal Bio if there is one), phone numbers (Duty, DNS, cell), their branch, and their position within the agency/directorate. Once installed, make sure to change the link within serve.json to your SPO tenant in order to be able to run the webpart. 

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Building for Microsoft teams](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/build-for-teams-overview)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/publish-to-marketplace-overview)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
