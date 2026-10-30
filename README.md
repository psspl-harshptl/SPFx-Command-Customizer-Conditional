# employee-operations-command-set

## Summary

A SharePoint Framework (SPFx) ListView Command Set that adds employee actions to the command bar of the `EmployeeDirectory` list. Commands appear conditionally, based on how many rows are selected.

| Command              | Visible when      | What it does                                                                                  |
| -------------------- | ----------------- | --------------------------------------------------------------------------------------------- |
| Employee Details     | Exactly 1 row     | Loads the employee from the list and shows their details in a dialog.                         |
| Bulk Status Update   | 2 or more rows    | Prompts for a new status, updates every selected employee, then reloads the page.             |
| Export Employees CSV | 1 or more rows    | Downloads the selected rows as `Employees.csv`.                                               |

Built with SPFx 1.23, TypeScript and [PnPjs](https://pnp.github.io/pnpjs/).

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.0-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)

> Get your own free development tenant by subscribing to [Microsoft 365 developer program](http://aka.ms/o365devprogram)

## Prerequisites

- Node.js 22 (`>=22.14.0 <23.0.0`)
- A SharePoint site with a custom list named exactly `EmployeeDirectory`, with these columns:

| Internal name    | Type                       | Notes                                                    |
| ---------------- | -------------------------- | -------------------------------------------------------- |
| `Title`          | Single line of text        | Employee name                                            |
| `EmployeeCode`   | Single line of text        |                                                          |
| `Department`     | Single line of text/Choice |                                                          |
| `Designation`    | Single line of text/Choice |                                                          |
| `EmployeeStatus` | Choice                     | Values: `Active`, `Inactive`, `On Leave`                 |
| `JoiningDate`    | Date                       |                                                          |
| `Manager`        | Single line of text        | A Person column won't work: the query doesn't expand it  |

The internal names must match, because the code refers to columns by these names.

## Getting started

1. Open the solution folder and install dependencies:

   ```bash
   npm install
   ```

2. In [config/serve.json](config/serve.json), set both `pageUrl` values to the `EmployeeDirectory` list in your own site. They currently point to a specific development tenant.

3. Start the local server:

   ```bash
   npm start
   ```

4. On the list page, choose **Load debug scripts** when prompted, then select rows to see the commands appear.

## Build and deploy

1. Build the production package:

   ```bash
   npm run build
   ```

   This creates `sharepoint/solution/employee-operations-command-set.sppkg`.

2. Upload the package to your App Catalog.

3. Either enable the app for all sites, or add it only to the site that holds `EmployeeDirectory`.

| Script          | Purpose                                           |
| --------------- | ------------------------------------------------- |
| `npm start`     | Clean, build and serve locally                    |
| `npm run build` | Run the production build and tests, then package  |
| `npm run clean` | Remove build output                               |

## How it works

- **Conditional visibility:** `_setCommandVisibility` runs on load and whenever the list view state changes, and shows or hides each command based on the selected row count.
- **Employee Details:** reads the item by ID from the list, so it shows current data even for columns that aren't in the view.
- **Bulk Status Update:** updates items in parallel. If any update fails, the error lists the failed item IDs; the other items may already have been updated.
- **Export Employees CSV:** runs in the browser using the values already loaded in the view, with no call to SharePoint.

### Project structure

```text
src/
├── extensions/employeeOperations/
│   ├── EmployeeOperationsCommandSet.ts            # Commands and visibility rules
│   └── EmployeeOperationsCommandSet.manifest.json # Command IDs and titles
├── services/EmployeeService.ts                    # Reads and updates list items
├── models/IListItem.ts                            # Employee item shape
├── components/BulkStatusUpdateDialog.tsx          # Fluent UI dialog (not used yet)
└── pnpjsConfig.ts                                 # PnPjs setup
```

## Known limitations

- **Commands show on every custom list.** The extension registers for all generic lists (template `100`), but the service always reads and updates `EmployeeDirectory`. On another list, Employee Details and Bulk Status Update would act on `EmployeeDirectory` items with the same IDs, so avoid enabling the app on sites with other custom lists until this is restricted.
- **Bulk status is typed by hand.** The command uses a text prompt, and the value must match `Active`, `Inactive` or `On Leave` exactly, including case. Other values are rejected with an error. `BulkStatusUpdateDialog` offers a dropdown instead, but isn't connected to the command yet.
- **CSV export depends on the view.** Columns missing from the current view export as empty values, and only selected rows are exported.
- **No unit tests yet.** `npm run build` runs the test step, but the project has no tests.

## Version history

| Version | Comments        |
| ------- | --------------- |
| 1.0.0.0 | Initial version |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Build your first ListView Command Set extension](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/get-started/building-simple-cmdset-with-dialog-api)
- [Tenant-wide deployment of SPFx extensions](https://learn.microsoft.com/sharepoint/dev/spfx/extensions/basics/tenant-wide-deployment-extensions)
- [PnPjs documentation](https://pnp.github.io/pnpjs/)
- [Heft documentation](https://heft.rushstack.io/)
