This module allows users to connect to Microsoft OneDrive and provides information on the files, which have been stored on OneDrive.

# Module Overview
This module contains operations for accessing the items stored in OneDrive.

## Supported Operations
- Get an item from the root directory
- Get an item from a non-root directory

## Compatibility
|                     |    Version     |
|:-------------------:|:--------------:|
| Ballerina Language  | 1.2.x   |
| Microsoftgraph REST API | v1.0          |

## Sample
Instantiate the connector by giving authentication details in an HTTP client config. The HTTP client config has built-in support for BasicAuth and OAuth 2.0. Microsoft Graph API uses OAuth 2.0 to authenticate and authorize requests. 

**Obtaining configuration information**
The Microsoft Graph connector can be minimally instantiated in the HTTP client config using the access token (\<MS_ACCESS_TOKEN>), the client ID (\<MS_CLIENT_ID>), the client secret (\<MS_CLIENT_SECRET>), and the refresh token (\<MS_REFRESH_TOKEN>). Specific details of obtaining these values is mentioned in the [README](https://github.com/ballerina-platform/module-microsoftgraph/blob/master/README.md).

**Add project configurations file**

Add the project configuration file by creating a `ballerina.conf` file under the root path of the project structure. This file should have following configurations. Add the tokens obtained in the previous step to the `ballerina.conf` file.

```
MS_BASE_URL="https://graph.microsoft.com"
MS_CLIENT_ID="<MS_CLIENT_ID>"
MS_CLIENT_SECRET="<MS_CLIENT_SECRET>"
MS_REFRESH_TOKEN="<MS_REFRESH_TOKEN>"
MS_REFRESH_URL="https://login.microsoftonline.com/common/oauth2/v2.0/token"
MS_ACCESS_TOKEN="<MS_ACCESS_TOKEN>"
TRUST_STORE_PATH=""
TRUST_STORE_PASSWORD=""
WORK_BOOK_NAME=""
WORK_SHEET_NAME=""
TABLE_NAME=""
```

**Example Code**
Creating a `microsoft.onedrive1:OneDriveClient` by giving the HTTP client config details. The module `microsoft.onedrive1`
is referred as the module prefix `onedrive`.

```
    import ballerinax/microsoft.onedrive1 as onedrive;

    onedrive:MicrosoftGraphConfiguration msGraphConfig = {
        baseUrl: config:getAsString("MS_BASE_URL"),
        msInitialAccessToken: config:getAsString("MS_ACCESS_TOKEN"),
        msClientID: config:getAsString("MS_CLIENT_ID"),
        msClientSecret: config:getAsString("MS_CLIENT_SECRET"),
        msRefreshToken: config:getAsString("MS_REFRESH_TOKEN"),
        msRefreshURL: config:getAsString("MS_REFRESH_URL"),
        trustStorePath: config:getAsString("TRUST_STORE_PATH"),
        trustStorePassword: config:getAsString("TRUST_STORE_PASSWORD"),
        bearerToken: config:getAsString("MS_ACCESS_TOKEN"),
        clientConfig: {
            accessToken: config:getAsString("MS_ACCESS_TOKEN"),
            refreshConfig: {
                clientId: config:getAsString("MS_CLIENT_ID"),
                clientSecret: config:getAsString("MS_CLIENT_SECRET"),
                refreshToken: config:getAsString("MS_REFRESH_TOKEN"),
                refreshUrl: config:getAsString("MS_REFRESH_URL")
            }
        }
    };

    onedrive:OneDriveClient msOneDriveClient = new(msGraphConfig);
```

Getting an item from OneDrive's root directory. Here we are getting an item called `Book.xlsx` from the root.

```onedrive:Item|error item = msOneDriveClient->getItemFromRoot("Book.xlsx");```

Getting an item from OneDrive's directory other than root. Here we are getting an item called `Book.xlsx` from the location `/myfolder`

```onedrive:Item|error item = msOneDriveClient->getItemFromNonRoot("/myfolder", "Book.xlsx");```

Getting the URL of an item.

```
    onedrive:Item|error item = msOneDriveClient->getItemFromRoot("Book.xlsx");
    if (item is onedrive:Item) {
        log:printInfo("The URL of the workbook is " + item.webUrl.toString());
    } else {
        log:printError("Error getting the spreadsheet URL", err = item);
    }
```
