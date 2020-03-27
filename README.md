# Module Microsoft Graph

Connects to Microsoft Graph API which acts as the gateway to data and intelligence in Microsoft 365. Microsoft Graph exposes a unified programming model which can be used to access vast amounts of data available in Office 365, Windows 10, and Enterprise Mobility & Security.

The current implementation of Microsoft Graph consists of the following sub modules.

**Spreadsheet Operations**

The `wso2/msspreadsheets` module contains operations to perform CRUD (Create, Read, Update, and Delete) operations on [Excel workbooks](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0) stored in Microsoft OneDrive.

**OneDrive Operations**
Microsoft OneDrive is a file hosting service and synchronization service run by Microsoft. It is operated as part of Microsoft Office 365.
The `wso2/msonedrive` module contains operations for accessing the items stored in OneDrive.

## Compatibility
|                     |    Version     |
|:-------------------:|:--------------:|
| Ballerina Language  | 1.2.0   |
| Microsoftgraph REST API | v1.0          |

## Getting started

1.  Refer the [Getting Started](https://ballerina.io/learn/getting-started/) guide to download and install Ballerina.

2.  To use the Microsoft Graph API, you need to provide the following configuration information in the ballerina.conf file:

       - MS_CLIENT_ID
       - MS_CLIENT_SECRET
       - MS_ACCESS_TOKEN
       - MS_REFRESH_TOKEN
       - TRUST_STORE_PATH
       - TRUST_STORE_PASSWORD
       - WORK_BOOK_NAME
       - WORK_SHEET_NAME
       - TABLE_NAME

    Following steps should be followed to obtain the above mentioned configuration information.

    Before you run the following steps you may have to create an account in [OneDrive](https://onedrive.live.com). Next, sign into [Azure Portal - App Registrations](https://portal.azure.com/#blade/Microsoft_AAD_RegisteredApps/ApplicationsListBlade). You may have to use your personal or work or school account.

    From the App registrations page, click on New registration. Enter a meaningful name in the name field. 
    In the Supported account types section, select Accounts in any organizational directory and personal Microsoft accounts (e.g. Skype, Xbox, Outlook.com). Click Register to create the application.

    Copy the Application (client) ID (\<MS_CLIENT_ID>). This is the unique identifier for your app.
    In the application's list of pages (under the Manage tab in left hand side menu), select Authentication.
    Under the Platform configurations click on "Add a platform" button.
    Under the "Configure platforms", click on "Web" button located under the Web applications.

    Under the redirect URIs text box, put [OAuth2 Native Client](https://login.microsoftonline.com/common/oauth2/nativeclient).
    Under the Implicit grant select Access tokens.
    Click on Configure.
    Under the Certificates & Secrets, create a new client secret (\<MS_CLIENT_SECRET>). This requires providing a description and a period of expiry. Next, click on Add button.

    Next, we need to obtain an access token and a refresh token to invoke the Microsoft Graph API.
    First, in a new browser enter this URL by replacing the \<MS_CLIENT_ID> with the application ID.

    ```
    https://login.microsoftonline.com/common/oauth2/v2.0/authorize?response_type=code&client_id=<MS_CLIENT_ID>&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&scope=Files.ReadWrite openid User.Read Mail.Send Mail.ReadWrite offline_access
    ```

    This may prompt you to enter the username and password for signing into  Azure Portal App.

    Once the username password pair successfully entered this will give a URL like follows on the browser address bar.

    https://login.microsoftonline.com/common/oauth2/nativeclient?code=M95780001-0fb3-d138-6aa2-0be59d402f32

    Copy the code parameter (M95780001-0fb3-d138-6aa2-0be59d402f32 in the above example) and in a new terminal enter the following CURL command with replacing the \<MS_CODE> with the code received from the above step. The \<MS_CLIENT_ID> and \<MS_CLIENT_SECRET> parameters are the same as above.

    ```
    curl -X POST --header "Content-Type: application/x-www-form-urlencoded" --header "Host:login.microsoftonline.com" -d "client_id=<MS_CLIENT_ID>&client_secret=<MS_CLIENT_SECRET>&grant_type=authorization_code&redirect_uri=https://login.microsoftonline.com/common/oauth2/nativeclient&code=<MS_CODE>&scope=Files.ReadWrite openid User.Read Mail.Send Mail.ReadWrite offline_access" https://login.microsoftonline.com/common/oauth2/v2.0/token
    ```

    The above CURL command should result in a response as follows,
    ```
    {
    "token_type": "Bearer",
    "scope": "Files.ReadWrite openid User.Read Mail.Send Mail.ReadWrite",
    "expires_in": 3600,
    "ext_expires_in": 3600,
    "access_token": "<MS_ACCESS_TOKEN>",
    "refresh_token": "<MS_REFRESH_TOKEN>",
    "id_token": "<ID_TOKEN>"
    }
    ```

    Set the path to your Ballerina distribution's trust store in the <TURST_STORE_PATH>. This is by default located in the following path.

    $BALLERINA_HOME/distributions/jballerina-<BALLERINA_VERSION>/bre/security/ballerinaTruststore.p12

    The default TRUST_STORE_PASSWORD is set to "ballerina".

    WORK_BOOK_NAME, WORK_SHEET_NAME, and TABLE_NAME corresponds to workbook file name (without the .xlsx extension), worksheet name, and the table name respectively. Makesure you create a workbook with the same WORK_BOOK_NAME on Microsoft OneDrive before using the connector.

3. Create a new Ballerina project by executing the following command.

	```shell
	<PROJECT_ROOT_DIRECTORY>$ ballerina init
	```

4. Import the Microsoft Graph connector to your Ballerina program as follows. The following sample program creates a new worksheet on an existing workbook on Microsoft OneDrive. Prior running this application please create a workbook on your Microsoft OneDrive account
having the name "MyShop.xlsx". There needs to be at least one worksheet (i.e., a tab) on the workbook for this sample program to work. Note that the sample application first tries to delete an existing worksheet named "Sales" from the workbook. If its not available it may throw an error and continue executing the rest of the program. This error may get thrown during the very first round of running the sample application. Makesure, you keep the ballerina.conf file with the above mentioned configuration information before running the
sample application.

## Sample Application

	```ballerina
	import ballerina/config;
    import ballerina/log;
    import ballerina/time;
    import wso2/msspreadsheets;
    import wso2/msonedrive;

    // Create Microsoft Graph Client configuration by reading from config file.
    msspreadsheets:MicrosoftGraphConfiguration msGraphConfig = {
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

    msspreadsheets:MSSpreadsheetClient msSpreadsheetClient = new(msGraphConfig);
    
    string WORK_BOOK_NAME = "MyShop";
    string WORK_SHEET_NAME = "Sales";
    string TABLE_NAME = "tbl";

    public function main() {
        msspreadsheets:Workbook|error workbook = msSpreadsheetClient->openWorkbook("/", WORK_BOOK_NAME);

        if (workbook is msspreadsheets:Workbook) {
            boolean|error resultRemove = workbook->removeWorksheet(WORK_SHEET_NAME);
            if (resultRemove is boolean) {
                log:printInfo("Worksheet was deleted");
            } else {
                log:printError("Could not delete the Worksheet", err = resultRemove);
            }

            msspreadsheets:Worksheet|error sheet = workbook->createWorksheet(WORK_SHEET_NAME);

            if (sheet is msspreadsheets:Worksheet) {
                log:printInfo("Worksheet was created");
                msspreadsheets:Table|error resultTable = sheet->createTable(TABLE_NAME, <@untainted> ("A1:E1"));

                if (resultTable is msspreadsheets:Table) {
                    boolean|error resultHeader = resultTable->setTableHeader(TABLE_NAME, 1, "ID");
                    resultHeader = resultTable->setTableHeader(TABLE_NAME, 2, "DateSold");
                    resultHeader = resultTable->setTableHeader(TABLE_NAME, 3, "ItemID");
                    resultHeader = resultTable->setTableHeader(TABLE_NAME, 4, "ItemName");
                    resultHeader = resultTable->setTableHeader(TABLE_NAME, 5, "Price");

                    if (resultHeader is boolean) {
                        json[][] valuesString=[];
                        time:Time time = time:currentTime();
                        string|error cString1 = time:format(time, "yyyy-MM-dd'T'HH:mm:ss.SSSZ");
                        string customTimeString = "";
                        if (cString1 is string) {
                            customTimeString = cString1;
                        }

                        foreach int counter in 1...5 {
                            int itemID = counter + 100;
                            json[] arr = [ counter.toString(), customTimeString, 
                            itemID.toString(), "Item-" + itemID.toString(), "10" ];
                            valuesString.push(arr);
                        }
                        json data = {"values": valuesString};
                        boolean|error result = resultTable->insertDataIntoTable(<@untainted> data);

                        if (result is boolean) {
                            if (result) {
                                log:printInfo("Inserted data into table");

                                msonedrive:OneDriveClient msOneDriveClient = new(msGraphConfig);
                                msonedrive:Item|error item = msOneDriveClient->getItemFromRoot(WORK_BOOK_NAME + ".xlsx");
                                if (item is msonedrive:Item) {
                                    log:printInfo("The URL of the workbook (" + item.name.toString() + ") is " + item.webUrl.toString());
                                } else {
                                     log:printError("Error getting the spreadsheet URL", err = item);
                                }
                            } else {
                                log:printError("Error inserting data into the table");
                            }
                        } else {
                            log:printError("Error inserting data into the table");
                        }
                    } else {
                        log:printError("Error setting table headers", err = resultHeader);
                    }
                } else {
                    log:printError("Error creating table", err = resultTable);
                }
            } else {
                log:printError("Error opening worksheet", err = sheet);
            }
        } else {
            log:printError("Error opening workbook", err = workbook);
        }
    }
	```