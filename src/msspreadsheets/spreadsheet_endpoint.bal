// Copyright (c) 2020 WSO2 Inc. (http://www.wso2.org) All Rights Reserved.
//
// WSO2 Inc. licenses this file to you under the Apache License,
// Version 2.0 (the "License"); you may not use this file except
// in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing,
// software distributed under the License is distributed on an
// "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
// KIND, either express or implied.  See the License for the
// specific language governing permissions and limitations
// under the License.

import ballerina/http;
import ballerina/log;
import ballerina/oauth2;
import ballerina/lang.'int as ints;
import ballerina/io;

# Microsoft Spreadsheet Client Object.
# + msSpreadsheetClient - HTTP client endpoint for the spreadsheet API
# + microsoftGraphConfig - Configurations for accessing spreadsheet API
public type MSSpreadsheetClient client object {
    http:Client msSpreadsheetClient;
    MicrosoftGraphConfiguration microsoftGraphConfig;

    public function __init(MicrosoftGraphConfiguration msGraphConfig) {
        self.microsoftGraphConfig = msGraphConfig;
        oauth2:OutboundOAuth2Provider oauth2Provider3 = new ({
            accessToken: msGraphConfig.msInitialAccessToken,
            refreshConfig: {
                clientId: msGraphConfig.msClientID,
                clientSecret: msGraphConfig.msClientSecret,
                refreshToken: msGraphConfig.msRefreshToken,
                refreshUrl: msGraphConfig.msRefreshURL,
                clientConfig: {
                    secureSocket: {
                        trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                        }
                    }
                }
            }
        });
        http:BearerAuthHandler oauth2Handler3 = new (oauth2Provider3);

        self.msSpreadsheetClient = new (msGraphConfig.baseUrl, {
            auth: {
                authHandler: oauth2Handler3
            },
            secureSocket: {
                trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                }
            }
        });
    }

    # Open a Workbook by the given name.
    # + path - Path to the workbook file
    # + workbookName - Name of the Workbook
    # + return - A Workbook client object on success, else returns an error
    public remote function openWorkbook(string path, string workbookName) returns Workbook|error {
        Workbook workBook = new(self.microsoftGraphConfig, path, workbookName);

        return workBook;
    }
};

# Workbook Client Object.
# + workbookClient - HTTP client endpoint for the workbook
# + properties - Workbook specific properties
# + microsoftGraphConfig - Configurations for accessing spreadsheet API
public type Workbook client object {
    http:Client workbookClient;
    WorkbookProperties properties = {"path":"", "workbookName":""};
    MicrosoftGraphConfiguration microsoftGraphConfig;

    public function __init(MicrosoftGraphConfiguration msGraphConfig, string path, string workbookName) {
        self.microsoftGraphConfig = msGraphConfig;
        self.properties = {"path":path, "workbookName":workbookName};

        oauth2:OutboundOAuth2Provider oauth2Provider3 = new ({
            accessToken: msGraphConfig.msInitialAccessToken,
            refreshConfig: {
                clientId: msGraphConfig.msClientID,
                clientSecret: msGraphConfig.msClientSecret,
                refreshToken: msGraphConfig.msRefreshToken,
                refreshUrl: msGraphConfig.msRefreshURL,
                clientConfig: {
                    secureSocket: {
                        trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                        }
                    }
                }
            }
        });

        http:BearerAuthHandler oauth2Handler3 = new (oauth2Provider3);

        self.workbookClient = new (msGraphConfig.baseUrl, {
            auth: {
                authHandler: oauth2Handler3
            },
            secureSocket: {
                trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                }
            }
        });
    }

    # Get the properties of the workbook.
    # + return - Properties of the Workbook
    public function getProperties() returns WorkbookProperties {
        return self.properties;
    }

    # Open a worksheet on this workbook.
    # + worksheetName - name of the worksheet to be opened
    # + return - A Worksheet client object on success, else returns an error
    public remote function openWorksheet(string worksheetName) returns @tainted (Worksheet|error) {
        http:Request request = new;
        http:Response|error httpResponse = self.workbookClient->get("/v1.0/me/drive/root:" + self.properties.path +
        self.properties.workbookName + ".xlsx:/workbook/worksheets/" + worksheetName, request);
        int position = -1;
        string sheetId = "";

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_OK) {
                json|error response = httpResponse.getJsonPayload();
                io:println("|" + response.toString() + "|");
                if (response is map<json>) {
                    json sheetIdItem = response["id"];
                    sheetId = sheetIdItem.toString();
                    json positionItem = response["position"];

                    int|error res1 = ints:fromString(positionItem.toString());
                    if (res1 is int) {
                        position = res1;
                    } else {
                        error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while " +
                                                                            "getting the worksheet position.");
                        return err;
                    }

                    Worksheet workSheet = new(self.microsoftGraphConfig, self.properties.path,
                                            self.properties.workbookName, sheetId, worksheetName, position);
                    return workSheet;
                } else {
                    error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while opening the worksheet.");
                    return err;
                }
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while opening the worksheet.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }
    }

    # Create a worksheet on this workbook.
    # + worksheetName - name of the worksheet to be created
    # + return - A Worksheet client object on success, else returns an error
    public remote function createWorksheet(string worksheetName) returns @tainted (Worksheet|error) {
        http:Request request = new;
        json payload = {"name": worksheetName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.workbookClient->post("/v1.0/me/drive/root:" +
                    self.properties.path + self.properties.workbookName + ".xlsx:/workbook/worksheets", request);
        int position = -1;
        string sheetId = "";

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_CREATED) {
                json|error response = httpResponse.getJsonPayload();
                if (response is map<json>) {
                    json sheetIdItem = response["id"];
                    sheetId = sheetIdItem.toString();
                    json positionItem = response["position"];

                    int|error res1 = ints:fromString(positionItem.toString());
                    if (res1 is int) {
                        position = res1;
                    } else {
                        error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while " +
                                                                            "getting the worksheet position.");
                        return err;
                    }

                    Worksheet workSheet = new(self.microsoftGraphConfig, self.properties.path,
                                            self.properties.workbookName, sheetId, worksheetName, position);
                    return workSheet;
                } else {
                    error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while creating the worksheet.");
                    return err;
                }
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while creating the worksheet.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }
    }

    # Remove a worksheet from this workbook.
    # + worksheetName - name of the worksheet to be removed
    # + return - boolean true on success, else returns an error
    public remote function removeWorksheet(string worksheetName) returns @tainted (boolean|error) {
        boolean result = false;
        http:Request request = new;
        http:Response|error httpResponse = self.workbookClient->delete("/v1.0/me/drive/root:/" +
            self.properties.workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_NO_CONTENT) {
                result = true;
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while deleting the worksheet.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }

        return result;
    }
};

# Worksheet Client Object.
# + worksheetClient - HTTP client endpoint for the worksheet
# + properties - worksheet specific properties
# + microsoftGraphConfig - Configurations for accessing spreadsheet API
public type Worksheet client object {
    http:Client worksheetClient;
    WorksheetProperties properties = {"path":"", "workbookName":"", "sheetId":"",
                                                "worksheetName":"", "position":0};
    MicrosoftGraphConfiguration microsoftGraphConfig;

    public function __init(MicrosoftGraphConfiguration msGraphConfig, string path, string workbookName, string sheetId,
                            string worksheetName, int position) {
        self.microsoftGraphConfig = msGraphConfig;
        self.properties = {"path":path, "workbookName":workbookName, "sheetId":sheetId,
                            "worksheetName":worksheetName, "position":position};
        oauth2:OutboundOAuth2Provider oauth2Provider3 = new ({
            accessToken: msGraphConfig.msInitialAccessToken,
            refreshConfig: {
                clientId: msGraphConfig.msClientID,
                clientSecret: msGraphConfig.msClientSecret,
                refreshToken: msGraphConfig.msRefreshToken,
                refreshUrl: msGraphConfig.msRefreshURL,
                clientConfig: {
                    secureSocket: {
                        trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                        }
                    }
                }
            }
        });

        http:BearerAuthHandler oauth2Handler3 = new (oauth2Provider3);

        self.worksheetClient = new (msGraphConfig.baseUrl, {
            auth: {
                authHandler: oauth2Handler3
            },
            secureSocket: {
                trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                }
            }
        });
    }

    # Get the properties of the Worksheet.
    # + return - Properties of the Worksheet
    public function getProperties() returns WorksheetProperties {
        return self.properties;
    }

    # Create a new Table on this Worksheet.
    # + tableName - name of the table to be created
    # + address - The location where the table should be created
    # + return - A Table client object on success, else returns an error
    public remote function createTable(string tableName, string address) returns @tainted (Table|error) {
        http:Request request = new;
        json payload = {"name": tableName, "address": address, "hasHeaders": false};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.worksheetClient->post("/v1.0/me/drive/root:" + self.properties.path +
        self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName +
        "/tables/add", request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_CREATED) {
                json|error response = httpResponse.getJsonPayload();
                if (response is map<json>) {
                    json nameItem = response["name"];
                    string createdTableName = nameItem.toString();
                    json newTableID = response["id"];
                    string tableID = newTableID.toString();

                    Table resultsTable = <@untainted> new(self.microsoftGraphConfig, self.properties.path,
                    self.properties.workbookName, self.properties.sheetId, self.properties.worksheetName, tableID,
                    address, createdTableName);

                    if (createdTableName != tableName) {
                        log:printInfo("Table created (" + createdTableName + ") carries different name than what " +
                        "was passed as the table name (" + tableName + "). Now patching the table with the correct " +
                        "table name.");
                        boolean|error renameResult = resultsTable->renameTable(tableName);

                        if (renameResult is boolean) {
                            if (renameResult) {
                                return resultsTable;
                            } else {
                                error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while renaming " +
                                                                                    "the created table.");
                                return err;
                            }
                        } else {
                            error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while creating the table.");
                            return err;
                        }
                    } else {
                        return resultsTable;
                    }
                } else {
                    error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while creating the table.");
                    return err;
                }
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error ocurred while creating the table.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }
    }

    # Open a Table.
    # + tableName - name of the table to be opened
    # + return - A Table client object on success, else returns an error
    public remote function openTable(string tableName) returns @tainted Table|error {
        boolean result = false;
        http:Request request = new;
        http:Response|error httpResponse = self.worksheetClient->get("/v1.0/me/drive/root:" + self.properties.path +
        self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName + "/tables/" +
        tableName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_OK) {
                json|error response = httpResponse.getJsonPayload();
                if (response is map<json>) {
                    string identifier = response["id"].toString();
                    string address = "";
                    Table resultsTable = new(self.microsoftGraphConfig, self.properties.path, self.properties.workbookName,
                    self.properties.sheetId, self.properties.worksheetName, identifier, address, tableName);
                    return resultsTable;
                } else {
                    error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while inserting data into table.");
                    return err;
                }
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while inserting data into table.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }
    }
};

# Table Client Object.
# + tableClient - HTTP client endpoint for the table
# + properties - table specific properties
public type Table client object {
    http:Client tableClient;
    TableProperties properties = {"path":"", "workbookName":"", "sheetId":"",
                                                "worksheetName":"", "address":"", "tableID":"", "tableName":""};

    public function __init(MicrosoftGraphConfiguration msGraphConfig, string path, string workbookName, string sheetId,
                            string worksheetName, string tableID, string address, string tableName) {
        self.properties = {"path":path, "workbookName":workbookName, "sheetId":sheetId,
                            "worksheetName":worksheetName, "tableID":tableID, "address":address, "tableName":tableName};

        oauth2:OutboundOAuth2Provider oauth2Provider3 = new ({
            accessToken: msGraphConfig.msInitialAccessToken,
            refreshConfig: {
                clientId: msGraphConfig.msClientID,
                clientSecret: msGraphConfig.msClientSecret,
                refreshToken: msGraphConfig.msRefreshToken,
                refreshUrl: msGraphConfig.msRefreshURL,
                clientConfig: {
                    secureSocket: {
                        trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                        }
                    }
                }
            }
        });

        http:BearerAuthHandler oauth2Handler3 = new (oauth2Provider3);

        self.tableClient = new (msGraphConfig.baseUrl, {
            auth: {
                authHandler: oauth2Handler3
            },
            secureSocket: {
                trustStore: {
                            path: msGraphConfig.trustStorePath,
                            password: msGraphConfig.trustStorePassword
                }
            }
        });
    }

    # Get the properties of the table.
    # + return - Properties of the Table
    public function getProperties() returns TableProperties {
        return self.properties;
    }

    # Insert data into the table.
    # + data - data to be inserted into the table
    # + return - boolean true on success, else returns an error
    public remote function insertDataIntoTable(json data) returns boolean|error {
        boolean result = false;
        http:Request request = new;
        json payload = data;
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->post("/v1.0/me/drive/root:" + self.properties.path +
        self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName + "/tables/" +
        self.properties.tableName + "/rows/add", request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_CREATED) {
                result = true;
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while inserting data into table.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }

        return result;
    }

    # Rename the table.
    # + newTableName - new name to be used with the table
    # + return - boolean true on success, else returns an error
    public remote function renameTable(string newTableName) returns @tainted (boolean|error) {
        boolean result = false;
        http:Request request = new;
        json payload = {"name": newTableName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->patch("/v1.0/me/drive/root:" + self.properties.path +
        self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName +
        "/tables/" + self.properties.tableName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_OK) {
                self.properties.tableName = newTableName;
                result = true;
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while renaming the table.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }

        return result;
    }

    # Set a table's header.
    # + tableName - name of the table to be changed
    # + columnID - ID of the tableColumn to change
    # + headerName - new name of the table header
    # + return - boolean true on success, else returns an error
    public remote function setTableHeader(string tableName, int columnID, string headerName) returns boolean|error {
        boolean result = false;
        http:Request request = new;
        json payload = {"name": headerName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->patch("/v1.0/me/drive/root:" + self.properties.path +
        self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName + "/tables/" +
        tableName + "/columns/" + columnID.toString(), request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == HTTP_STATUS_OK) {
                result = true;
            } else {
                error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while setting the table header.");
                return err;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API.");
            return err;
        }

        return result;
    }
};

# Microsoft Graph client configuration.
# + baseUrl - The Microsoft Graph endpoint URL
# + msInitialAccessToken - Initial access token
# + msClientID - Microsoft client identifier
# + msClientSecret - client secret
# + msRefreshToken - refresh token
# + msRefreshURL - refresh URL
# + trustStorePath - trust store path
# + trustStorePassword - trust store password
# + bearerToken - bearer token
# + clientConfig - OAuth2 direct token configuration
public type MicrosoftGraphConfiguration record {
    string baseUrl;
    string msInitialAccessToken;
    string msClientID;
    string msClientSecret;
    string msRefreshToken;
    string msRefreshURL;
    string trustStorePath;
    string trustStorePassword;
    string bearerToken;
    oauth2:DirectTokenConfig clientConfig;
};