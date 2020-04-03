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
        http:Response|error response = self.workbookClient->get("/v1.0/me/drive/root:" + self.properties.path +
                            self.properties.workbookName + ".xlsx:/workbook/worksheets/" + worksheetName, request);

        if response is error {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
                                        errorCode = HTTP_ERROR, cause = response);
            return httpError;
        }

        http:Response httpResponse = <http:Response> response;

        if httpResponse.statusCode != http:STATUS_CREATED {
            HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while opening the worksheet.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
            return httpResponseHandlingError;
        }

        //If the worksheet is available we will get a JSON response with the worksheet's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid identifier; expected a `map<json>` found " +  typeOfRespone.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        map<json> payload = <map<json>> responseJson;
    
        json|error identifier = payload.id;

        if !(identifier is string) {
            typedesc<any|error> typeOfIdentifier = typeof identifier;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid identifier; expected a `string` found " +  typeOfIdentifier.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        string sheetId = <string> identifier;

        json|error sheetPosition = payload.position;

        if !(sheetPosition is int) {
            typedesc<any|error> typeOfPosition = typeof sheetPosition;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid sheet position; expected a `int` found " +  typeOfPosition.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        int position = <int> sheetPosition;
        
        //Populate a new worksheet object using the information from the properties as well as from the received JSON object
        Worksheet workSheet = new(self.microsoftGraphConfig, self.properties.path,
                                self.properties.workbookName, sheetId, worksheetName, position);
        return workSheet;
    }

    # Create a worksheet on this workbook.
    # + worksheetName - name of the worksheet to be created
    # + return - A Worksheet client object on success, else returns an error
    public remote function createWorksheet(string worksheetName) returns @tainted (Worksheet|error) {
        //Make a POST request and create the worksheet
        http:Request request = new;
        json payload = {"name": worksheetName};
        request.setJsonPayload(payload);
        http:Response|error response = self.workbookClient->post("/v1.0/me/drive/root:" +
                    self.properties.path + self.properties.workbookName + ".xlsx:/workbook/worksheets", request);

        if response is error {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
                                        errorCode = HTTP_ERROR, cause = response);
            return httpError;
        }

        http:Response httpResponse = <http:Response> response;

        if httpResponse.statusCode != http:STATUS_CREATED {
            HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while creating the worksheet.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
            return httpResponseHandlingError;
        }

        //If the worksheet was created we will get a JSON response with the newly created worksheet's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid response; expected a `map<json>` found " +  typeOfRespone.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        map<json> responsePayload = <map<json>> responseJson;
    
        json|error identifier = responsePayload.id;

        if !(identifier is string) {
            typedesc<any|error> typeOfIdentifier = typeof identifier;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid identifier; expected a `string` found " +  typeOfIdentifier.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        string sheetId = <string> identifier;

        json|error sheetPosition = responsePayload.position;

        if !(sheetPosition is int) {
            typedesc<any|error> typeOfPosition = typeof sheetPosition;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid sheet position; expected an `int` found " +  typeOfPosition.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        int position = <int> sheetPosition;
        
        //Populate a new worksheet object using the information from the properties as well as from the received JSON object
        Worksheet workSheet = new(self.microsoftGraphConfig, self.properties.path,
                                self.properties.workbookName, sheetId, worksheetName, position);
        return workSheet;
    }

    # Remove a worksheet from this workbook.
    # + worksheetName - name of the worksheet to be removed
    # + return - boolean true on success, else returns an error
    public remote function removeWorksheet(string worksheetName) returns @tainted error? {
        http:Request request = new;
        http:Response|error httpResponse = self.workbookClient->delete("/v1.0/me/drive/root:/" +
            self.properties.workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_NO_CONTENT) {
                return ();
            } else {
                HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while deleting the worksheet.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
                return httpResponseHandlingError;
            }
        } else {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
            errorCode = HTTP_ERROR, cause = httpResponse);
            return httpError;
        }
    }
};

# Worksheet Client Object.
# + worksheetClient - HTTP client endpoint for the worksheet
# + properties - worksheet specific properties
# + microsoftGraphConfig - Configurations for accessing spreadsheet API
public type Worksheet client object {
    http:Client worksheetClient;
    WorksheetProperties properties;
    MicrosoftGraphConfiguration microsoftGraphConfig;

    public function __init(MicrosoftGraphConfiguration msGraphConfig, string path, string workbookName, string sheetId,
                            string worksheetName, int position) {
        self.microsoftGraphConfig = msGraphConfig;
        self.properties = {
            path: path, 
            workbookName: workbookName,
            sheetId: sheetId,
            worksheetName: worksheetName,
            position: position
        };
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
        //Make a POST request and create the table
        http:Request request = new;
        request.setJsonPayload({"name": tableName, "address": address, "hasHeaders": false});
        http:Response|error response = self.worksheetClient->post(<@untainted> ("/v1.0/me/drive/root:" + self.properties.path +
                        self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName +
                        "/tables/add"), request);

        if response is error {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
                                        errorCode = HTTP_ERROR, cause = response);
            return httpError;
        }

        http:Response httpResponse = <http:Response> response;

        if httpResponse.statusCode != http:STATUS_CREATED {
            HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while creating the table.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
            return httpResponseHandlingError;
        }

        //If the table was created we will get a JSON response with the newly created table's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid response; expected a `map<json>` found " +  typeOfRespone.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        map<json> payload = <map<json>> responseJson;
    
        json|error nameItem = payload.name;

        if !(nameItem is string) {
            typedesc<any|error> typeOfNameItem = typeof nameItem;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid name; expected a `string` found " +  typeOfNameItem.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        string createdTableName = <string> nameItem;

        json|error newTableID = payload.id;
    
        if !(newTableID is string) {
            typedesc<any|error> typeOfTableId = typeof newTableID;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid table ID; expected a `string` found " +  typeOfTableId.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        //Populate a new Table object using the information from the properties as well as from the received JSON object
        Table resultsTable = <@untainted> new (self.microsoftGraphConfig, self.properties.path,
                                            self.properties.workbookName, self.properties.sheetId, 
                                            self.properties.worksheetName, newTableID.toString(), address, 
                                            createdTableName);

        if (createdTableName == tableName) {
            return resultsTable;
        }

        log:printInfo("Table created (" + createdTableName + ") carries different name than what " +
                "was passed as the table name (" + tableName + "). Now patching the table with the correct " +
                "table name.");

        error? renameResult = resultsTable->renameTable(tableName);

        if (renameResult is ()) {
            return resultsTable;
        } else {
            TableError tableError = error(TABLE_ERROR_CODE, message = "Error ocurred while renaming the created table.", 
                        errorCode = TABLE_ERROR_CODE, cause = renameResult);
            return tableError;
        }
    }

    # Open a Table.
    # + tableName - name of the table to be opened
    # + return - A Table client object on success, else returns an error
    public remote function openTable(string tableName) returns @tainted Table|error {
        //Make a GET request and retrieve the table information
        http:Request request = new;
        http:Response|error response = self.worksheetClient->get(<@untainted> ("/v1.0/me/drive/root:" + self.properties.path +
                                            self.properties.workbookName + ".xlsx:/workbook/worksheets/" + 
                                            self.properties.worksheetName + "/tables/" + tableName), request);

        if response is error {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
                                        errorCode = HTTP_ERROR, cause = response);
            return httpError;
        }

        http:Response httpResponse = <http:Response> response;

        if httpResponse.statusCode != http:STATUS_OK {
            HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while inserting data into table.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
            return httpResponseHandlingError;
        }

        //If the table exists we will get a JSON response with the table's information
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfRespone = typeof responseJson;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid response; expected a `map<json>` found " +  typeOfRespone.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        map<json> payload = <map<json>> responseJson;
    
        json|error identifier = payload.id;
        if !(identifier is string) {
            typedesc<any|error> typeOfIdentifier = typeof identifier;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid identifier; expected a `string` found " +  typeOfIdentifier.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);
            return typeError;
        }

        string sheetIdentifier = <string> identifier;

        //Address is not returned from the above API call. Hence the address is initialized to an empty string
        string address = "";

        Table resultsTable = <@untainted> new (self.microsoftGraphConfig, self.properties.path, 
                    self.properties.workbookName, self.properties.sheetId, self.properties.worksheetName,
                    sheetIdentifier, address, tableName);

        return resultsTable;
    }
};

# Table Client Object.
# + tableClient - HTTP client endpoint for the table
# + properties - table specific properties
public type Table client object {
    http:Client tableClient;
    TableProperties properties;

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
    public remote function insertDataIntoTable(json data) returns error? {
        http:Request request = new;
        json payload = data;
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->post(<@untainted> ("/v1.0/me/drive/root:" + self.properties.path +
            self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName + "/tables/" +
            self.properties.tableName + "/rows/add"), request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_CREATED) {
                return ();
            } else {
                HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while inserting data into table.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
                return httpResponseHandlingError;
            }
        } else {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
                        errorCode = HTTP_ERROR, cause = httpResponse);
            return httpError;
        }
    }

    # Rename the table.
    # + newTableName - new name to be used with the table
    # + return - boolean true on success, else returns an error
    public remote function renameTable(string newTableName) returns @tainted error? {
        http:Request request = new;
        json payload = {"name": newTableName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->patch("/v1.0/me/drive/root:" + self.properties.path +
                    self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName +
                    "/tables/" + self.properties.tableName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_OK) {
                self.properties.tableName = newTableName;
                return ();
            } else {
                HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while renaming the table.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
                return httpResponseHandlingError;
            }
        } else {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
            errorCode = HTTP_ERROR, cause = httpResponse);
            return httpError;
        }
    }

    # Set a table's header.
    # + columnID - ID of the tableColumn to change
    # + headerName - new name of the table header
    # + return - boolean true on success, else returns an error
    public remote function setTableHeader(int columnID, string headerName) returns error? {
        http:Request request = new;
        json payload = {"name": headerName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.tableClient->patch(<@untainted> ("/v1.0/me/drive/root:" + self.properties.path +
            self.properties.workbookName + ".xlsx:/workbook/worksheets/" + self.properties.worksheetName + "/tables/" +
            self.properties.tableName + "/columns/" + columnID.toString()), request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == http:STATUS_OK) {
                return ();
            } else {
                HttpResponseHandlingError httpResponseHandlingError = error(HTTP_RESPONSE_HANDLING_ERROR,
                message = "Error occurred while setting the table header.", errorCode = HTTP_RESPONSE_HANDLING_ERROR);
                return httpResponseHandlingError;
            }
        } else {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
            errorCode = HTTP_ERROR, cause = httpResponse);
            return httpError;
        }
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