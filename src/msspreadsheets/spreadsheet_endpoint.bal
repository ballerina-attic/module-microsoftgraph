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
import ballerina/io;
import ballerina/config;
import ballerina/oauth2;
import ballerina/log;

# Microsoft Spreadsheet Client Object
public type MSSpreadsheetClient client object {
    http:Client msSpreadsheetClient;

    public function __init(MicrosoftGraphConfiguration msGraphConfig) {
        oauth2:OutboundOAuth2Provider oauth2Provider3 = new({
            accessToken: config:getAsString("MS_INITIAL_ACCESS_TOKEN"),
            refreshConfig: {
                clientId: config:getAsString("MS_CLIENT_ID"),
                clientSecret: config:getAsString("MS_CLIENT_SECRET"),
                refreshToken: config:getAsString("MS_REFRESH_TOKEN"),
                refreshUrl: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
                clientConfig: {
                    secureSocket: {
                        trustStore: {
                            path: config:getAsString("KEYSTORE_PATH"),
                            password: config:getAsString("KEYSTORE_PASSWORD")
                        }
                    }
                }
            }
        });
        http:BearerAuthHandler oauth2Handler3 = new(oauth2Provider3);

        self.msSpreadsheetClient = new("https://graph.microsoft.com", {
            auth: {
                authHandler: oauth2Handler3
            },
            secureSocket: {
                trustStore: {
                            path: config:getAsString("KEYSTORE_PATH"),
                            password: config:getAsString("KEYSTORE_PASSWORD")
                }
            }
        });
    }

    # Function to create a worksheet on a given workbook
    # + workbookName - name of the workbook
    # + worksheetName - name of the worksheet to be created
    # + return - string with the identifier of the newly created worksheet. If not returns an error.
    public remote function createWorksheet(string workbookName, string worksheetName) returns @tainted (string|error) {
        string result = "";
        http:Request request = new;
        json payload = {"name" : worksheetName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.msSpreadsheetClient->post("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets", request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 201) {
                json|error response = httpResponse.getJsonPayload();
                if (response is map<json>) {
                    json nameItem = response["id"];
                    string createdTableID = nameItem.toString();
                    result = createdTableID;

                    return result;
                }
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }

    # Function to check whether a given table name exists in a workbook
    # Here, on which worksheet the table exists does not matter. If worksheetName corresponds to any worksheet in the workbook
    # and there is a table with the table name is located in the worksbook, this function returns true.
    # + workbookName - name of the workbook where table exists
    # + worksheetName - name of the worksheet where table exists
    # + tableName - name of the table to check the existence.
    # + return - boolean flag indicating whether table exists or not. If not returns an error.
    public remote function tableExists(string workbookName, string worksheetName, string tableName) returns boolean|error {
        boolean result = false;
        http:Request request = new;
        http:Response|error httpResponse = self.msSpreadsheetClient->get("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName + "/tables/" + tableName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 200) {
                result = true;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }

    # Function to create a new table
    # + workbookName - name of the workbook where the table should get created
    # + worksheetName - name of the worksheet where the table should get created
    # + tableName - name of the table to be created
    # + address - The location where the table should be created
    # + return - a flag indicating whether the table creation was successful or not. If not returns an error.
    public remote function createTable(string workbookName, string worksheetName, string tableName, string address) returns @tainted (boolean|error) {
        boolean result = false;
        http:Request request = new;
        json payload = {"name" : tableName, "address": address, "hasHeaders": false};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.msSpreadsheetClient->post("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName + "/tables/add", request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 201) {

                json|error response = httpResponse.getJsonPayload();
                if (response is map<json>) {
                    json nameItem = response["name"];
                    string createdTableName = nameItem.toString();
                    if (createdTableName != tableName) {
                        log:printInfo("Table created (" + createdTableName + ") carries different name than what was passed as the table name (" + tableName + "). Now patching the table with the correct table name.");
                        
                        boolean|error result2 = self->renameTable(workbookName, worksheetName, <@untainted > createdTableName, tableName);
                        if (result2 is boolean) {
                            result = result2;
                        }
                    }
                }

                result = true;
            } else if (httpResponse.statusCode == 400) {
                json|error response = httpResponse.getJsonPayload();
                if (response is map<json>) {
                    json errItem = response["error"];

                    if (errItem is map<json>) {
                        error err = error(WORKSHEET_ERROR_CODE, message = errItem["message"].toString());

                        return err;
                    }
                }
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }

    # Function to rename a table
    # + workbookName - name of the workbook where the table exists
    # + worksheetName - name of the worksheet where the table exists
    # + oldTableName - name of the table to be renamed
    # + newTableName - new name to be used with the table
    # + return - a flag indicating whether the table creation was successful or not. If not returns an error.
    public remote function renameTable(string workbookName, string worksheetName, string oldTableName, string newTableName) returns @tainted (boolean|error) {
        boolean result = false;
        http:Request request = new;
        json payload = {"name" : newTableName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.msSpreadsheetClient->patch("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName + "/tables/" + oldTableName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 200) {
                result = true;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }

    # Function for changing a table's header
    # + workbookName - name of the workbook where table exists
    # + worksheetName - name of the worksheet where table exists
    # + tableName - name of the table to be changed
    # + tableColumnID - ID of the tableColumn to change
    # + headerName - new name of the table header
    # + return - a flag indicating whether the table creation was successful or not. If not returns an error.
    public remote function setTableheader(string workbookName, string worksheetName, string tableName, int tableColumnID, string headerName) returns boolean|error {
        boolean result = false;
        http:Request request = new;
        json payload = {"name" : headerName};
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.msSpreadsheetClient->patch("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName + "/tables/" + tableName + "/columns/" + tableColumnID.toString(), request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 200) {
                result = true;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }

    # Function to insert data into a table
    # + workbookName - name of the workbook where table exists
    # + worksheetName - name of the worksheet where table exists
    # + tableName - name of the table where the data will be inserted
    # + data - data to be inserted into the table
    # + return - a flag indicating whether the table creation was successful or not. If not returns an error.
    public remote function insertDataIntoTable(string workbookName, string worksheetName, string tableName, json data) returns boolean|error {
        boolean result = false;
        http:Request request = new;
        json payload = data;
        request.setJsonPayload(payload);
        http:Response|error httpResponse = self.msSpreadsheetClient->post("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName + "/tables/" + tableName + "/rows/add", request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 200) {
                result = true;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

         return result;
    }

    # Function to delete a worksheet
    # + workbookName - name of the workbook where worksheet exists
    # + worksheetName - name of the worksheet to be deleted
    # + return - a flag indicating whether the worksheet deletion was successful or not. If not returns an error.
    public remote function deleteWorksheet(string workbookName, string worksheetName) returns boolean|error {
        boolean result = false;
        http:Request request = new;
        http:Response|error httpResponse = self.msSpreadsheetClient->delete("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 204) {
                result = true;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }

    # Function to check whether a worksheet exists or not
    # + workbookName - name of the workbook where the worksheet exists
    # + worksheetName - name of the worksheet
    # + return - a flag indicating whether the worksheet exists or not. If not returns an error.
    public remote function worksheetExists(string workbookName, string worksheetName) returns boolean|error {
        boolean result = false;
        http:Request request = new;
        http:Response|error httpResponse = self.msSpreadsheetClient->get("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets/" + worksheetName, request);

        if (httpResponse is http:Response) {
            if (httpResponse.statusCode == 200) {
                result = true;
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }

    # Function to get the list of worksheet names in a given workbook
    # + workbookName - name of the workbook from which the worksheet names list be be fecthed
    # + return - an arry of worksheet names. If not returns an error.
    public remote function getWorksheetNames(string workbookName) returns @tainted string[]|error {
        http:Request request = new;
        http:Response|error httpResponse = self.msSpreadsheetClient->get("/v1.0/me/drive/root:/" + workbookName  + ".xlsx:/workbook/worksheets", request);

        string[] result = [];

        if (httpResponse is http:Response) {
            var msg = httpResponse.getJsonPayload();

            if (msg is json) {
                if (msg is map<json>) {
                    if (msg["error"] == null) {
                        json[] arr = <json[]> msg["value"];
                        int i = 0;
                        foreach (json item in arr) {
                            map<json> member2 = <map<json>> item;
                            result[i] = member2["name"].toJsonString();
                            i += 1;
                            io:println("Name: " + member2["name"].toJsonString());
                        }
                    } else {
                        io:println("Error ocurred.");
                        io:println(msg.toJsonString());
                    }
                } else {
                    io:println("Not JSON");
                }
            } else {
                io:println("Invalid payload received:" , msg.reason());
            }
        } else {
            error err = error(WORKSHEET_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }

        return result;
    }
};

# Microsoft Graph client configuration.
# + baseUrl - The Microsoft Graph endpoint URL
# + bearerToken - Token for bearer authentication
# + clientConfig - OAuth2 direct token configuration
# + secureSocketConfig - HTTPS secure socket configuration
public type MicrosoftGraphConfiguration record {
    string baseUrl;
    string bearerToken;
    oauth2:DirectTokenConfig clientConfig;
    http:ClientSecureSocket secureSocketConfig?;
};