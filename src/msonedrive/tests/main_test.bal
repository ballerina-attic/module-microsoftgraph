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
import ballerina/io;
import ballerina/test;
import ballerina/log;
import ballerina/config;
//
//// Create Microsoft Graph Client configuration by reading from config file.
MicrosoftGraphConfiguration msGraphConfig = {
    baseUrl: config:getAsString("MS_EP_URL"),
    bearerToken: config:getAsString("BEARER_TOKEN"),

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

@test:Config {}
function testGetURLofItem() {
    OneDriveClient oneDriveClient = new(msGraphConfig);

    string|error result = oneDriveClient->getItemURL("Book.xlsx");
    if (result is string) {
        io:println(result);
    } else {
        log:printError("Error getting the WorkBook URL", err = result);
    }
}
