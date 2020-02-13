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
import ballerina/stringutils;
import ballerina/config;
import ballerina/oauth2;

# Microsoft One Drive Client Object
public type OneDriveClient client object {
    http:Client oneDriveClient;

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

        self.oneDriveClient = new("https://graph.microsoft.com", {
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

    # Function to get the URL of an item located in Onedrive
    # + itemName - name of the item (e.g., Workbook) for which the URL has to be fetched
    # + return - string with the specific item's URL.
    public remote function getItemURL(string itemName) returns @tainted (string|error) {
        string result = "";
        http:Request request = new;
        http:Response|error httpResponse = self.oneDriveClient->get("/v1.0/me/drive/root/children", request);

        if (httpResponse is http:Response) {
            json|error payload = httpResponse.getJsonPayload();
            
            if (payload is map<json>) {
                json value = payload["value"];
                if (value is json[]) {
                    foreach var item in value {
                        if (item is map<json>){
                            if (stringutils:equalsIgnoreCase(item["name"].toString(), itemName)){
                                return item["webUrl"].toString();
                            }
                        }
                    }
                }
            }

            return result;
        } else {
             error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
             return err;
        }
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