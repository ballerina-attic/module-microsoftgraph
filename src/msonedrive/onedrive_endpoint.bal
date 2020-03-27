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
import ballerina/oauth2;

# Microsoft OneDrive Client Object.
# + oneDriveClient - HTTP client endpoint for accessing the OneDrive
public type OneDriveClient client object {
    http:Client oneDriveClient;

    public function __init(MicrosoftGraphConfiguration msGraphConfig) {
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

        self.oneDriveClient = new (msGraphConfig.baseUrl, {
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

    # Get an item located at the root level of the OneDrive.
    # + itemName - name of the item (e.g., Workbook) to be fetched
    # + return - item from the root on success, else returns an error
    public remote function getItemFromRoot(string itemName) returns @tainted (Item|error) {
        http:Request request = new;
        http:Response|error httpResponse = self.oneDriveClient->get("/v1.0/me/drive/root/children", request);
        Item resultItem = new();

        if (httpResponse is http:Response) {
            json|error payload = httpResponse.getJsonPayload();
            
            if (payload is map<json>) {
                json value = payload["value"];
                if (value is json[]) {
                    foreach var item in value {
                        if (item is map<json>){
                            if (stringutils:equalsIgnoreCase(item["name"].toString(), itemName)){
                                resultItem.id = item["id"].toString();
                                resultItem.name = item["name"].toString();
                                resultItem.webUrl = item["webUrl"].toString();
                                return resultItem;
                            }
                        } else {
                            error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while getting an item " +
                                                                             "from the root.");
                            return err;
                        }
                    }

                    return resultItem;
                } else {
                    error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while getting an item from the root.");
                    return err;    
                }
            } else {
                error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while getting an item from the root.");
                return err;
            }
        } else {
            error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }
    }

    # Get an item located at a non-root level location of the OneDrive.
    # + path - Path to the item (e.g., /foo/bar)
    # + itemName - name of the item (e.g., Workbook) to be fetched
    # + return - item from the non-root on success, else returns an error
    public remote function getItemFromNonRoot(string path, string itemName) returns @tainted (Item|error) {
        http:Request request = new;
        http:Response|error httpResponse = self.oneDriveClient->get("https://graph.microsoft.com/v1.0/me/drive/root:" +
                                            path + ":/children", request);
        Item resultItem = new();

        if (httpResponse is http:Response) {
            json|error payload = httpResponse.getJsonPayload();

            if (payload is map<json>) {
                json value = payload["value"];
                if (value is json[]) {
                    foreach var item in value {
                        if (item is map<json>){
                            if (stringutils:equalsIgnoreCase(item["name"].toString(), itemName)){
                                resultItem.id = item["id"].toString();
                                resultItem.name = item["name"].toString();
                                resultItem.webUrl = item["webUrl"].toString();
                                return resultItem;
                            }
                        } else {
                            error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while getting an item " +
                                                                             "from non-root.");
                            return err;
                        }
                    }

                    return resultItem;
                } else {
                    error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while getting an item " +
                                                                     "from non-root.");
                    return err;
                }
            } else {
                error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while getting an item from non-root.");
                return err;
            }
        } else {
            error err = error(ONEDRIVE_ERROR_CODE, message = "Error occurred while accessing the Microsoft Graph API");
            return err;
        }
    }
};

# Client Object which represents an item on Microsoft OneDrive.
# + id - unique identifier for the item
# + name - name of the item
# + webUrl - unique URL for accessing the item via a web browser
public type Item client object {
    public string id = "";
    public string name = "";
    public string webUrl = "";
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