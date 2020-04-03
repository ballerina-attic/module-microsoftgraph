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
# + oneDriveClient - HTTP client endpoint for accessing OneDrive
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

    # Get an item located at the root level of OneDrive.
    # + itemName - name of the item (e.g., Workbook) to be fetched
    # + return - item from the root if fetching is successful or else returns an error
    public remote function getItemFromRoot(string itemName) returns @tainted (Item|error) {
        //Make a GET request and collect the information about the items on the root
        http:Request request = new;
        http:Response|error response = self.oneDriveClient->get("/v1.0/me/drive/root/children", request);
        Item resultItem = new();

        if response is error {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
                                        errorCode = HTTP_ERROR, cause = response);
            return httpError;
        }

        http:Response httpResponse = <http:Response> response;

        //If the request was successful it will return the details in a JSON response
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfResponse = typeof responseJson;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                message = "Invalid response; expected a `map<json>` found " +  typeOfResponse.toString(), 
                errorCode = TYPE_CONVERSION_ERROR);

            return typeError;
        }

        map<json> responsePayload = <map<json>> responseJson;

        json|error value = responsePayload.value;

        if !(value is json[]) {
            typedesc<any|error> typeOfValue = typeof value;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                message = "Invalid value; expected a `json[]` found " +  typeOfValue.toString(), 
                errorCode = TYPE_CONVERSION_ERROR);

            return typeError;
        }

        json[] itemsArray = <json[]> value;

        //Iterate through the array of items until the specified item was found
        foreach var item in itemsArray {
            if (item is map<json>){
                if (stringutils:equalsIgnoreCase(item["name"].toString(), itemName)){
                    resultItem.id = item["id"].toString();
                    resultItem.name = item["name"].toString();
                    resultItem.webUrl = item["webUrl"].toString();
                    return resultItem;
                }
            } else {
                typedesc<any|error> typeOfResponse = typeof responseJson;
                TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid response; expected a `map<json>` found " +  typeOfResponse.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);

                return typeError;
            }
        }

        return resultItem;
    }

    # Get an item located at a non-root level location of OneDrive.
    # + path - path to the item (e.g., /foo/bar)
    # + itemName - name of the item (e.g., Workbook) to be fetched
    # + return - item from the non-root if fetching is successful or else returns an error
    public remote function getItemFromNonRoot(string path, string itemName) returns @tainted (Item|error) {
        //Make a GET request and collect the information about the items from the non-root location
        http:Request request = new;
        http:Response|error response = self.oneDriveClient->get("https://graph.microsoft.com/v1.0/me/drive/root:" +
                                            path + ":/children", request);
        Item resultItem = new();

        if response is error {
            HttpError httpError = error(HTTP_ERROR, message = "Error occurred while accessing the Microsoft Graph API.", 
                errorCode = HTTP_ERROR, cause = response);
            return httpError;
        }

        http:Response httpResponse = <http:Response> response;

        //If the request was successful it will return the details in a JSON response
        json|error responseJson = httpResponse.getJsonPayload();

        if !(responseJson is map<json>) {
            typedesc<any|error> typeOfResponse = typeof responseJson;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                message = "Invalid response; expected a `map<json>` found " +  typeOfResponse.toString(), 
                errorCode = TYPE_CONVERSION_ERROR);

            return typeError;
        }

        map<json> responsePayload = <map<json>> responseJson;

        json|error value = responsePayload.value;

        if !(value is json[]) {
            typedesc<any|error> typeOfValue = typeof value;
            TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                message = "Invalid value; expected a `json[]` found " +  typeOfValue.toString(), 
                errorCode = TYPE_CONVERSION_ERROR);

            return typeError;
        }

        json[] itemsArray = <json[]> value;

        //Iterate through the array of items until the specified item was found
        foreach var item in itemsArray {
            if (item is map<json>){
                if (stringutils:equalsIgnoreCase(item["name"].toString(), itemName)){
                    resultItem.id = item["id"].toString();
                    resultItem.name = item["name"].toString();
                    resultItem.webUrl = item["webUrl"].toString();
                    return resultItem;
                }
            } else {
                typedesc<any|error> typeOfItem = typeof item;
                TypeConversionError typeError = error(TYPE_CONVERSION_ERROR, 
                    message = "Invalid response; expected a `map<json>` found " +  typeOfItem.toString(), 
                    errorCode = TYPE_CONVERSION_ERROR);

                return typeError;
            }
        }

        return resultItem;
    }
};

# Client Object, which represents an item on Microsoft OneDrive.
# + id - unique identifier for the item
# + name - name of the item
# + webUrl - unique URL for accessing the item via a web browser
public type Item client object {
    public string id = "";
    public string name = "";
    public string webUrl = "";
};

# Microsoft Graph client configuration.
# + baseUrl - the Microsoft Graph endpoint URL
# + msInitialAccessToken - initial access token
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