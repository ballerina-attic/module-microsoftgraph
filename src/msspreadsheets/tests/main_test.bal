//
import ballerina/io;
import ballerina/test;
import ballerina/log;
import ballerina/config;
//
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
function testCreateSpreadsheet() {   
    MSSpreadsheetClient msGraphClient = new(msGraphConfig);
    
    boolean|error result = msGraphClient->deleteWorksheet("Book", "ABC");
    if (result is boolean) {
        io:println(result);
    } else {
        log:printError("Error deleting worksheet", err = result);
    }

    string|error resultStr = msGraphClient->createWorksheet("Book", "ABC");
    if (resultStr is string) {
        io:println(resultStr);
    } else {
        log:printError("Error creating worksheet", err = resultStr);
    }

    result = msGraphClient->createTable("Book", "ABC", "tableOpportunities", "A1:D1");
    if (result is boolean) {
        io:println(result);
    } else {
        log:printError("Error creating table", err = result);
    }
    result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 1, "MyColumn1");
    io:println(result);
    result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 2, "MyColumn2");
    io:println(result);
    result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 3, "MyColumn3");
    io:println(result);
    result = msGraphClient->setTableheader("Book", "ABC", "tableOpportunities", 4, "MyColumn4");
    io:println(result);
    json data = {"values": [[1, 2, 3, 4]]};
    result = msGraphClient->insertDataIntoTable("Book", "ABC", "tableOpportunities", data);
    io:println(result);
}