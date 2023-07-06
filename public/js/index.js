// ----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.
// ----------------------------------------------------------------------------

let models = window["powerbi-client"].models;
let reportContainer = $("#report-container")[0];
let questionText;
let previousValue;
var timeoutId;

questionText = $("#txtQuestion").val();
$('#txtQuestion').on('change keyup', function() {
    console.log('Textbox Change');
   
    clearTimeout(timeoutId);
    timeoutId = setTimeout(function() {  
        questionText = $("#txtQuestion").val();
        if (previousValue !== questionText) {
            $("#loadingTxtId")[0].style.display = "block"
            previousValue = JSON.parse(JSON.stringify(questionText));
            configureAgainAndAgain();
        }
    }, 1000);
});
// 

// Initialize iframe for embedding report
powerbi.bootstrap(reportContainer, { type: "report" });

// AJAX request to get the report details from the API and pass it to the UI
function configureAgainAndAgain() {

    $.ajax({
        type: "GET",
        url: "/getEmbedToken",
        dataType: "json",
        success: function (embedData) {
        console.log(embedData.accessToken);
            // Create a config object with type of the object, Embed details and Token Type
            let reportLoadConfig = {
                type: "report",
                tokenType: models.TokenType.Embed,
                accessToken: embedData.accessToken,
                settings: models.IEmbedSettings,
                // Use other embed report config based on the requirement. We have used the first one for demo purpose
                embedUrl: embedData.embedUrl[0].embedUrl,
            };
            txtEmbedUrl = "https://app.powerbi.com/qnaEmbed?groupId=273777b9-50ca-4e2c-ab7b-af9fcd4d1331";
           
            var qnaconfig = { 
                type: 'qna', 
                tokenType: models.TokenType.Embed, 
                accessToken: embedData.accessToken, 
                embedUrl: txtEmbedUrl, 
                datasetIds: ["f2d37119-970c-46fc-8ec3-12706708604e"], 
                viewMode: models.QnaMode.Interactive,
                question: questionText 
            }; 
        
            // Use the token expiry to regenerate Embed token for seamless end user experience
            // Refer https://aka.ms/RefreshEmbedToken
            tokenExpiry = embedData.expiry;

            // Embed Power BI report when Access token and Embed URL are available
            var report = powerbi.embed(reportContainer, qnaconfig);

            // Clear any other loaded handler events
            //report.off("loaded");
            report.on("visualRendered", async function () {
                console.log("Report visualRendered successful");
                $("#loadingTxtId")[0].style.display = "none"
            });
        
            // Triggers when a report schema is successfully loaded
            report.on("loaded", async function () {
                console.log("Report load successful");                
            });

            // Clear any other rendered handler events
            report.off("rendered");
            let pages;
            // Triggers when a report is successfully embedded in UI
            report.on("rendered", async function () {
                console.log("Report render successful");
                report.switchMode("edit");
             });

            // Clear any other error handler events
            report.off("error");

            // Handle embed errors
            report.on("error", function (event) {
                let errorMsg = event.detail;
                console.error(errorMsg);
                return;
            });
        },
    
        
        error: function (err) {
            $("#loadingTxtId").style.display = 'none';
            // Show error container
            let errorContainer = $(".error-container");
            $(".embed-container").hide();
            errorContainer.show();

            // Get the error message from err object
            let errMsg = JSON.parse(err.responseText)['error'];

            // Split the message with \r\n delimiter to get the errors from the error message
            let errorLines = errMsg.split("\r\n");

            // Create error header
            let errHeader = document.createElement("p");
            let strong = document.createElement("strong");
            let node = document.createTextNode("Error Details:");

            // Get the error container
            let errContainer = errorContainer.get(0);

            // Add the error header in the container
            strong.appendChild(node);
            errHeader.appendChild(strong);
            errContainer.appendChild(errHeader);

            // Create <p> as per the length of the array and append them to the container
            errorLines.forEach(element => {
                let errorContent = document.createElement("p");
                let node = document.createTextNode(element);
                errorContent.appendChild(node);
                errContainer.appendChild(errorContent);
            });
        }
    });
}