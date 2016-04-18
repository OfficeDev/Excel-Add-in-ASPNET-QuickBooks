/// <reference path="/Scripts/FabricUI/MessageBanner.js" />

/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function () {
    "use strict";

    var accessToken = {};
    var dlg;
    var messageBanner = null;

    Office.initialize = function (reason) {
        $(document).ready(function () {

            $('#btnSignIn').click(signIn);
            $('#btnSignOut').click(signOut);

            $('#btnGetExpenses').click(getExpenses);
            $('#btnCreateChart').click(addChart);

            $("#welcomePanel").show();
            $("#actionsPanel").hide();

            // Initialize the FabricUI notification mechanism and hide it
            messageBanner = new fabric.MessageBanner($(".ms-MessageBanner")[0]);
            messageBanner.hideBanner();

            checkSignIn();
        });
    };

    function signIn() {
        var signInUrl = "https://appcenter.intuit.com/Connect/SessionStart?datasources=quickbooks&grantUrl="
            + encodeURIComponent("https://localhost:44300/OAuthManager.aspx?connect=true");
        Office.context.ui.displayDialogAsync(signInUrl,
            { height: 70, width: 40},
            function (result) {
                dlg = result.value;
                dlg.addEventHandler("dialogMessageReceived", processMessage);
            });
    }

    function processMessage(arg) {

        //Work around for Mac:
        try {
            dlg.close();
            OSF.ClientHostController.closeDialog();
        } catch (e) { }

        accessToken = JSON.parse(arg.message);

        $.get("/api/setToken?" + $.param(accessToken))
            .done(function (data) {
                console.log(data);
            });

        $("#welcomePanel").hide();
        $("#actionsPanel").show();


    }

    function signOut() {
        $.get("/api/clearToken", function (data, status) {
            accessToken = null;
            $("#welcomePanel").show();
            $("#actionsPanel").hide();
        });
    }

    function checkSignIn() {
        if (typeof accessToken.token == "undefined") {
            $.get("api/getToken", function (data, status) {
                $("#welcomePanel").hide();
                $("#actionsPanel").fadeIn("slow");
            });
        }
    }


    function getExpenses() {
        var url = "/api/getExpenses?n=100";

        $.get(url, function (data, status) {
            var tableBody = getFormattedArray(data);
            addExpensesSheet(tableBody);
        });
    }


    // Import sample transactions into the workbook
    function addExpensesSheet(tableBody) {

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {

            // Add a new worksheet to store the transactions
            var dataSheet = ctx.workbook.worksheets.add("Expenses");

           // Fill white color in the sheet for improved look
            dataSheet.getRange("A2:M1000").format.fill.color = "white";

            // Add Sheet Title
            var range = dataSheet.getRange("B1:E1");
            range.values = "Contoso Expenses";
            range.format.font.name = "Corbel";
            range.format.font.size = 30;
            range.format.font.color = "white";
            range.merge();
            // Fill color in the brand bar
            dataSheet.getRange("A1:M1").format.fill.color = "#41AEBD";

            // Queue a command to add a new table
            var startRowNumber = 2;
            var masterTableAddress = 'Expenses!B' + startRowNumber + ':G' + (startRowNumber + tableBody.length).toString();
            var masterTable = ctx.workbook.tables.add(masterTableAddress, true);
            masterTable.name = "ExpensesTable";

            // Queue a command to get the newly added table
            masterTable.getHeaderRowRange().values = [["DATE", "AMOUNT", "MERCHANT", "CATEGORY", "TYPEOFDAY", "MONTH"]];

            masterTable.getDataBodyRange().formulas = tableBody;
            masterTable.columns.getItem("AMOUNT").getRange().numberFormat = '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)';

            // Format the table header and data rows
            range = dataSheet.getRange('Expenses!B' + startRowNumber + ':G' + startRowNumber);
            range.format.font.name = "Corbel";
            range.format.font.size = 10;
            range.format.font.bold = true;
            range.format.font.color = "black";

            range = dataSheet.getRange(masterTableAddress);
            range.format.font.name = "Corbel";
            range.format.font.size = 10;
            range.format.borders.getItem('EdgeBottom').style = 'Continuous';
            range.format.borders.getItem('EdgeTop').style = 'Continuous';

            // Sort by most recent transactions at the top (Date, descending order)
            var sortRange = masterTable.getDataBodyRange().getColumn(0).getUsedRange();
            sortRange.sort.apply([
            {
                key: 0,
                ascending: false,
            },
            ]);

            // Auto-fit columns and rows
            dataSheet.getUsedRange().getEntireColumn().format.autofitColumns();
            dataSheet.getUsedRange().getEntireRow().format.autofitRows();

            // Set the sheet as active
            dataSheet.activate();

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        }).catch(errorHandler);

    }

    function getFormattedArray(expenses) {
        var result = [];

        $.each(expenses, function (i, item) {
            //var tmpDate = new Date(item.txnDateField);
            //var date = tmpDate.getMonth()+1 + "/" + tmpDate.getDate() + "/" + tmpDate.getFullYear();
            var date = item.txnDateField.substr(0, 10);
            var type = "Unspecified";
            switch (item.paymentTypeField) {
                case 0:
                    type = "Cash";
                    break;
                case 1:
                    type = "Check";
                    break;
                case 2:
                    type = "Credit Card";
                    break;
            }
            var payee = "";
            if (item.entityRefField)
                payee = item.entityRefField.nameField;
            var cat = "";
            if (item.lineField.length > 0) {
                switch (item.lineField[0].detailTypeField) {
                    case 5:
                        cat = item.lineField[0].itemField.accountRefField.nameField;
                        break;
                    case "itemBasedExpenseLineDetailField":
                        cat = item.lineField[0].itemBasedExpenseLineDetailField.itemRefField.nameField;
                        break;

                }
            }
            var amount = item.totalAmtField;

            result.push([date, amount, payee, cat, '=IF(OR((TEXT([DATE], "dddd") = "Saturday"), (TEXT([DATE], "dddd") = "Sunday")), "Weekend", "Weekday")', '=TEXT([DATE], "mmm - yyyy")']);
        });

        return result;
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notificationHeader").text(header);
        $("#notificationBody").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
