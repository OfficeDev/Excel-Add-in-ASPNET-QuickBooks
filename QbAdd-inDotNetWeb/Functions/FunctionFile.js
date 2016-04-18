/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

function addChart(event) {

    // Run a batch operation against the Excel object model
    Excel.run(function (ctx) {

        // Create a proxy object for the selected range and load its address and values properties
        var sourceRange = ctx.workbook.getSelectedRange().load("address");


        // Run the queued-up commands, and return a promise to indicate task completion
        return ctx.sync()
            .then(function () {
                var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                // Get the table
                var masterTable = ctx.workbook.tables.getItem("ExpensesTable");

                // Queue a command to add the new chart
                var chartDataRangeColumn1 = masterTable.columns.getItemAt(0).getDataBodyRange();
                var chartDataRangeColumn2 = masterTable.columns.getItemAt(1).getDataBodyRange();

                // Insert the chart in the sheet and format the chart
                var chartDataRange = chartDataRangeColumn1.getBoundingRect(chartDataRangeColumn2);
                var chart = sheet.charts.add("Line", chartDataRange, Excel.ChartSeriesBy.auto);
                chart.setPosition(sourceRange.address);
                chart.title.text = "Expense Trends";
                chart.title.format.font.color = "#41AEBD";
                chart.series.getItemAt(0).format.line.color = "#2E81AD";
            })
    });
}
