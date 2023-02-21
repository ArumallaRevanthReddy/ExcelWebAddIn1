'use strict';

function testResults(form) {
    let startyear = form.startyear.value;
    let endyear = form.endyear.value;
    let year = getYearString(startyear, endyear);

    Office.onReady(function () {
        $(document).ready(function () {
            getData(year, "AccountsPayableCurrent");
        });
    });
}

function getData(year, facts) {
    fetch('https://ew6mbmqfkj.execute-api.us-east-2.amazonaws.com/test/get-xbrl-facts', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({
            "CIK": 1001601,
            "year": year,
            "facts": facts
        })
    })
        .then(response => {
            if (!response.ok) {
                throw new Error('Network response was not ok');
            }
            return response.json();
        })
        .then(data => {
            console.log(data);
            populateData(data['AccountsPayableCurrent']);
        })
        .catch(error => {
            console.error('There was a problem with the API request:', error);
        });
}

async function populateData(response) {
    var data = prepareDataObject(response);

    Excel.run(function (ctx) {
        // make space for the data to be inserted
        var sheet1 = ctx.workbook.worksheets.getItem("Sheet1");
        var firstCell = sheet1.getCell(0, 0);
        var lastCell = sheet1.getCell(data.length - 1, data[0].length - 1);
        var range = firstCell.getBoundingRect(lastCell).insert('down');

        range.values = data; // insert data

        return ctx.sync();
    }).catch(function (error) {
        console.log(error);
    })
}

function prepareDataObject(response) {
    let data = [["val", "fy", "fp", "form", "filed", "end"]];
    if (response.val !== undefined) {
        let length = response.val.length;
        for (let i = 0; i < length; i++) {
            let row = [];
            row.push(response.val[i]);
            row.push(response.fy[i]);
            row.push(response.fp[i]);
            row.push(response.form[i]);
            row.push(response.filed[i]);
            row.push(response.end[i]);
            data.push(row);
        }
    }

    return data;
}

function getYearString(startyear, endyear) {
    let year = "";
    console.log(startyear, endyear)
    for (let i = startyear; i <= endyear; i++) {
        if (i !== startyear) year += ",";
        year += i;
    }

    return year;
}

(function () {

    Office.onReady(function () {
        // Office is ready
        $(document).ready(function () {
            // The document is ready
            $('#set-color').click(setColor);
            $('#get-data').click(getData);
        });
    });

    async function setColor() {
        await Excel.run(async (context) => {
            testResults();
            const range = context.workbook.getSelectedRange();
            range.format.fill.color = 'green';

            await context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    async function getData() {
        await Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange().load("values, rowCount, columnCount");
            range.values = range.values[0][0] * 100;

            await context.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
})();
