// Declare the necessary variables. 
const Excel = require("exceljs"),
	fs = require("fs"),
	workbook = new Excel.Workbook();


// Check if the necessary excel file exists and if not send a notification informing the user. 
if(fs.existsSync("Bi-Weekly Timesheets.xlsx")) {
	// Read the excel file, save the current file, and clear the file for new entries. 
	workbook.xlsx.readFile("Bi-Weekly Timesheets.xlsx").then(() => {
		var copyindex = 0;
		// Iterate through all of the worksheets for each timesheet. 
		for(var j = 0; j < workbook.worksheets.length; j++) {
			// Extract the start and end dates. 
			var arr1 = JSON.stringify(workbook.worksheets[j].getCell("D2").value).split("-"),
				arr2 = JSON.stringify(workbook.worksheets[j].getCell("D3").value).split("-");
			arr1[0] = arr1[0].slice(1, arr1[0].length);	
			arr2[0] = arr2[0].slice(1, arr2[0].length);	
			arr1[2] = arr1[2].slice(0, 2);
			arr2[2] = arr2[2].slice(0, 2);
	    	// Set the formulas for the totals of hours, regular hours, and overtime hours. 
	    	workbook.worksheets[j].getCell("B14").value = { "formula": "SUM(D7:D13)" };
			workbook.worksheets[j].getCell("B23").value = { "formula": "SUM(D16:D22)" };
			workbook.worksheets[j].getCell("E14").value = { "formula": "IF(B14-\"40:00:00\">0,\"40:00:00\",B14)" };
			workbook.worksheets[j].getCell("E23").value = { "formula": "IF(B23-\"40:00:00\">0,\"40:00:00\",B23)" };
			workbook.worksheets[j].getCell("F14").value = { "formula": "IF(E14=\"40:00:00\",B14-\"40:00:00\",\"0:00:00\")" };
			workbook.worksheets[j].getCell("F23").value = { "formula": "IF(E23=\"40:00:00\",B23-\"40:00:00\",\"0:00:00\")" };
	    	// Check to see that the folder to which the current file is to be saved in fact exists. If not, create it and notify the user.
	    	for(; copyindex < 1; copyindex++) {
	    		var startdate = arr1[1].concat("-", arr1[2], "-", arr1[0]),
	    			enddate = arr2[1].concat("-", arr2[2], "-", arr2[0]);
	    		if(!fs.existsSync(arr1[0])) {
	    			fs.mkdirSync(arr1[0]);
					console.log("The folder for " + arr1[0] + " has been created. No action is needed!");
	    		}
	    	}
	    }
	    // Finish writing the current excel file to a safe folder. 
		return workbook.xlsx.writeFile(arr1[0].concat("\\", startdate,"_", enddate, ".xlsx"));
	}).then(() => {
		// Extract the start and end dates to determine the new ones. 
		var arr1 = JSON.stringify(workbook.worksheets[0].getCell("D2").value).split("-"),
			arr2 = JSON.stringify(workbook.worksheets[0].getCell("D3").value).split("-");
		arr1[0] = arr1[0].slice(1, arr1[0].length);	
		arr2[0] = arr2[0].slice(1, arr2[0].length);	
		arr1[2] = arr1[2].slice(0, 2);
		arr2[2] = arr2[2].slice(0, 2);
		var holder1 = new Date(parseInt(arr2[0]),parseInt(arr2[1]) - 1, parseInt(arr2[2])),
			holder2 = new Date(parseInt(arr2[0]),parseInt(arr2[1]) - 1, parseInt(arr2[2]));
		holder1.setDate(holder1.getDate() + 1);
		holder2.setDate(holder2.getDate() + 14);
		// Iterate through all of the worksheets for each timesheet. 
		for(var j = 0; j < workbook.worksheets.length; j++) {
	    	// Empty the previous entries on the timesheets. 
	    	for(var i = 7; i <= 22; i++) {
	    		if((7 <= i && i <= 13) || (16 <= i && i <= 22)) {
		    		workbook.worksheets[j].getCell("B".concat(i)).value = null;
		    		workbook.worksheets[j].getCell("C".concat(i)).value = null;
		    		workbook.worksheets[j].getCell("D".concat(i)).value = { "formula": "C".concat(i, "-", "B", i) };
		    		workbook.worksheets[j].getCell("E".concat(i)).value = null;
		    		workbook.worksheets[j].getCell("F".concat(i)).value = null;
		    	}
	    	}
	    	// Set the new start and end dates for the timesheets.
	    	workbook.worksheets[j].getCell("D2").value = holder1;
	    	workbook.worksheets[j].getCell("D3").value = holder2;
	    	// Set the formulas for the totals of hours, regular hours, and overtime hours. 
	    	workbook.worksheets[j].getCell("B14").value = { formula: "SUM(D7:D13)" };
			workbook.worksheets[j].getCell("B23").value = { formula: "SUM(D16:D22)" };
			workbook.worksheets[j].getCell("E14").value = { formula: "IF(B14-\"40:00:00\">0,\"40:00:00\",B14)" };
			workbook.worksheets[j].getCell("E23").value = { formula: "IF(B23-\"40:00:00\">0,\"40:00:00\",B23)" };
			workbook.worksheets[j].getCell("F14").value = { formula: "IF(E14=\"40:00:00\",B14-\"40:00:00\",\"0:00:00\")" };
			workbook.worksheets[j].getCell("F23").value = { formula: "IF(E23=\"40:00:00\",B23-\"40:00:00\",\"0:00:00\")" };
	    }
	    // Finish writing the new empty timesheets. 
		return workbook.xlsx.writeFile("Bi-Weekly Timesheets.xlsx");
	}).then(() => {
		console.log("The current timesheets have been saved and the new ones are ready to be edited. No action is needed.");
	});	
}
else {
	console.log("The 'Bi-Weekly Timesheets' excel file is missing!");
}