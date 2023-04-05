# NodeExcelExporter

Excel4node is a Node.js package that allows you to create and manipulate Excel spreadsheets using JavaScript. With this code, you can generate an Excel file with multiple sheets, each containing different data and formatting. We have included a invoice-template.xlsx file for reference purposes.

## Project Initialization

To get started, clone the repository to your local machine using the following command:

```bash
  git clone github_url
```

Go to the project directory

```bash
  cd excel_node
```

#### Install dependencies

Once you have cloned the repository, navigate to the project directory in your terminal and install the project dependencies using the following command:

```bash
npm install
```

#### Start the application

After installing the dependencies, you can start the application using the following command:

```bash
npm start
```

## File Structure

When you open the code in an IDE, there will be a few folders and files.

### server.js

The Express server is configured to use middleware such as express.json() and body-parser to parse incoming request data. The server also enables CORS using the cors() middleware. The dotenv package is used to load environment variables from a .env file.

The server listens on the specified port, which is the port defined in the environment variable. You can change the port from the .env file situated in the root folder.

The ExcelRoute is mounted on the /api route, which means that all requests starting with /api will be handled by the ExcelRoute. The route can be changed as per your requirement.

```javascript
// Import the excel4node package as excel
import excel from 'excel4node';

// Import the fs (file system) library as fs
import fs from 'fs';

// Import path as path. Path is a built in node js module
import path from 'path';

// Import the data from JSON file or you can get data from databse
import data from '../data/WorksheetSheetThree.json' assert { type: 'json' };

// Declare an asynchronous function in any name. eg: generateExcel
export const generateExcel = async (req, res) => {
    try {
        // Initialize a workbook using excel4node Workbook
        let workbook = new excel.WorkBook();

        // Create work sheets as per your need using addWorksheet function
        let worksheet = workbook.addWorksheet("Sheet 1");

        // create different styles as per your need using createStyle function
        const style = workbook.createStyle({
            // Add background color to the cell using fill property
            fill: {
                type: "pattern",
				patternType: "solid",
				bgColor: "#086AD8",
				fgColor: "#086AD8",
            },
            // Add border color and thickness using border property
            border: {
                top: {
					style: "thin",
					color: "#FFFFFF",
				},
				bottom: {
					style: "thin",
					color: "#FFFFFF",
				},
				left: {
					style: "thin",
					color: "#FFFFFF",
				},
				right: {
					style: "thin",
					color: "#FFFFFF",
				}
			},
            // Adjust alignment of the content using alignment property
            // Adjust the text wrapping using the wrapText property
            alignment: {
				wrapText: true,
				vertical: "center",
                horizontal: "center"
			},
            // Add font color, size, bold and font family(name) using font property
            // All global font families can be used
            font: {
				color: "#000000",
				size: 13,
				name: "Roboto",
			},
        });

        // Add the style to to cells using the cell numbers
        worksheet.cell(1, 1).style(style);

        // The cells can be merged using merge property
        // The below code will merge cells on first row first cell to 47th cell
        worksheet.cell(1, 1, 1, 47, true);

        // You can add height each rows using height property
        // The below example will set the first row's height to 30pts
        worksheet.row(1).setHeight(30);

        // You can set width to each columns using width property
        // The below example will set the first column' width to 100pts
        worksheet.column(1).setWidth(100);

        // If you want to loop the data from databse or JSON file use this foreach method
        data.forEach((item) => {
            worksheet.cell(1, 1).string(item.order_date);
            worksheet.cell(1, 2).string(item.order_no);
        })

        // Check below code to generate the xlsx file and save to any folder
        let fileName = "invoice-template.xlsx"; // declared the filename

        // Define the absolute file path for the file to be generated
        // The below code will save the generated file to the media folder in the defined fileName
        let filePath = path.resolve(process.cwd(), "media", fileName);

        // Rename the file using fs package to replace the same file when generating multipl times
        fs.rename(filePath, () => {
            console.log("File Renamed!");
        });

        // Write and generate the excel using write property
        workbook.write(filePath);

        // Send the response to the client indicating that the excel file has been generated
        res.status(200).send("Excel file generated!");
    } catch (error) {
        console.log(error);
    }
}

```



### excel.controller.js

This is a custom controller that defines the function to generate the Excel file using the excel4node package. The controller exports an object with a generateExcel method that takes in the HTTP request and response objects as arguments.

Inside the generateExcel method, a new Excel workbook object is created using the workbook() method from the excel4node package. The workbook object is then used to create three new worksheets, each with a different name.

Data is added to the worksheets using various methods provided by excel4node. The data includes strings, numbers, and dates, and is formatted using various options such as font style, font size, font color, background color, and borders.

```javascript
// Import the excel4node package as excel
import excel from 'excel4node';

// Import the fs (file system) library as fs
import fs from 'fs';

// Import path as path. Path is a built in node js module
import path from 'path';

// Declare an asynchronous function in any name. eg: generateExcel
export const generateExcel = async (req, res) => {
    try {
        // Initialize a workbook using excel4node Workbook
        let workbook = new excel.WorkBook();

        // Create work sheets as per your need using addWorksheet function
        let worksheet = workbook.addWorksheet("Sheet 1");

        // create different styles as per your need using createStyle function
        const style = workbook.createStyle({
            // Add background color to the cell using fill property
            fill: {
                type: "pattern",
				patternType: "solid",
				bgColor: "#086AD8",
				fgColor: "#086AD8",
            },
            // Add border color and thickness using border property
            border: {
                top: {
					style: "thin",
					color: "#FFFFFF",
				},
				bottom: {
					style: "thin",
					color: "#FFFFFF",
				},
				left: {
					style: "thin",
					color: "#FFFFFF",
				},
				right: {
					style: "thin",
					color: "#FFFFFF",
				}
			},
            // Adjust alignment of the content using alignment property
            // Adjust the text wrapping using the wrapText property
            alignment: {
				wrapText: true,
				vertical: "center",
                horizontal: "center"
			},
            // Add font color, size, bold and font family(name) using font property
            // All global font families can be used
            font: {
				color: "#000000",
				size: 13,
				name: "Roboto",
			},
        });

        // Add the style to to cells using the cell numbers
        worksheet.cell(1, 1).style(style);

        // The cells can be merged using merge property
        // The below code will merge cells on first row first cell to 47th cell
        worksheet.cell(1, 1, 1, 47, true);

        // You can add height each rows using height property
        // The below example will set the first row's height to 30pts
        worksheet.row(1).setHeight(30);

        // You can set width to each columns using width property
        // The below example will set the first column' width to 100pts
        worksheet.column(1).setWidth(100);

        // Check below code to generate the xlsx file and save to any folder
        let fileName = "invoice-template.xlsx"; // declared the filename

        // Define the absolute file path for the file to be generated
        // The below code will save the generated file to the media folder in the defined fileName
        let filePath = path.resolve(process.cwd(), "media", fileName);

        // Rename the file using fs package to replace the same file when generating multipl times
        fs.rename(filePath, () => {
            console.log("File Renamed!");
        });

        // Write and generate the excel using write property
        workbook.write(filePath);

        // Send the response to the client indicating that the excel file has been generated
        res.status(200).send("Excel file generated!");
    } catch (error) {
        console.log(error);
    }
}
```

### excel.routes.js

This file provides a custom router object that defines a single route for generating an excel file. The router object is created using the express.Router() method.

The router object defines a single route for the GET HTTP method, which is mounted on the /excel route. When this route is accessed, the generateExcel function from the excel.controller.js file is executed. The generateExcel function generates an Excel file and sends it back as a response to the HTTP request.

```javascript
// Import the necessary modules
import express from "express"; // Import the Express web framework
import { generateExcel } from "../controllers/excel.controller.js"; // Import the generateExcel controller function

// Create a new router instance with the Express Router() method
const router = express.Router();

// Define a GET route for "/excel" that calls the generateExcel controller function
router.get("/excel", generateExcel);

// Export the router instance for use in other modules
export default router;
```

### Assets Folder

We have included a logo file in the asset folder and added the logo into the excel file using the addImage option in the excel4node package. You can replace the logo file or add the neccessory images using the same function mentioned in the controller file.

### Media Folder

The generated excel file will be saved into this folder with the name we have mentioned in the controller.

## Related

For more information, please have a look at excel4node package github repository.

[excel4node](https://github.com/advisr-io/excel4node)

## Authors

- [@softnotions](https://github.com/softnotions)