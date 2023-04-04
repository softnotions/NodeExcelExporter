
# NodeExcelExporter

Excel4node is a Node.js package that allows you to create and manipulate Excel spreadsheets using JavaScript. With this code, you can generate an Excel file with multiple sheets, each containing different data and formatting.

We have included a invoice-template.xlsx file for reference purposes.
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

#### server.js

The Express server is configured to use middleware such as express.json() and body-parser to parse incoming request data. The server also enables CORS using the cors() middleware. The dotenv package is used to load environment variables from a .env file.

The server listens on the specified port, which is the port defined in the environment variable. You can change the port from the .env file situated in the root folder.

The ExcelRoute is mounted on the /api route, which means that all requests starting with /api will be handled by the ExcelRoute. The route can be changed as per your requirement.

![Server Screenshot](https://imgtr.ee/images/2023/04/04/ko2Lx.png)

#### excel.controller.js

This is a custom controller that defines the function to generate the Excel file using the excel4node package. The controller exports an object with a generateExcel method that takes in the HTTP request and response objects as arguments.

Inside the generateExcel method, a new Excel workbook object is created using the workbook() method from the excel4node package. The workbook object is then used to create three new worksheets, each with a different name.

Data is added to the worksheets using various methods provided by excel4node. The data includes strings, numbers, and dates, and is formatted using various options such as font style, font size, font color, background color, and borders.

We have included detail instruction as comments in the controller file to understand each options provided by excel4node package.

![Controller Screenshot](https://imgtr.ee/images/2023/04/04/kowfJ.png)

#### excel.routes.js

This file provides a custom router object that defines a single route for generating an excel file. The router object is created using the express.Router() method.

The router object defines a single route for the GET HTTP method, which is mounted on the /excel route. When this route is accessed, the generateExcel function from the excel.controller.js file is executed. The generateExcel function generates an Excel file and sends it back as a response to the HTTP request.

![Route Screenshot](https://imgtr.ee/images/2023/04/04/koSki.png)

#### Assets Folder

We have included a logo file in the asset folder and added the logo into the excel file using the addImage option in the excel4node package. You can replace the logo file or add the neccessory images using the same function mentioned in the controller file.

#### Media Folder

The generated excel file will be saved into this folder with the name we have mentioned in the controller.
## Related

For more information, please have a look at excel4node package github repository.

[excel4node](https://github.com/advisr-io/excel4node)


## Authors

- [@softnotions](https://github.com/softnotions)

