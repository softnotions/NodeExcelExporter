// Imported the excel4node package as excel.
import excel from "excel4node";

// Imported the fs (file system) library as fs.
// A built-in Node.js module for interacting with the file system.
import fs from "fs";

// A  built-in Node.js module for working with file and directory paths in a platform-independent way.
import path from "path";

import process from "process";

// Declared an asynchronous function named `generateExcel`.
export const generateExcel = async (req, res) => {
	try {
		
		// Created a new Excel workbook using the excel4node library
		let workbook = new excel.Workbook();

		// Added three new worksheets to the workbook with the given names.
		let worksheet1 = workbook.addWorksheet("Summary");
		let worksheet2 = workbook.addWorksheet("Payout Invoice");
		let worksheet3 = workbook.addWorksheet("All Orders");

		// Set the height of certain rows in the first worksheet.
		worksheet1.row(1).setHeight(15);
		worksheet1.row(7).setHeight(30);
		worksheet1.row(9).setHeight(20);
		worksheet1.row(17).setHeight(20);
		worksheet1.row(20).setHeight(20);
		worksheet1.row(26).setHeight(20);
		worksheet1.row(44).setHeight(30);
		worksheet1.row(38).setHeight(30);
		worksheet1.row(32).setHeight(30);
		worksheet1.row(41).setHeight(30);

		// Set the width of certain columns in the first worksheet
		worksheet1.column(1).setWidth(10);
		worksheet1.column(2).setWidth(40);
		worksheet1.column(3).setWidth(50);

		// Defined different styles for cells in different sheets.
		const titleStyle = workbook.createStyle({
			// Defined the bckground color for the cells
			fill: {
				type: "pattern",
				patternType: "solid",
				bgColor: "#086AD8",
				fgColor: "#086AD8",
			},
			// Defined the border thickness and color of cells
			border: {
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
				},
				top: {
					style: "thin",
					color: "#FFFFFF",
				}
			},

		});

		const style1 = workbook.createStyle({
			font: {
				color: "#000000",
				size: 30,
				name: "Arial",
				bold: true,
			},
			border: {
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
				},
				top: {
					style: "thin",
					color: "#FFFFFF",
				}
			},

		});

		const style2 = workbook.createStyle({

			border: {
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
				},
				top: {
					style: "thin",
					color: "#FFFFFF",
				}
			},
			// Defined the horizontal and vertical alignment of the text.
			// Also the wrapText to show all content inside the cell.
			alignment: {
				wrapText: true,
				vertical: "center",
			},
			// Defined the font color, size and font family name
			font: {
				color: "#000000",
				size: 13,
				name: "Arial",

			},
		});

		const style3 = workbook.createStyle({
			font: {
				color: "#000000",
				size: 15,
				name: "Arial",

			},
			alignment: {
				wrapText: true,
				vertical: "center"
			},
		});

		const style4 = workbook.createStyle({
			font: {
				bold: true,
				underline: false,
				color: "#000000",
				name: "Arial",

			},
			alignment: {
				wrapText: true,
				vertical: "center",
				horizontal:"center"
			},
			border: {
				bottom: {
					style: "thin",
					color: "#DEDEDE",
				},
				left: {
					style: "thin",
					color: "#DEDEDE",
				},
				right: {
					style: "thin",
					color: "#DEDEDE",
				},
				top: {
					style: "thin",
					color: "#DEDEDE",
				}
			},
		});

		const style5 = workbook.createStyle({
			alignment: {
				wrapText: true,
				vertical: "center",
				horizontal:"right"

			},
			font: {
				color: "#086AD8",
				bold: true,
				size: 15,
				name: "Arial",

			},
			border: {
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
				},
				top: {
					style: "thin",
					color: "#FFFFFF",
				}
			},
		});

		const style6 = workbook.createStyle({
			alignment: {
				wrapText: true,
				vertical: "center",
				horizontal:"left"
			},
			border: {
				bottom: {
					style: "thin",
					color: "#DEDEDE",
				},
				left: {
					style: "thin",
					color: "#DEDEDE",
				},
				right: {
					style: "thin",
					color: "#DEDEDE",
				},
				top: {
					style: "thin",
					color: "#DEDEDE",
				}
			},
		});

		const style7 = workbook.createStyle({
			font: {
				color: "#1D1D1D",
				bold: false,
				size: 11,
				name:"Arial"
			},
			alignment: {
				wrapText: true,
				vertical: "center",
				horizontal:"center"
			},
		});

		const style8 = workbook.createStyle({
			font: {
				color: "#000000",
				bold: true,
				size: 18,
				name: "Arial",

			},
			border: {
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
				},
				top: {
					style: "thin",
					color: "#FFFFFF",
				}
			},
		});

		const style9 = workbook.createStyle({
			font: {
				color: "#000000",
				bold: true,
				size: 15,
				name: "Arial",
			},
			fill: {
				type: "pattern",
				patternType: "solid",
				bgColor: "#DEDEDE",
				fgColor: "#DEDEDE"
			},

		});

		const style10 = workbook.createStyle({
			font: {
				color: "#000000",
				size: 12,
				name: "Arial",

			},
			alignment: {
				wrapText: true,
				vertical: "center"
			},
			border: {
				bottom: {
					style: "thin",
					color: "#DEDEDE",
				},
				left: {
					style: "thin",
					color: "#DEDEDE",
				},
				right: {
					style: "thin",
					color: "#DEDEDE",
				},
				top: {
					style: "thin",
					color: "#DEDEDE",
				}
			},
		});

		const style11 = workbook.createStyle({
			border: {
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
				},
				top: {
					style: "thin",
					color: "#FFFFFF",
				}
			},
			alignment: {
				wrapText: true,
				vertical: "center",
			},
			font: {
				color: "#6A6A6A",
				size: 12,
				name: "Arial",
			},
		});

		const style12 = workbook.createStyle({
			alignment: {
				wrapText: true,
				vertical: "center",
				horizontal:"center"
			},
			fill: {
				type: "pattern",
				patternType: "solid",
				bgColor: "#D4E8FF",
				fgColor: "#D4E8FF"
			},
			border: {
				bottom: {
					style: "thin",
					color: "#DEDEDE",
				},
				left: {
					style: "thin",
					color: "#DEDEDE",
				},
				right: {
					style: "thin",
					color: "#DEDEDE",
				},
				top: {
					style: "thin",
					color: "#DEDEDE",
				}
			},
			font: {
				color: "#000000",
				size: 11,
				name: "Arial",
			},
		});

		const style13 = workbook.createStyle({
			alignment: {
				wrapText: true,
				vertical: "center",
				horizontal:"right"
			},
			border: {
				bottom: {
					style: "thin",
					color: "#DEDEDE",
				},
				left: {
					style: "thin",
					color: "#DEDEDE",
				},
				right: {
					style: "thin",
					color: "#DEDEDE",
				},
				top: {
					style: "thin",
					color: "#DEDEDE",
				}
			},
		});

		// Defined this style to remove border of all unused cells.
		// There is no way to remove the border, instead we have added a white color for border
		const globalBorder = workbook.createStyle({
			border: {
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
				},
				top: {
					style: "thin",
					color: "#FFFFFF",
				}
			},
		});

		// This code is to merge multiple cells and add common style for the cells.
		worksheet1.cell(1, 1, 1, 15, true).style(titleStyle);
		worksheet1.cell(1, 1, 100, 1, true).style(style2);
		worksheet1.cell(1, 2, 100, 2, true).style(style2);
		worksheet1.cell(1, 3, 100, 3, true).style(style2);
		worksheet1.cell(1, 4, 100, 4, true).style(style2);
		worksheet1.cell(1, 5, 100, 5, true).style(style2);
		worksheet1.cell(1, 6, 100, 6, true).style(style2);
		worksheet1.cell(1, 7, 100, 7, true).style(style2);
		worksheet1.cell(1, 8, 100, 8, true).style(style2);
		worksheet1.cell(1, 9, 100, 9, true).style(style2);
		worksheet1.cell(1, 10, 100, 10, true).style(style2);
		worksheet1.cell(1, 11, 100, 11, true).style(style2);
		worksheet1.cell(1, 12, 100, 12, true).style(style2);
		worksheet1.cell(1, 13, 100, 13, true).style(style2);
		worksheet1.cell(1, 14, 100, 14, true).style(style2);
		worksheet1.cell(1, 15, 100, 15, true).style(style2);
		worksheet1.cell(2, 2, 4, 10, true).style(style1);
		worksheet1.cell(5, 2, 5, 10, true).style(style2);
		worksheet1.cell(6, 2, 6, 10, true).style(style2);
		worksheet1.cell(8, 2).string("Your Payout").style(style1);
		worksheet1.cell(9, 2, 9, 10, true).style(style2);

		// Defined the contents for each cells in worksheet 1. This will be common in all worksheets.
		worksheet1.cell(10, 2).string("SOFTNOTIONS").style(style3);
		worksheet1.cell(11, 2).string("Module No. B,6th Floor,\nBhavani Building").style(style2);
		worksheet1.cell(12, 2).string("Phase-I Campus,\nTechnopark,Trivandrum").style(style2);
		worksheet1.cell(13, 2).string("Karyavattom P.O, Pin - 695 581").style(style2);
		worksheet1.cell(14, 2, 15, 10, true).style(style2);
		worksheet1.cell(16, 2).string("Payout Period").style(style11);
		worksheet1.cell(16, 3).string("Payout On").style(style11);
		worksheet1.cell(17, 2).string("26 March - 20 April").style(style8);
		worksheet1.cell(17, 3).string("21 April").style(style8);
		worksheet1.cell(21, 2).string("Total Payout").style(style11);
		worksheet1.cell(21, 3).string("Total Orders").style(style11);
		worksheet1.cell(22, 2).string("5020.5").style(style8);
		worksheet1.cell(22, 3).string("35").style(style8);
		worksheet1.cell(27, 2,27,3,true).string("How to use this Annexure").style(style8);
		worksheet1.cell(28, 2,28,3,true).string("Let's walk you through the tabs listed at the bottom of this sheet.").style(style11);
		worksheet1.cell(31, 2,31,3,true).string("Payout Invoice").style(style9);
		worksheet1.cell(32, 2,32,3,true).string("The payout summary gives you an overall breakup of your payout with specific breakup of your earnings, fees and deductions.").style(style11);
		worksheet1.cell(34,2,34, 3,true).string("All Orders").style(style9);
		worksheet1.cell(35,2,35, 3,true).string("You can see the the breakup of your earnings at an order level.").style(style2);
		worksheet1.cell(37,2,37, 3,true).string("Discounts P&L").style(style9);
		worksheet1.cell(38,2,38, 3,true).string("This is where you can look at the financial breakup of the discounts which were used by your customers in this payout period. ").style(style11);
		worksheet1.cell(40,2,40, 3,true).string("Adjustment Details").style(style9);
		worksheet1.cell(41,2,41, 3,true).string("This is where you can see the breakup & details about the adjustments made from your weekly payout").style(style11);
		worksheet1.cell(43,2,43, 3,true).string("FAQ/Glossary").style(style9);
		worksheet1.cell(44,2,44, 3,true).string("If you have any queries on the payout, you can look at the most frequently asked questions. Also contains a glossary of the terms used in this annexure.").style(style11);

		// Inserted the logo from assets using the fs dependancy. The image file path can be changed.
		const imgFilePath = "./assets/logo.png";
		// const imgData = fs.readFileSync(imgFilePath);
		const pic = worksheet1.addImage({
			path: imgFilePath,
			type: "picture",
			position: {
				type: "twoCellAnchor",
				from: {
					col: 1,
					colOff: "30mm",
					row: 2,
					rowOff: "5mm"
				},
				to: {
					col: 3,
					colOff: "10mm",
					row: 7,
					rowOff: "5mm"
				}
			}
		});
		pic.editAs = "twoCell";

		//Payout Invoice Sheet 

		// Defined the width of columns in worksheet 2
		worksheet2.column(1).setWidth(10);
		worksheet2.column(2).setWidth(30);
		worksheet2.column(3).setWidth(30);
		worksheet2.column(4).setWidth(30);
		worksheet2.column(5).setWidth(30);
		worksheet2.column(6).setWidth(500);

		// Defined the height of rows in worksheet 2
		worksheet2.row(4).setHeight(50);
		worksheet2.row(6).setHeight(20);
		worksheet2.row(8).setHeight(20);
		worksheet2.row(9).setHeight(20);
		worksheet2.row(10).setHeight(30);
		worksheet2.row(11).setHeight(20);
		worksheet2.row(12).setHeight(20);
		worksheet2.row(13).setHeight(20);
		worksheet2.row(14).setHeight(30);
		worksheet2.row(15).setHeight(20);
		worksheet2.row(16).setHeight(20);
		worksheet2.row(17).setHeight(20);
		worksheet2.row(18).setHeight(20);
		worksheet2.row(19).setHeight(20);
		worksheet2.row(20).setHeight(20);
		worksheet2.row(21).setHeight(20);
		worksheet2.row(22).setHeight(20);
		worksheet2.row(23).setHeight(30);
		worksheet2.row(24).setHeight(30);
		worksheet2.row(25).setHeight(30);
		worksheet2.row(26).setHeight(800);
		worksheet2.row(27).setHeight(20);
		worksheet2.row(28).setHeight(20);
		worksheet2.row(29).setHeight(20);
		worksheet2.row(30).setHeight(20);
		worksheet2.row(31).setHeight(20);
		worksheet2.row(32).setHeight(20);
		worksheet2.row(33).setHeight(20);
		worksheet2.row(34).setHeight(20);
		worksheet2.row(35).setHeight(20);
		worksheet2.row(36).setHeight(20);

		// Merged multiple cells together and added common style.
		worksheet2.cell(1, 1, 1, 10, true).style(titleStyle);
		worksheet2.cell(2, 6, 20, 6, true);
		worksheet2.cell(1, 1, 20, 1, true).style(style2);
		worksheet2.cell(2, 1, 2, 5, true).style(style2);
		worksheet2.cell(3, 1, 3, 5, true).style(style2);
		worksheet2.cell(4, 2, 4, 5, true).style(style2);
		worksheet2.cell(7, 2, 7, 5, true).style(style2);
		worksheet2.cell(5, 1, 5, 5, true).style(style2);
		worksheet2.cell(6, 2).style(style2);
		worksheet2.cell(15, 2, 15, 5, true).style(style2);
		worksheet2.cell(16, 2, 16, 5, true).style(style2);

		// Added the content for each cells same as in worksheet 1
		worksheet2.cell(4, 2).string("Payout Invoice").style(style1);
		worksheet2.cell(6, 3).string("Delivered Orders").style(style5);
		worksheet2.cell(6, 4).string("Cancelled Orders").style(style5);
		worksheet2.cell(6, 5).string("Total").style(style5);
		worksheet2.cell(8, 2).string("Number of Orders").style(style10);
		worksheet2.cell(9, 2).string("Item Total").style(style10);
		worksheet2.cell(10, 2).string("Packaging and Service Charges").style(style10);
		worksheet2.cell(11, 2).string("Discounts").style(style10);
		worksheet2.cell(12, 2).string("Net Bill Value").style(style10);
		worksheet2.cell(13, 2).string("GST on order(Including Cess)").style(style10);
		worksheet2.cell(8, 3).string("34").style(style13);
		worksheet2.cell(9, 3).string("5460").style(style13);
		worksheet2.cell(10, 3).string("0.0").style(style13);
		worksheet2.cell(11, 3).string("987").style(style13);
		worksheet2.cell(12, 3).string("6900").style(style13);
		worksheet2.cell(13, 3).string("149").style(style13);
		worksheet2.cell(8, 4).string("1").style(style13);
		worksheet2.cell(9, 4).string("160.0").style(style13);
		worksheet2.cell(10, 4).string("0.0").style(style13);
		worksheet2.cell(11, 4).string("32.0").style(style13);
		worksheet2.cell(12, 4).string("128.0").style(style13);
		worksheet2.cell(13, 4).string("0.0").style(style13);
		worksheet2.cell(8, 5).string("35").style(style13);
		worksheet2.cell(9, 5).string("8545").style(style13);
		worksheet2.cell(10, 5).string("0.0").style(style13);
		worksheet2.cell(11, 5).string("1111").style(style13);
		worksheet2.cell(12, 5).string("7282").style(style13);
		worksheet2.cell(13, 5).string("150").style(style13);
		worksheet2.cell(14, 1).string("A").style(style3);
		worksheet2.cell(14, 2).string("Total Customer Payable").style(style4);
		worksheet2.cell(14, 3).string("7232").style(style13);
		worksheet2.cell(14, 4).string("128.0").style(style13);
		worksheet2.cell(14, 5).string("7485").style(style13);
		worksheet2.cell(17, 2).string("Platform Service Fee").style(style10);
		worksheet2.cell(18, 2).string("Discount on service Fee").style(style10);
		worksheet2.cell(19, 2).string("Collection Charges").style(style10);
		worksheet2.cell(20, 2).string("Access Charges").style(style10);
		worksheet2.cell(21, 2).string("Merchant Cancellation charges").style(style10);
		worksheet2.cell(22, 2).string("Call Center Service Fee").style(style10);
		worksheet2.cell(23, 2).string("Total Service Fees (Before Taxes)").style(style10);
		worksheet2.cell(24, 2).string("Taxes(GST,CESS)over service fee").style(style10);
		worksheet2.cell(17, 3).string("1652").style(style13);
		worksheet2.cell(18, 3).string("0.0").style(style13);
		worksheet2.cell(19, 3).string("0.0").style(style13);
		worksheet2.cell(20, 3).string("0.0").style(style13);
		worksheet2.cell(21, 3).string("0.0").style(style13);
		worksheet2.cell(22, 3).string("0.0").style(style13);
		worksheet2.cell(23, 3).string("1742").style(style13);
		worksheet2.cell(24, 3).string("319").style(style13);
		worksheet2.cell(17, 4).string("0.0").style(style13);
		worksheet2.cell(18, 4).string("0.0").style(style13);
		worksheet2.cell(19, 4).string("0.0").style(style13);
		worksheet2.cell(20, 4).string("0.0").style(style13);
		worksheet2.cell(21, 4).string("0.0").style(style13);
		worksheet2.cell(22, 4).string("0.0").style(style13);
		worksheet2.cell(23, 4).string("0.0").style(style13);
		worksheet2.cell(24, 4).string("0.0").style(style13);
		worksheet2.cell(17, 5).string("1669").style(style13);
		worksheet2.cell(18, 5).string("0.0").style(style13);
		worksheet2.cell(19, 5).string("0.0").style(style13);
		worksheet2.cell(20, 5).string("0.0").style(style13);
		worksheet2.cell(21, 5).string("0.0").style(style13);
		worksheet2.cell(22, 5).string("0.0").style(style13);
		worksheet2.cell(23, 5).string("1884").style(style13);
		worksheet2.cell(24, 5).string("302").style(style13);
		worksheet2.cell(25, 2).string("Total Service Fees").style(style4);
		worksheet2.cell(25, 3).string("2048").style(style13);
		worksheet2.cell(25, 4).string("0.0").style(style13);
		worksheet2.cell(25, 5).string("3045").style(style13);
		worksheet2.cell(25, 1).string("B").style(style3);

		// Added the global style to remove the border for unwanted cells
		worksheet2.cell(24, 1).style(globalBorder);
		worksheet2.cell(22, 1).style(globalBorder);
		worksheet2.cell(26, 1).style(globalBorder);
		worksheet2.cell(26, 2).style(globalBorder);
		worksheet2.cell(26, 3).style(globalBorder);
		worksheet2.cell(26, 4).style(globalBorder);
		worksheet2.cell(26, 5).style(globalBorder);
		worksheet2.cell(26, 6).style(globalBorder);
		worksheet2.cell(25, 6).style(globalBorder);
		worksheet2.cell(24, 6).style(globalBorder);
		worksheet2.cell(23, 6).style(globalBorder);
		worksheet2.cell(22, 6).style(globalBorder);
		worksheet2.cell(21, 6).style(globalBorder);

		//All Orders Sheet

		// Defined this code to merge the cells and give a common style and added content.
		worksheet3.cell(1, 1, 1, 47, true).style(titleStyle);
		worksheet3.cell(2,1,2,4,true).string("Order Details").style(style4);
		worksheet3.cell(2,5,2,9,true).string("Merchant Details").style(style4);
		worksheet3.cell(2,10,2,28,true).string("Service charges").style(style4);

		// Set the height for each rows in worksheet 3
		worksheet3.row(3).setHeight(60);
		worksheet3.column(1).setWidth(50);
		worksheet3.column(2).setWidth(30);
		worksheet3.column(3).setWidth(20);
		worksheet3.column(4).setWidth(20);
		worksheet3.column(5).setWidth(20);
		worksheet3.column(6).setWidth(20);
		worksheet3.column(7).setWidth(20);
		worksheet3.column(8).setWidth(20);
		worksheet3.column(9).setWidth(20);
		worksheet3.column(10).setWidth(40);
		worksheet3.column(11).setWidth(20);
		worksheet3.column(12).setWidth(25);
		worksheet3.column(13).setWidth(20);
		worksheet3.column(14).setWidth(20);
		worksheet3.column(15).setWidth(20);
		worksheet3.column(16).setWidth(20);
		worksheet3.column(17).setWidth(20);
		worksheet3.column(18).setWidth(20);
		worksheet3.column(19).setWidth(20);
		worksheet3.column(20).setWidth(20);
		worksheet3.column(21).setWidth(20);
		worksheet3.column(22).setWidth(20);
		worksheet3.column(23).setWidth(20);
		worksheet3.column(24).setWidth(20);
		worksheet3.column(25).setWidth(20);
		worksheet3.column(26).setWidth(20);
		worksheet3.column(27).setWidth(20);
		worksheet3.column(28).setWidth(20);
		worksheet3.column(29).setWidth(20);
		worksheet3.column(30).setWidth(20);
		worksheet3.column(31).setWidth(20);
		worksheet3.column(32).setWidth(20);
		worksheet3.column(33).setWidth(20);
		worksheet3.column(34).setWidth(20);
		worksheet3.column(35).setWidth(20);
		worksheet3.column(36).setWidth(20);
		worksheet3.column(37).setWidth(20);
		worksheet3.column(38).setWidth(20);
		worksheet3.column(39).setWidth(20);
		worksheet3.column(40).setWidth(20);
		worksheet3.column(41).setWidth(20);
		worksheet3.column(42).setWidth(20);
		worksheet3.column(43).setWidth(20);
		worksheet3.column(44).setWidth(20);
		worksheet3.column(45).setWidth(20);
		worksheet3.column(46).setWidth(20);
		worksheet3.column(47).setWidth(20);

		// Defined the content for each cells in worksheet 3
		worksheet3.cell(3, 1).string("Order Date").style(style7);
		worksheet3.cell(3, 2).string("Order No").style(style7);
		worksheet3.cell(3, 3).string("Order Status").style(style7);
		worksheet3.cell(3, 4).string("Order Category").style(style7);
		worksheet3.cell(3, 6).string("Item's total A").style(style7);
		worksheet3.cell(3, 5).string("Cancelled By?").style(style7);
		worksheet3.cell(3, 6).string("Item's total A").style(style7);
		worksheet3.cell(3, 7).string("Packaging & Service charges \n B").style(style7);
		worksheet3.cell(3, 8).string("Merchant Discount \n C1").style(style7);
		worksheet3.cell(3, 9).string("Exclusive Offer\n C2").style(style7);
		worksheet3.cell(3, 10).string("Total Merchant Discount C= C1 + C2").style(style7);
		worksheet3.cell(3, 11).string("Net Bill Value (without taxes) \n D = A + B - C)").style(style7);
		worksheet3.cell(3, 12).string("GST on order (including cess) E ").style(style7);
		worksheet3.cell(3, 13).string("Customer payable\n(Net bill value after taxes & discount)\nF = D + E").style(style12);
		worksheet3.cell(3, 14).string("Order Date").style(style7);
		worksheet3.cell(3, 15).string("Platform Service Fee Chargeable On").style(style7);
		worksheet3.cell(3, 16).string("Platform Service Fee % (%)").style(style7);
		worksheet3.cell(3, 17).string("Platform Service Fee \n G").style(style7);
		worksheet3.cell(3, 18).string("Discount on Platform Service Fee H").style(style7);
		worksheet3.cell(3, 19).string("Collection Charges I").style(style7);
		worksheet3.cell(3, 20).string("Access Charges J").style(style7);
		worksheet3.cell(3, 21).string("Merchant Cancellation Charges K").style(style7);
		worksheet3.cell(3, 22).string("Call Center Service Fees  L").style(style7);
		worksheet3.cell(3, 23).string("Total Service fee (without taxes) M = G-H+I+J+K+L").style(style7);
		worksheet3.cell(3, 24).string("Taxes on Service fee (Including Cess) N").style(style7);
		worksheet3.cell(3, 25).string("Total service fee (including taxes) O = M + N").style(style12);
		worksheet3.cell(3, 26).string("Cash Prepayment to Merchant P").style(style7);
		worksheet3.cell(3, 27).string("Merchant Share of Cancelled Orders Q = D*x%").style(style7);
		worksheet3.cell(3, 28).string("Delivery fee  (sponsored by merchant) R1").style(style7);
		worksheet3.cell(3, 29).string("GST Deduction U/S 9(5) R2").style(style7);
		worksheet3.cell(3, 30).string("Refund for Customer Complaints R3").style(style7);
		worksheet3.cell(3, 31).string("Disputed Order Remarks").style(style12);
		worksheet3.cell(3, 32).string("Total of Order Level Adjustments S = P + Q + R1 + R2 + R3").style(style7);
		worksheet3.cell(4, 1).string("2022-03-06 10:12:12").style(style6);
		worksheet3.cell(5, 1).string("2022-03-06 10:12:12").style(style6);
		worksheet3.cell(6, 1).string("2022-03-06 10:12:12").style(style6);
		worksheet3.cell(4, 2).string("384452632154").style(style6);
		worksheet3.cell(5, 2).string("384452632154").style(style6);
		worksheet3.cell(6, 2).string("384452632154").style(style6);
		worksheet3.cell(4, 3).string("delivered").style(style6);
		worksheet3.cell(5, 3).string("delivered").style(style6);
		worksheet3.cell(6, 3).string("delivered").style(style6);
		worksheet3.cell(4, 4).string("regular").style(style6);
		worksheet3.cell(5, 4).string("regular").style(style6);
		worksheet3.cell(6, 4).string("regular").style(style6);
		worksheet3.cell(4, 6).string("150").style(style6);
		worksheet3.cell(5, 6).string("150").style(style6);
		worksheet3.cell(6, 6).string("150").style(style6);
		worksheet3.cell(4, 7).string("0").style(style6);
		worksheet3.cell(5, 7).string("0").style(style6);
		worksheet3.cell(6, 7).string("0").style(style6);
		worksheet3.cell(4, 8).string("0").style(style6);
		worksheet3.cell(5, 8).string("45").style(style6);
		worksheet3.cell(6, 8).string("0").style(style6);
		worksheet3.cell(4, 9).string("0").style(style6);
		worksheet3.cell(5, 9).string("0").style(style6);
		worksheet3.cell(6, 9).string("0").style(style6);
		worksheet3.cell(4, 10).string("0").style(style6);
		worksheet3.cell(5, 10).string("47").style(style6);
		worksheet3.cell(6, 10).string("0").style(style6);
		worksheet3.cell(4, 11).string("150").style(style6);
		worksheet3.cell(5, 11).string("345").style(style6);
		worksheet3.cell(6, 11).string("69").style(style6);
		worksheet3.cell(4, 12).string("0").style(style6);
		worksheet3.cell(5, 12).string("6").style(style6);
		worksheet3.cell(6, 12).string("4").style(style6);
		worksheet3.cell(4, 13).string("150").style(style6);
		worksheet3.cell(5, 13).string("150").style(style6);
		worksheet3.cell(6, 13).string("150").style(style6);
		worksheet3.cell(4, 14).string("150").style(style6);
		worksheet3.cell(5, 14).string("150").style(style6);
		worksheet3.cell(6, 14).string("150").style(style6);
		worksheet3.cell(4, 15).string("24.0%").style(style6);
		worksheet3.cell(5, 15).string("24.0%").style(style6);
		worksheet3.cell(6, 15).string("24.0%").style(style6);
		worksheet3.cell(4, 16).string("36").style(style6);
		worksheet3.cell(5, 16).string("36").style(style6);
		worksheet3.cell(6, 16).string("36").style(style6);
		worksheet3.cell(4, 17).string("0").style(style6);
		worksheet3.cell(5, 17).string("0").style(style6);
		worksheet3.cell(6, 17).string("0").style(style6);
		worksheet3.cell(4, 18).string("0").style(style6);
		worksheet3.cell(5, 18).string("0").style(style6);
		worksheet3.cell(6, 18).string("0").style(style6);
		worksheet3.cell(4, 19).string("0").style(style6);
		worksheet3.cell(5, 19).string("0").style(style6);
		worksheet3.cell(6, 19).string("0").style(style6);
		worksheet3.cell(4, 20).string("0").style(style6);
		worksheet3.cell(5, 20).string("0").style(style6);
		worksheet3.cell(6, 20).string("0").style(style6);
		worksheet3.cell(4, 21).string("0").style(style6);
		worksheet3.cell(5, 21).string("0").style(style6);
		worksheet3.cell(6, 21).string("0").style(style6);
		worksheet3.cell(4, 22).string("36").style(style6);
		worksheet3.cell(5, 22).string("36").style(style6);
		worksheet3.cell(6, 22).string("36").style(style6);
		worksheet3.cell(4, 23).string("6").style(style6);
		worksheet3.cell(5, 23).string("6").style(style6);
		worksheet3.cell(6, 23).string("6").style(style6);
		worksheet3.cell(4, 24).string("40").style(style6);
		worksheet3.cell(5, 24).string("40").style(style6);
		worksheet3.cell(6, 24).string("40").style(style6);
		worksheet3.cell(4, 25).string("0").style(style6);
		worksheet3.cell(5, 25).string("0").style(style6);
		worksheet3.cell(6, 25).string("0").style(style6);
		worksheet3.cell(4, 26).string("0").style(style6);
		worksheet3.cell(5, 26).string("0").style(style6);
		worksheet3.cell(6, 26).string("0").style(style6);
		worksheet3.cell(4, 27).string("0").style(style6);
		worksheet3.cell(5, 27).string("0").style(style6);
		worksheet3.cell(6, 27).string("0").style(style6);
		worksheet3.cell(4, 28).string("6").style(style6);
		worksheet3.cell(5, 28).string("12").style(style6);
		worksheet3.cell(6, 28).string("4").style(style6);
		worksheet3.cell(4, 29).string("0").style(style6);
		worksheet3.cell(5, 29).string("0").style(style6);
		worksheet3.cell(6, 29).string("0").style(style6);

		// Defined the name of the file to be generated
		// The file extension can be changed csv or xls, but xlsx will be the better choice.
		let fileName = "invoice-template.xlsx";

		// Defined the absolute file path for the file to be generated
		let filePath = path.resolve(process.cwd(), "media", fileName);

		// Defined the absolute file path for the renamed file.
		// This is for replacing the file while regenerating.
		let newPath = path.resolve(process.cwd(), "media", fileName);

		// Rename the file to the new path
		fs.rename(filePath, newPath, () => {
			console.log("File Renamed!");
		});

		// Log a message to the console indicating that the Excel file has been generated
		console.log("Excel Generated!");

		// Write the Excel workbook to the original file path
		workbook.write(filePath);

		// Send a response to the client indicating that the Excel file has been generated
		res.status(200).send("Excel file generated!");
	} catch (error) {
		console.log(error);
	}
};