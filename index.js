'use strict';

var name = process.env.Name || 'ALL',
	express = require('express'),
	Parser = require('parse-xl'),
	path = require('path'),
	fs = require('fs'),
	PDFDocument = require('pdfkit'),
	_ = require('underscore'),
	app = express();

	app.listen(3000, function connection(err) {
	    if (err instanceof Error) {
	        console.error('Unable to start Server', 3000);
	    } else {
	        console.info('Server started at 3000');
	    }
	});

	console.log('Looking for Payslip of '+name+'\nPlease wait ...');

	var PATH = path.resolve('documents/excel/NOW.xlsx');
	
	var parse = new Parser(PATH);
	
	var sheet = "Sheet1";
	var records = parse.records(sheet);

	if(name === 'ALL'){
		for (var i = 0 ; i < records.length; i++) {
			console.log(records[i]);
		}
	}else if(name === 'INDIVIDUAL'){
		 records.forEach(function(data){ 
		 	var doc = new PDFDocument;
		 	var writeStream = null;
		 	var totalGROSS = parseFloat(data['Basic Pay'].replace(",","")) + parseFloat(data['OT'].replace(",",""));
				var totalDeduction = parseFloat(data['PHIC'].replace(",","")) + parseFloat(data['Pag-Ibig'].replace(",","")) + parseFloat(data['SSS'].replace(",","")) + parseFloat(data['Withholding Tax'].replace(",","")) + parseFloat(data['P I S O'].replace(",","")) + parseFloat(data['Loans'].replace(",",""));

				var NETTOTAL = totalGROSS - totalDeduction;
				writeStream = fs.createWriteStream('documents/pdf/' + data.Name + '.pdf');
 				doc.image('img/logo.png', 170, 60,{fit: [80, 80]})
 					.font('Times-Roman')
 					.fontSize(20)
 					.text('PAY SLIP', 380, 70)
 					.moveDown(0.5)
					.moveTo(140, 110)
   					.lineTo(490, 110)
   					.stroke('#808080')
   					.moveTo(140, 111)
   					.lineTo(490, 111)
   					.stroke('#808080')
   					.moveTo(140, 112)
   					.lineTo(490, 112)
   					.stroke('#808080')
   					.moveTo(140, 113)
   					.lineTo(490, 113)
   					.stroke('#808080')
   					.moveTo(140, 114)
   					.lineTo(490, 114)
   					.stroke('#808080')
   					.moveTo(140, 115)
   					.lineTo(490, 115)
   					.stroke('#808080')
   					.moveTo(140, 116)
   					.lineTo(490, 116)
   					.stroke('#808080')
   					.moveTo(140, 117)
   					.lineTo(490, 117)
   					.stroke('#808080')
   					.moveTo(140, 118)
   					.lineTo(490, 118)
   					.stroke('#808080')
   					.moveTo(140, 119)
   					.lineTo(490, 119)
   					.stroke('#808080')
   					.moveTo(140, 120)
   					.lineTo(490, 120)
   					.stroke('#808080')
   					.font('fonts/Calibri Bold.ttf')
 					.fontSize(7)
 					.fillColor('white')
 					.text(data['Payroll Date'],280, 112);

   				doc.font('fonts/Calibri.ttf')
   					.fillColor('#383131')
   					.fontSize(9)
   					.text('Employee: ',150, 140, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['Name'])
   					.moveTo(150, 150)
   					.lineTo(290, 150)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('ID No.: ',350, 140, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['ID'])
   					.moveTo(350, 150)
   					.lineTo(470, 150)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Position: ',150, 155, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['Position'])
   					.moveTo(150, 165)
   					.lineTo(290, 165)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Status: ',350, 155, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['Status'])
   					.moveTo(350, 165)
   					.lineTo(470, 165)
   					.stroke('#544545');

   				doc.font('fonts/Calibri Bold.ttf')
   					.fontSize(8)
   					.text('GROSS INCOME',290, 190);

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Basic Pay',150, 200)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Special Non-Working Holiday Pay',150, 210)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Rest day Pay (2 days)',150, 220)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('OT',150, 230)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['Basic Pay'],440, 200,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('00.00 ',440, 210,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('00.00 ',440, 220,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['OT'],440, 230,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Total',390, 250)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(totalGROSS,410, 250,{
   						width: 60,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri Bold.ttf')
   					.fontSize(8)
   					.text('DEDUCTIONS',292, 270);

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('PHIC',150, 280)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('HDMF',150, 290)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('SSS',150, 300)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('TAX',150, 310)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('PISO',150, 320)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('PISO INS',150, 330)


   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['PHIC'],440, 280,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['Pag-Ibig'],440, 290,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['SSS'],440, 300,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['Withholding Tax'],440, 310,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['P I S O'],440, 320,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(data['Loans'],440, 330,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Total',390, 350)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(totalDeduction,410, 350,{
   						width: 60,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri Bold.ttf')
   					.fontSize(8)
   					.text('NET PAY',380, 370)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(NETTOTAL,410, 370,{
   						width: 60,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Received by: ',150, 410, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('')
   					.moveTo(150, 420)
   					.lineTo(290, 420)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Approved by: ',350, 410, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('')
   					.moveTo(350, 420)
   					.lineTo(470, 420)
   					.stroke('#544545');

				doc.rect(140,40, 350, doc.y).stroke('#544545');

				doc.pipe(writeStream);
				doc.end()
		 })
		
	}else {
		for (var i = 0 ; i < records.length; i++) {
			var regexp = new RegExp(name, 'g');
			if(records[i].Name.match(regexp)){
				var doc = new PDFDocument;
		 		var writeStream = null;
				var totalGROSS = parseFloat(records[i]['Basic Pay'].replace(",","")) + parseFloat(records[i]['OT'].replace(",",""));
				var totalDeduction = parseFloat(records[i]['PHIC'].replace(",","")) + parseFloat(records[i]['Pag-Ibig'].replace(",","")) + parseFloat(records[i]['SSS'].replace(",","")) + parseFloat(records[i]['Withholding Tax'].replace(",","")) + parseFloat(records[i]['P I S O'].replace(",","")) + parseFloat(records[i]['Loans'].replace(",",""));

				var NETTOTAL = totalGROSS - totalDeduction;
				writeStream = fs.createWriteStream('documents/pdf/' + records[i].Name + '.pdf');
 				doc.image('img/logo.png', 170, 60,{fit: [80, 80]})
 					.font('Times-Roman')
 					.fontSize(20)
 					.text('PAY SLIP', 380, 70)
 					.moveDown(0.5)
					.moveTo(140, 110)
   					.lineTo(490, 110)
   					.stroke('#808080')
   					.moveTo(140, 111)
   					.lineTo(490, 111)
   					.stroke('#808080')
   					.moveTo(140, 112)
   					.lineTo(490, 112)
   					.stroke('#808080')
   					.moveTo(140, 113)
   					.lineTo(490, 113)
   					.stroke('#808080')
   					.moveTo(140, 114)
   					.lineTo(490, 114)
   					.stroke('#808080')
   					.moveTo(140, 115)
   					.lineTo(490, 115)
   					.stroke('#808080')
   					.moveTo(140, 116)
   					.lineTo(490, 116)
   					.stroke('#808080')
   					.moveTo(140, 117)
   					.lineTo(490, 117)
   					.stroke('#808080')
   					.moveTo(140, 118)
   					.lineTo(490, 118)
   					.stroke('#808080')
   					.moveTo(140, 119)
   					.lineTo(490, 119)
   					.stroke('#808080')
   					.moveTo(140, 120)
   					.lineTo(490, 120)
   					.stroke('#808080')
   					.font('fonts/Calibri Bold.ttf')
 					.fontSize(7)
 					.fillColor('white')
 					.text(records[i]['Payroll Date'],280, 112);

   				doc.font('fonts/Calibri.ttf')
   					.fillColor('#383131')
   					.fontSize(9)
   					.text('Employee: ',150, 140, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['Name'])
   					.moveTo(150, 150)
   					.lineTo(290, 150)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('ID No.: ',350, 140, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['ID'])
   					.moveTo(350, 150)
   					.lineTo(470, 150)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Position: ',150, 155, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['Position'])
   					.moveTo(150, 165)
   					.lineTo(290, 165)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Status: ',350, 155, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['Status'])
   					.moveTo(350, 165)
   					.lineTo(470, 165)
   					.stroke('#544545');

   				doc.font('fonts/Calibri Bold.ttf')
   					.fontSize(8)
   					.text('GROSS INCOME',290, 190);

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Basic Pay',150, 200)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Special Non-Working Holiday Pay',150, 210)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Rest day Pay (2 days)',150, 220)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('OT',150, 230)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['Basic Pay'],440, 200,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('00.00 ',440, 210,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('00.00 ',440, 220,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['OT'],440, 230,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Total',390, 250)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(totalGROSS,410, 250,{
   						width: 60,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri Bold.ttf')
   					.fontSize(8)
   					.text('DEDUCTIONS',292, 270);

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('PHIC',150, 280)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('HDMF',150, 290)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('SSS',150, 300)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('TAX',150, 310)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('PISO',150, 320)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('PISO INS',150, 330)


   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['PHIC'],440, 280,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['Pag-Ibig'],440, 290,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['SSS'],440, 300,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['Withholding Tax'],440, 310,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['P I S O'],440, 320,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(records[i]['Loans'],440, 330,{
   						width: 30,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('Total',390, 350)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(totalDeduction,410, 350,{
   						width: 60,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri Bold.ttf')
   					.fontSize(8)
   					.text('NET PAY',380, 370)

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text(NETTOTAL,410, 370,{
   						width: 60,
   						align: 'right'
   					})

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Received by: ',150, 410, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('')
   					.moveTo(150, 420)
   					.lineTo(290, 420)
   					.stroke('#544545');

   				doc.font('fonts/Calibri.ttf')
   					.fontSize(9)
   					.text('Approved by: ',350, 410, {
   						lineBreak: true,
                        continued: true
   					})
   					.font('fonts/Calibri.ttf')
   					.fontSize(8)
   					.text('')
   					.moveTo(350, 420)
   					.lineTo(470, 420)
   					.stroke('#544545');

				doc.rect(140,40, 350, doc.y).stroke('#544545');

				doc.pipe(writeStream);

				doc.end()
				break;
			}
		}
	}

	


