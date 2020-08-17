var express =   require("express");
var multer  =   require('multer');
const puppeteer = require('puppeteer');
var excel = require('excel4node');
var app     =   express();
//var port = process.env.PORT || 3000;

function exportToExcel(data) {
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();

    // Add Worksheets to the workbook
    var worksheet = workbook.addWorksheet('Sheet 1');

    const headings = ['Link'];//, '', 'Gift2','Gift3', 'Gift4', 'Gift5', 'Gift6', 'Gift7'];;

    // Writing from cell A1 to I1
    headings.forEach((heading, index) => {
        worksheet.cell(1, index + 1).string(heading);
    })

    // Writing from cell A2 to I2 , A3 to I3, .....
    data.forEach((item, index) => {
        worksheet.cell(index + 2, 1).string(item.link);
    });
    var filename = "Linkcholon"  + '.xlsx'; //+ Date.now().toString()
    workbook.write("./public/"+filename);
}

function getLink(file){
const XLSX = require('xlsx');
var workbook = XLSX.readFile(file);
var sheet_name_list = workbook.SheetNames;
var urls = [];
data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); 
for(var key in data){
urls.push(data[key]['Article']);
}
return urls;
};

var storage =   multer.diskStorage({
  destination: function (req, file, callback) {
    callback(null, './uploads');
  },
  filename: function (req, file, callback) {
    callback(null,file.originalname);
  }
});


var upload = multer({ storage : storage}).single('userPhoto');

app.get('/',function(req,res){
      res.sendFile(__dirname + "/views/index.html");
});

app.post('/upload',function(req,res){
    upload(req,res,function(err) {
        if(err) {
            return res.end("Error uploading file.");
        }else{ // open else
        	var file = "./uploads/" + req.file.originalname;
        	(async () => {
			    const browser = await puppeteer.launch({ headless: true })
			    const page = await browser.newPage()
			    const urls = getLink(file);
			    console.log(urls);
			    let arrInfo = [];
			    for (let i = 0; i < urls.length; i++) {
			        try {
			            await page.goto("https://dienmaycholon.vn/tu-khoa/"+urls[i], { timeout: 300000 });
			            const info = await page.evaluate(() => {
			                // let checkweb = document.querySelector(".themNoel")
			                let checkweb = document.querySelector(".content_search")
			                if (checkweb !== null) {
			                    const checklink = document.querySelector(".item_product .pro_infomation a")
			                    const link = document.querySelector(checklink !== null ? ".item_product .pro_infomation a" : ".khongcoclass");
			                    return {
			                        //...data,
			                        link: link ? "https://dienmaycholon.vn"+link.getAttribute('href') : "Not found",
			                    }

			                }

			                return {
			                    link: "https://dienmaycholon.vn/tu-khoa/khong-tim-thay"
			                };

			            })
			            if (info) {
			                arrInfo.push(info)
			                console.log(info)
			            }
			        } catch (err) {
			            console.log("Có lỗi xảy ra", err);
			        }
			    }
			    exportToExcel(arrInfo);
			    res.end("<a href='" + __dirname + "/public/Linkcholon.xlsx'>Download File</a>");
			    //await browser.close();
			})();
        } // end else
        // var name = req.file.originalname;
        // res.end("<a href='./uploads/"+name+"'>Download File</a>");
    });
});

app.listen(process.env.PORT || 3000);
