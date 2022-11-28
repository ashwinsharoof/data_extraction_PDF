const fs = require('fs')
const pdfParse = require('pdf-parse')
const excelJS = require("exceljs");

let extractPDF = async (file) => {
  let fileSync = fs.readFileSync(file)
 
    let Parse = await pdfParse(fileSync)
    //console.log('Content: ', Parse.text)
    //console.log('PDF pages: ', Parse.numpages)
    var result = []
    let content = Parse.text
    let arr = content.split("Name of the Exhibitor")
    let address = content.split("Address")
    let contact = content.split("Contact Person")
    let desgination = content.split("Designation")
    let mobile = content.split("Mobile")
    let email = content.split("Email")
    let website = content.split("Website")
    let profile = content.split("Profile")
    for(let i =0; i<arr.length; i++){
       
        var arr2 = arr[i].split("Address")
        
        var arr3 = address[i].split("Contact Person")
      /*  if (typeof contact[i] === "undefined"){
            console.log("yes")
        }
        else {
            var arr4 = contact[i].split("Designation")
            result.push(
                {
                    //"Name_of_the_Exhibitor":arr2[0],
                    //"Address" : arr3[0],
                    "Contact_Person" : arr4[0],
                    //"Designation" : arr5[0],
                    //"Mobile" : arr6[0],
                    //"Email": arr7[0],
                    //"Website": arr8[0],
                    //"Profile": arr9[0]
                    
                }) 

        }*/
     
        var arr5 = desgination[i].split("Mobile")
        var arr6 = mobile[i].split("Email")
        var arr7 = email[i].split("Website")
        var arr8 = website[i].split("Profile")
        var arr9 = profile[i].split("Name of the Exhibitor")
        result.push(
            {
                //"Name_of_the_Exhibitor":arr2[0],
                //"Address" : arr3[0],
                //"Contact_Person" : arr4[0],
                //"Designation" : arr5[0],
                //"Mobile" : arr6[0],
                //"Email": arr7[0],
                //"Website": arr8[0],
                "Profile": arr9[0]
                
            })  
       
        
    }
    console.log(result)
    const workbook = new excelJS.Workbook();  // Create a new workbook
    const worksheet = workbook.addWorksheet("My Users"); // New Worksheet
    const path = "./";  // Path to download excel
    worksheet.columns = [
        { header: "Name of the Exhibitor", key: "Name_of_the_Exhibitor", width: 5 },
        { header: "Address", key: "Address", width: 5 },
        { header: "Contact Person", key: "Contact_Person", width: 5 },
        { header: "Designation", key: "Designation", width: 5 },
        { header: "Mobile", key: "Mobile", width: 5 },
        { header: "Email", key: "Email", width: 5 },
        { header: "Website", key: "Website", width: 5 },
        { header: "Profile", key: "Profile", width: 5 },
      ];

      result.forEach((user) => {
        worksheet.addRow(user); // Add data in worksheet
      });
      worksheet.getRow(1).eachCell((cell) => {
        cell.font = { bold: true };
      });
      try {
        const data = await workbook.xlsx.writeFile(`${path}/profile.xlsx`)
          .then(() => {
            console.log("success");
          });
      } catch (err) {
        console.log(err);
      }
    };
 

let pdfRead = './sample.pdf'
extractPDF(pdfRead)