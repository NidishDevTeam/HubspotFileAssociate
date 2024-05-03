const express = require('express');
const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const multer = require('multer');
const upload = multer();
const app = express();
const PORT = process.env.PORT || 3000;
const xlsx = require('xlsx');
const path = require('path');

// Middleware to parse JSON bodies
// app.use(express.json());
// app.use(bodyParser.json());
// Replace these with your actual HubSpot API key and file path
const HUBSPOT_API_KEY = 'api key goes here';

const FILE_PATH = 'path/to/your/file.pdf';



app.get('/getdeal', async (req, res) => {
    const response = await axios.get(`https://api.hubapi.com/crm/v3/objects/deals/search?filterGroups=[{"filters":[{"propertyName":"zoho_id","operator":"EQ","value":"8596320147"}]}]&hapikey=${HUBSPOT_API_KEY}`);
    //  response.data.results[0];
    res.status(200).json({ response: response.data });
})
app.post('/testupload', async (req, res) => {
    try {
        upload.single('file')(req, res, async function (err) {
            if (err) {
                throw new Error("Error handling file upload: " + err.message);
            }

            if (!req.file) {
                throw new Error("No file uploaded.");
            }

            const form = new FormData();
            const fileOptions = {
                access: 'PUBLIC_NOT_INDEXABLE',
                overwrite: false,
                duplicateValidationStrategy: 'NONE',
                duplicateValidationScope: 'ENTIRE_PORTAL',
            };

            form.append('file', req.file.buffer, req.file.originalname); // Append file buffer
            form.append('options', JSON.stringify(fileOptions));
            form.append('folderPath', 'testFolder');

            const config = {
                method: 'post',
                url: 'https://api.hubapi.com/files/v3/files',
                headers: {
                    Authorization: `Bearer ${HUBSPOT_API_KEY}`,
                    ...form.getHeaders(),
                },
                data: form,
            };

            const response = await axios(config);

            console.log(JSON.stringify(response.data));
            res.status(200).json({ response: response.data });
        });
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({ errormsg: "changes " + error.message });
    }
})


// app.post('/testuploadMulti', async (req, res) => {
//     try {
//         var UploadedFilesUrl=[];
//         var fileresponse;
//         var noteresponse;
//         var zohoid="8596320147";
//         const dealResponse = await axios({
//             method: 'get',
//             url: `https://api.hubapi.com/crm/v3/objects/deals/${zohoid}?idProperty=zohoid`,
//             headers: {
//                 Authorization: `Bearer ${HUBSPOT_API_KEY}`,

//             },
//         });
//         if(dealResponse.status!=200)
//         {

//           return  res.status(500).json({ errormsg: "Deal was not found" });

//         }

//          console.log("id of deal"+JSON.stringify(dealResponse.data["id"]));
//         var dealID=dealResponse.data["id"];
//         upload.single('file')(req, res, async function (err) {

//             if (err) {
//                 throw new Error("Error handling file upload: " + err.message);
//             }

//             if (!req.file) {
//                 throw new Error("No file uploaded.");
//             }

//             const form = new FormData();
//             const fileOptions = {
//                 access: 'PUBLIC_NOT_INDEXABLE',
//                 overwrite: false,
//                 duplicateValidationStrategy: 'NONE',
//                 duplicateValidationScope: 'ENTIRE_PORTAL',
//             };

//             form.append('file', req.file.buffer, req.file.originalname); // Append file buffer
//             form.append('options', JSON.stringify(fileOptions));
//             form.append('folderPath', 'testFolder');

//                     const config = {
//                         method: 'post',
//                         url: 'https://api.hubapi.com/files/v3/files',
//                         headers: {
//                             Authorization: `Bearer ${HUBSPOT_API_KEY}`,
//                             ...form.getHeaders(),
//                         },
//                         data: form,
//                     };

//                     fileresponse = await axios(config);

//                     console.log("after fiel response ")
//                     console.log(fileresponse.status)
//                         if(fileresponse.status==200||fileresponse.status==201)
//                         {//create note 
//                             noteresponse= await axios({
//                                 method: 'post',
//                                 url: 'https://api.hubapi.com/crm/v3/objects/notes',
//                                 headers: {
//                                     Authorization: `Bearer ${HUBSPOT_API_KEY}`,
//                                 },
//                                 data: {
//                                     "properties":{"hs_attachment_ids": fileresponse.data["id"],
//                                     "hs_timestamp":new Date().toISOString()}

//                                  },
//                              })
//                              console.log("after notecreate response ")
//                              console.log(noteresponse.status)

//                              if(noteresponse.status==201||noteresponse.status==200)
//                              {
//                                 console.log("inside noteresponse")
//                                 //associateapo call
//                                // https://api.hubspot.com/crm/v3/objects/notes/50666634569/associations/deal/18530520855/214
//                                try{
//                                 var associateResponse= await axios({
//                                     method: 'put',
//                                     url: `https://api.hubspot.com/crm/v3/objects/notes/${noteresponse.data["id"]}/associations/deal/${dealID}/214`,
//                                     headers: {
//                                         Authorization: `Bearer ${HUBSPOT_API_KEY}`,
//                                     }
//                                    })
//                                    console.log("after associate")
//                                    console.log(
//                                     associateResponse.status
//                                     )
//                                }catch(e){

//                                  console.log("errori n assoxiate")
//                                  res.status(400).json({error :e})

//                                }


//                                         // if(associateResponse.status==200)
//                                         // {
//                                         //     console.log(
//                                         //     "file associated "
//                                         //     )

//                                         //     res.status(200).json({ response: associateResponse.data });
//                                         // }
//                              }
//                         }

//         });
//     } catch (error) {
//         // console.error("Error:", error);
//         res.status(500).json({ errormsg: "error "});
//     }
// })


// app.get('/getexceldata',async(req,res)=>{
// try{
//     console.log("get called")
//     var listofdata=[];
// // Path to your Excel file
// const filePath = path.join(__dirname, '/test.xlsx');

// // Load the Excel workbook
// const workbook = xlsx.readFile(filePath);

// // Assuming the first sheet in the workbook contains the data
// const sheetName = workbook.SheetNames[0];
// console.log("sheetname"+sheetName)
// const sheet = workbook.Sheets[sheetName];

// // Assuming your data starts from cell A1
// const range = xlsx.utils.decode_range(sheet['!ref']);

// // Iterate through each cell in the range
// for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
//     for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
//         const cellAddress = xlsx.utils.encode_cell({ r: rowNum, c: colNum });
//         const cell = sheet[cellAddress];
//         // Access the value of the cell
//         const cellValue = cell ? cell.v : null;
//         listofdata.push(cellValue)
//         console.log(`Cell ${cellAddress} has value: ${cellValue}`);
//     }
// }
// res.send({message:"success",list:listofdata});
// }
// catch(e)
// {     console.log("error on get"+e)
//     res.send({message:e});

// }

// })
app.get('/getexceldata', (req, res) => {
    try {
        // Path to your Excel file
        const filePath = path.join(__dirname, '/test.xlsx');

        // Load the Excel workbook
        const workbook = xlsx.readFile(filePath);

        // Assuming the first sheet in the workbook contains the data
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        // Assuming your data starts from cell A2 (assuming A1 contains headers)
        const range = xlsx.utils.decode_range(sheet['!ref']);

        const zohoIdFilePathArray = [];

        // Iterate through each row in the range
        for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
            const zohoIdCellAddress = `B${rowNum + 1}`;
            const filePathCellAddress = `C${rowNum + 1}`;

            const zohoId = sheet[zohoIdCellAddress] ? sheet[zohoIdCellAddress].v : null;
            const filePath = sheet[filePathCellAddress] ? sheet[filePathCellAddress].v : null;

            if (zohoId && filePath) {
                zohoIdFilePathArray.push({ zohoId, filePath });
            }
        }

        // Send the collected data as JSON response
        res.json({ data: zohoIdFilePathArray });
    } catch (error) {
        console.error("Error:", error);
        res.status(500).json({ error: "Internal server error" });
    }
});

app.get('/testuploadMulti', async (req, res) => {
    const baseFileName = 'response';
    const fileExtension = 'txt';
    const now = new Date();
    let outputPath = `${baseFileName}.${fileExtension}`;
    let counter = 1;
    try {
     
        const commanRoot = "C:/Nidish Projects/hubspotassociateServer/";
        const filePath = path.join(__dirname, '/test.xlsx');

        const workbook = xlsx.readFile(filePath);
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const range = xlsx.utils.decode_range(sheet['!ref']);

        const zohoIdFilePathArray = [];
        var typeOF = "Id";
        for (let rowNum = range.s.r + 1; rowNum <= range.e.r; rowNum++) {
            const zohoIdCellAddress = `B${rowNum + 1}`;
            const filePathCellAddress = `C${rowNum + 1}`;

            const zohoId = sheet[zohoIdCellAddress] ? sheet[zohoIdCellAddress].v : null;
            const filePath = sheet[filePathCellAddress] ? sheet[filePathCellAddress].v : null;

            if (zohoId && filePath) {
                zohoIdFilePathArray.push({ zohoId, filePath });
            }
        }
        console.log(JSON.stringify(zohoIdFilePathArray))
        var ArrayofAssociate = []
        var errorList = [];
        var ID;

        for (const { zohoId, filePath } of zohoIdFilePathArray) {
            console.log("zohoid" + zohoId + "FilePath" + filePath)

            var isDealfilter = false;

            try {
                const DealResponse = await axios.get(`https://api.hubapi.com/crm/v3/objects/deals/${zohoId}?idProperty=zoho_deal_record_id_updated`, {

                    // const DealResponse = await axios.get(`https://api.hubapi.com/crm/v3/objects/deals/${zohoId}?idProperty=zohoid`, {
                    headers: { Authorization: `Bearer ${HUBSPOT_API_KEY}` }
                });
                if (DealResponse.status == 200) {

                    isDealfilter = true;
                    ID = DealResponse.data.id
                    typeOF = "dealId";
                }
            }

            catch (err) {
                if (err.message = "Request failed with status code 404") {
                    isDealfilter = false;
                    try {

                        const CompanyResponse = await axios.get(`https://api.hubapi.com/crm/v3/objects/companies/${zohoId}?idProperty=zoho_company_record_id_updated`, {

                            // const CompanyResponse = await axios.get(`https://api.hubapi.com/crm/v3/objects/companies/${zohoId}?idProperty=zohoid`, {
                            headers: { Authorization: `Bearer ${HUBSPOT_API_KEY}` }

                        });
                        if (
                            CompanyResponse.status != 200
                        ) {
                            errorList.push({ "zohoid": zohoId, "message": "company not found" })
                            continue;

                        } else {
                            ID = CompanyResponse.data.id
                            typeOF = "CompanyId";

                        }

                    } catch (e) {

                        errorList.push({ "zohoid": zohoId, "message": "not found in deal and company" })
                        continue;
                    }



                } else {
                    errorList.push({ "zohoid": zohoId, "message": "error in deal filter" })

                    continue;
                }
            }



            var fileName

            console.log("get deal" + ID)
            try {
                fileName = filePath;

                const fileBuffer = fs.readFileSync(commanRoot + filePath);
                const form = new FormData();
                const fileOptions = {
                    access: 'PUBLIC_NOT_INDEXABLE',
                    overwrite: false,
                    duplicateValidationStrategy: 'NONE',
                    duplicateValidationScope: 'ENTIRE_PORTAL',
                };

                form.append('file', fileBuffer, fileName);
                form.append('options', JSON.stringify(fileOptions));
                form.append('folderPath', 'testFolder');

                const config = {
                    headers: {
                        Authorization: `Bearer ${HUBSPOT_API_KEY}`,
                        ...form.getHeaders(),
                    },
                    data: form,
                };
                var fileresponse;


                fileresponse = await axios.post('https://api.hubapi.com/files/v3/files', form, config);

            } catch (e) {

                errorList.push({ "ID": ID, "type": typeOF, "zohoid": zohoId, "message": "error in file upload", "error": e })
                continue;
            }

            if (fileresponse.status >= 200 && fileresponse.status <= 299) {
                console.log("after file upload id " + fileresponse.data.id);
                var noteresponse;
                try {
                    noteresponse = await axios.post('https://api.hubapi.com/crm/v3/objects/notes', {
                        properties: {
                            hs_attachment_ids: fileresponse.data.id,
                            hs_timestamp: new Date().toISOString()
                        }
                    }, {
                        headers: { Authorization: `Bearer ${HUBSPOT_API_KEY}` }
                    });

                } catch (e) {
                    errorList.push({ "ID": ID, "type": typeOF, "zohoid": zohoId, "message": "error in note creation", "error": e })
                    continue;
                }

                if (noteresponse.status >= 200 && noteresponse.status <= 299) {
                    console.log("noteresponse" + noteresponse.data.id)
                    console.log("dealid" + ID)
                    var associateUrl;
                    // https://api.hubspot.com/crm/v3/objects/notes/50679212303/associations/deal/18530520855/214

                    if (isDealfilter) {
                        associateUrl = `https://api.hubspot.com/crm/v3/objects/notes/${noteresponse.data.id}/associations/deal/${ID}/214`;
                    } else {
                        console.log("inside comapny associte")
                        https://api.hubspot.com/crm/v4​/objects​/notes​/51165406522​/associations​/default​/company​/20290499093






                        associateUrl = `https://api.hubspot.com/crm/v3/objects/notes/${noteresponse.data.id}/associations/company/${ID}/190`;
                        console.log("company  urll " + associateUrl)
                    }
                    try {
                        const associateResponse = await axios.put(
                            associateUrl,
                            null,
                            {
                                headers: {
                                    Authorization: `Bearer ${HUBSPOT_API_KEY}`,
                                    'Content-Type': 'application/json'
                                }
                            }
                        );
                        if (associateResponse.status == 200) {
                            ArrayofAssociate.push({ "ID": ID, "type": typeOF, "zohoid": zohoId, "noteID": noteresponse.data.id, "associateId": associateResponse.data.id, "fileName": fileName })
                            console.log("successfully associated " + associateResponse.data.id)
                            // return res.status(200).json({ message:"successfully associated id " ,response: associateResponse.data });

                        } else {
                            errorList.push({ "ID": ID, "type": typeOF, "zohoid": zohoId, "noteID": noteresponse.data.id, "fileName": fileName, "message": "error on association", "error": e })
                            console.log("Error in associate:");
                            continue;
                            // return res.status(400).json({ error: associateResponse });

                        }
                    } catch (e) {
                        errorList.push({ "ID": ID, "type": typeOF, "zohoid": zohoId, "noteID": noteresponse.data.id, "fileName": fileName, "message": "error on association", "error": e })
                        continue;
                    }



                } else {
                    errorList.push({ "ID": ID, "type": typeOF, "zohoid": zohoId, "fileName": fileName, "message": "error on creating note" });
                    continue;

                }
            } else {
                errorList.push({ "ID": ID, "type": typeOF, "zohoid": zohoId, "fileName": fileName, "message": "error on fileupload" });
                continue;
            }

        }

        // const zohoid = "8596320147";
        const jsonString = JSON.stringify({ message: "successfully executed", dateTime: now, "Associated": ArrayofAssociate, "ErrorList": errorList }, null, 2); 


        // Function to check if file exists
        function fileExists(outputPath) {
            try {
                fs.accessSync(outputPath);
                return true;
            } catch (err) {
                return false;
            }
        }

        // Check if file exists, if so, increment counter and try again
        while (fileExists(outputPath)) {
            outputPath = `${baseFileName}_${counter}.${fileExtension}`;
            counter++;
        }

        // Write JSON string to a text file
        fs.writeFile(outputPath, jsonString, (err) => {
            if (err) {
                console.error('Error writing file:', err);
            } else {
                console.log('JSON response has been written to', outputPath);
            }
        });
        res.status(200).json({ message: "successfully associated", details: ArrayofAssociate, errorlist: errorList })

    } catch (error) {
        console.error("Error:", error);
        // Function to check if file exists
        function fileExists(outputPath) {
            try {
                fs.accessSync(outputPath);
                return true;
            } catch (err) {
                return false;
            }
        }

        // Check if file exists, if so, increment counter and try again
        while (fileExists(outputPath)) {
            outputPath = `${baseFileName}_${counter}.${fileExtension}`;
            counter++;
        }

        // Write JSON string to a text file
        const jsonString = JSON.stringify({ message: "faced Error", dateTime: now, "Associated": ArrayofAssociate, "ErrorList": errorList, error: error }, null, 2);
        fs.writeFile(outputPath, jsonString, (err) => {
            if (err) {
                console.error('Error writing file:', err);
            } else {
                console.log('JSON response has been written to', outputPath);
            }
        });
        res.status(500).json({ errormsg: "Internal server error", errorlist: errorList, error: error });
    }
});


// Start the server
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
