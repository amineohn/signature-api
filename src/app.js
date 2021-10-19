const fs = require("fs");
const express = require("express");
const app = express();
const cors = require("cors");
const path = require("path");
const port = process.env.PORT || 3001;
const multer = require("multer");
app.use(cors());
app.use(express());
app.use(express.json());

var bodyParser = require("body-parser");
app.use(bodyParser.urlencoded({ extended: false }));

app.use((req, res, next) => {
  const origin = req.get("origin");
  res.header("Access-Control-Allow-Origin", origin);
  res.header("Access-Control-Allow-Credentials", true);
  res.header(
    "Access-Control-Allow-Methods",
    "GET, POST, OPTIONS, PUT, PATCH, DELETE"
  );
  res.header(
    "Access-Control-Allow-Headers",
    "Origin, X-Requested-With, Content-Type, Accept, Authorization, Cache-Control, Pragma"
  );
  if (req.method === "OPTIONS") {
    res.sendStatus(204);
  } else {
    next();
  }
});
app.use(
  "/img",
  express.static(path.join(__dirname, "/generated/assets/images"))
);
app.use("/assets", express.static(path.join(__dirname, "/generated/assets")));

app.get("/preview", (req, res) => {
  res.sendFile(__dirname + "/generated/index.html");
});

app.listen(port, () => console.log(`Listening on port ${port}`));

var storage = multer.diskStorage({
  destination: (req, file, callBack) => {
    callBack(null, path.join(__dirname, "/generated/assets/images"));
  },
  filename: (req, file, callBack) => {
    callBack(
      null,
      file.fieldname + "-" + Date.now() + path.extname(file.originalname)
    );
  },
});

var upload = multer({
  storage: storage,
});
app.post(`/generate`, upload.single("file"), (req, res) => {
  if (!req.file) {
    console.log("No file upload");
  } else {
    console.log(req.file.filename);
    console.log("file uploaded");
  }
  let data = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:w="urn:schemas-microsoft-com:office:word"
    xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
    xmlns="http://www.w3.org/TR/REC-html40">
    
    <head>
    <meta charset="utf-8">
    <meta http-equiv=Content-Type content="text/html; charset=windows-1252">
    <meta name=ProgId content=Word.Document>
    <meta name=Generator content="Microsoft Word 15">
    <meta name=Originator content="Microsoft Word 15">
    <link rel=File-List href="/assets/filelist.xml">
    <link rel=themeData href="/assets/themedata.thmx">
    <link rel=colorSchemeMapping href="/assets/colorschememapping.xml">
    <meta content="text/html; charset=utf-8" http-equiv=Content-Type>
    <table style="border-spacing: 0px;">  
    
      <tr>
          <td>
            <a href="https://les-detritivores.org/" style="text-decoration:none"><img moz-do-not-send="false" style="" src="./assets/images/${req.file.filename}" alt="image profil"/>
          </td>
    
          <td style="padding-top: 0px; padding-left: 10px;">

            <span style="font-family: Arial, Helvetica, Sans-Serif; font-size: 15pt;  color: #e94e1b;"><strong>${req.body.FirstName} ${req.body.LastName}</strong>
              <br />
      
              <span style="font-size: 11pt; font-family: Arial, Helvetica, Sans-Serif; color: #263b29;">
                <strong>${req.body.Function}</strong>
              </span>
            </span>

            <br />
            <br />

            <span style="color: #263b29 ;font-size: 9pt;font-family: Arial, Helvetica, Sans-Serif; white-space: nowrap; font-weight: 700">
              ${req.body.Mail}
            </span>
            <br />
            <span style="color: #263b29 ;font-size: 9pt;font-family: Arial, Helvetica, Sans-Serif; white-space: nowrap; font-weight: 500">
              ${req.body.ProNumber}
            </span>
            <br />
            <span style="color: #263b29 ;font-size: 9pt;font-family: Arial, Helvetica, Sans-Serif; white-space: nowrap; font-weight: 500">
              ${req.body.Number}
            </span>
      
            <br />
            <span style="font-family: Arial, Helvetica, Sans-Serif; font-size: 9pt;  color: #263b29; font-weight: 500">${req.body.Adress}</span>
            <br />
            <a href="https://${req.body.Link}/">
              <span style="font-family:  Arial, Helvetica, Sans-Serif; font-size: 8.5pt; color: #e94e1b; font-weight: 900">${req.body.Link}</span>
            </a>

            <table style="border-spacing: 0px;">
              <th>
                <a mc:disable-tracking href="https://www.facebook.com/lesdetritivores/" style="text-decoration: none;">
                  <img style="vertical-align: bottom; padding-top: 10px; margin-left: 0px;" data-input="facebook" data-tab="social" src="./assets/images/facebook.png" />
                </a>
              </th>
      
              <th>
                <a mc:disable-tracking href="https://www.instagram.com/lesdetritivores/?hl=fr" style="text-decoration: none;">
                  <img style="vertical-align: bottom; padding-top: 12px; margin-left: 4px" data-input="insta" data-tab="social" src="./assets/images/insta.png" />
                </a>
              </th>
      
              <th>
                <a mc:disable-tracking href="https://www.linkedin.com/company/les-d%C3%A9tritivores/?originalSubdomain=fr" style="text-decoration: none; margin-left: 4px">
                  <img style=" vertical-align: bottom; padding-top: 12px;" data-input="linkedin" data-tab="social" src="./assets/images/linkedin.png" />
                </a>
              </th>
          </table>
    </table>`;
  res.json({
    firstname: req.body.firstname,
  });
  fs.writeFile("src/generated/index.html", data, (err) => {
    if (err) throw err;
  });
});
