const express = require("express");
const cors = require("cors");
const axios = require("axios");
const { google } = require("googleapis");
const path = require("path");
const xlsx = require("xlsx");
const multer = require("multer");
const fs = require("fs");
const { promisify } = require("util");
const unlinkAsync = promisify(fs.unlink);
require("dotenv").config();
const app = express();
app.use(express.urlencoded({ extended: true }));
app.use(cors());
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const REFRESH_TOKEN = process.env.REFRESH_TOKEN;
const PORT = process.env.PORT || 8080;

const storage = multer.diskStorage({
  destination: path.join(__dirname, "uploads"),
  filename: function (req, file, cb) {
    cb(null, file.originalname);
  },
});
const upload = multer({ storage: storage });

app.use(express.static(path.join(__dirname, "template")));

app.get("/", (req, res) => {
  res.sendFile(`index.html`);
});

const getZohoToken = async () => {
  try {
    const res = await axios.post(
      `https://accounts.zoho.com/oauth/v2/token?client_id=${CLIENT_ID}&grant_type=refresh_token&client_secret=${CLIENT_SECRET}&refresh_token=${REFRESH_TOKEN}`
    );
    console.log(res.data);
    const token = res.data.access_token;
    return token;
  } catch (error) {
    res.send({
      error,
    });
  }
};

const updateContactOnZoho = async ({ phone, config, correct }) => {
  const date = new Date();
  const day = date.getDate();
  const month = date.getMonth() + 1;
  const year = date.getFullYear();
  const attemptDate = `${year}-${month <= 9 ? "0" + month : month}-${
    day <= 9 ? "0" + day : day
  }`;
  // const attemptDate = `2023-09-03`;
  const contact = await axios.get(
    `https://www.zohoapis.com/crm/v3/Contacts/search?phone=${phone}`,
    config
  );
  if (!contact.data || !contact.data.data[0] || !contact.data.data[0].id) {
    console.log("Not a Zoho Contact");
    return "Not a Zoho Contact";
  }

  // console.log("contact", contact.data.data[0]);
  const contactid = contact.data.data[0].id;
  const contactBody = {
    data: [
      {
        id: contactid,
        Workshop_Quiz_Score: correct,
        Workshop_Quiz_Attended_Date: attemptDate,
        $append_values: {
          Workshop_Quiz_Score: true,
          Workshop_Quiz_Attended_Date: true,
        },
      },
    ],
    duplicate_check_fields: ["id"],
    apply_feature_execution: [
      {
        name: "layout_rules",
      },
    ],
    trigger: ["workflow"],
  };

  const updateContact = await axios.post(
    `https://www.zohoapis.com/crm/v3/Contacts/upsert`,
    contactBody,
    config
  );
};

app.post("/view", upload.array("file", 50), async (req, res) => {
  const files = req.files;
  if (files.length === 0) {
    return res
      .status(400)
      .send(
        `<h1 style="display:grid;place-items:center;min-height:100vh;">No files were uploaded.</h1>`
      );
  }
  try {
    const finalUsers = [];
    for (const file of files) {
      console.log(file.path);
      const date = new Date().toDateString();
      const workbook = xlsx.readFile(file.path);
      const sheetName1 = workbook.SheetNames[0];
      const sheet1 = workbook.Sheets[sheetName1];
      const data1 = xlsx.utils.sheet_to_json(sheet1);
      const currentUsers = [];
      for (let i = 8; i < data1.length; i++) {
        const firstname = data1[i][""];
        const lastname = data1[i]["__EMPTY"];
        const attemptDate = new Date(
          data1[i]["__EMPTY_1"].substring(0, 11)
        ).toDateString();
        const obj = { firstname, lastname, attemptDate };

        // ------------------Change date to today--------------------
        // toDateString() format === "Sun Sep 03 2023"
        if (date === attemptDate) {
          currentUsers.push(obj);
        }

        // ----------------------------------------------------------
      }
      const sheetName2 = workbook.SheetNames[2];
      const sheet2 = workbook.Sheets[sheetName2];
      const data2 = xlsx.utils.sheet_to_json(sheet2);
      for (let i = 7; i < data2.length - 2; i++) {
        const firstname = data2[i]["Polling Results"];
        const lastname = data2[i]["__EMPTY"];
        const correct = data2[i]["__EMPTY_1"];

        // ---------change phone field according to the question number----------

        const phone = data2[i]["__EMPTY_2"]?.toString().replace(/[^0-9]/g, "");

        // ----------------------------------------------------------------------

        const userFound = currentUsers.find(
          (user) => user.firstname === firstname && user.lastname === lastname
        );
        if (userFound && phone) {
          const obj = { firstname, lastname, correct, phone };
          finalUsers.push(obj);
        }
      }
      await unlinkAsync(file.path);
    }
    const token = await getZohoToken();
    const config = {
      headers: {
        Authorization: `Zoho-oauthtoken ${token}`,
        "Content-Type": "application/json",
      },
    };
    console.log(finalUsers);
    for (let i = 0; i < finalUsers.length; i++) {
      await updateContactOnZoho({
        phone: finalUsers[i].phone,
        config,
        correct: finalUsers[i].correct,
      });
    }
    let table = "";
    finalUsers.map(
      (user) =>
        (table += `<tr>
          <td style="border:1px solid; padding:10px 20px;">${user.firstname}</td>
          <td style="border:1px solid; padding:10px 20px;">${user.lastname}</td>
          <td style="border:1px solid; padding:10px 20px;">${user.correct}</td>
          <td style="border:1px solid; padding:10px 20px;">${user.phone}</td>
        </tr>`)
    );
    return res.status(200).send(`
    <div style="width : 80%; margin : 50px auto; text-align : center; display : grid; place-items:center;">
      <h1>Excel file uploaded and processed successfully.</h1>
      <Table style="text-align : center; font-size : 20px; margin-top : 20px; border-collapse: collapse; ">
        <Thead>
          <th style="border:1px solid; padding:10px 20px;">First Name</th>
          <th style="border:1px solid; padding:10px 20px;">Last Name</th>
          <th style="border:1px solid; padding:10px 20px;">Correct Answer</th>
          <th style="border:1px solid; padding:10px 20px;">Phone</th>
        </Thead>
        ${table}
      </Table>
    </div>
    `);
  } catch (error) {
    console.error("Error reading Excel file:", error);
    res.status(500).send({ error: "Error reading Excel file." });
    for (const file of files) {
      await unlinkAsync(file.path);
    }
    return;
  }
});

app.listen(PORT, () => {
  console.log(`http://localhost:${PORT}`);
});
