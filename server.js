const express = require("express");
const bodyParser = require("body-parser");
const { Octokit } = require("@octokit/rest");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
app.use(bodyParser.json());

const GITHUB_TOKEN = "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"; // Replace with your token
const REPO_OWNER = "Rahulkrishnanj";
const REPO_NAME = "Chendaclass";
const FILE_PATH = "Namedemo.xlsx";
const BRANCH = "main";

const octokit = new Octokit({ auth: GITHUB_TOKEN });

// Fetch Excel file
async function getExcelFile() {
  const { data } = await octokit.repos.getContent({
    owner: REPO_OWNER,
    repo: REPO_NAME,
    path: FILE_PATH,
    ref: BRANCH,
  });

  const content = Buffer.from(data.content, "base64");
  fs.writeFileSync("temp.xlsx", content);
  return "temp.xlsx";
}

// Update Excel file
async function updateExcelFile(filePath) {
  const content = fs.readFileSync(filePath).toString("base64");

  const { data } = await octokit.repos.getContent({
    owner: REPO_OWNER,
    repo: REPO_NAME,
    path: FILE_PATH,
    ref: BRANCH,
  });

  await octokit.repos.createOrUpdateFileContents({
    owner: REPO_OWNER,
    repo: REPO_NAME,
    path: FILE_PATH,
    message: "Update name entry",
    content: content,
    sha: data.sha,
    branch: BRANCH,
  });
}

// Endpoint to save name to Excel
app.post("/update-excel", async (req, res) => {
  try {
    const { name } = req.body;
    const filePath = await getExcelFile();

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const sheet = workbook.getWorksheet(1) || workbook.addWorksheet("Names");
    sheet.addRow([name]);
    await workbook.xlsx.writeFile(filePath);

    await updateExcelFile(filePath);

    res.status(200).send("Entry saved successfully!");
  } catch (error) {
    console.error(error);
    res.status(500).send("Error updating Excel file");
  }
});

app.listen(3000, () => console.log("Server running on http://localhost:3000"));
