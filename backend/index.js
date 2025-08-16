const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const ExcelJS = require("exceljs");
const fs = require("fs");

const app = express();
const PORT = 3001;

app.use(cors());
app.use(bodyParser.json());

app.post("/save", async (req, res) => {
  const data = req.body;
  const filePath = "./datos.xlsx";
  let workbook;

  if (fs.existsSync(filePath)) {
    workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
  } else {
    workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Datos");
    sheet.addRow([
      "Saludo", "Nombre", "Apellido", "Gènero",
      "Email", "Fecha de Nacimiento", "Direccion"
    ]);
  }

  const sheet = workbook.getWorksheet("Datos");
  sheet.addRow([
    data.Saludo,
    data.Saludo.includes("Ninguno") ? 1 : 0,
    data.Saludo.includes("Sr.") ? 1 : 0,
    data.Saludo.includes("Sra.") ? 1 : 0,
    data.Saludo.includes("Srta..") ? 1 : 0,
    data.Nombre,
    data.Apellido,
    data.Gènero,
    data.Email,
    data.FechadeNacimiento, 
    data.Direcciòn, 
  ]);

  await workbook.xlsx.writeFile(filePath);
  res.send("Datos guardados correctamente");
});

app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});