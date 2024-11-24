import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import "./App.css";

function App() {
  const [fileData, setFileData] = useState(null);
  const [totalCount, setTotalCount] = useState(0);
  const [successCount, setSuccessCount] = useState(0);
  const [totalAmount, setTotalAmount] = useState(0);

  const handleFileUpload  = (e) => {
    const file = e.target.files[0];
    if (!file) return alert("No file selected");

    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: "array" });
      const worksheet = workbook.Sheets[workbook.SheetNames[0]];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        range: 1,
      });

      const filteredData = [];
      let totalDataCount = 0;
      let totalSuccessCount = 0;
      let totalSheetAmount = 0;

      for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];

        totalDataCount = totalDataCount + 1;

        if (row.every((cell) => cell === null || cell === "")) {
          break;
        }

        if (row[5] == null || row[5] === "") {
          continue;
        }

        totalSuccessCount = totalSuccessCount + 1;
        totalSheetAmount = totalSheetAmount + row[5];

        filteredData.push(
          ["N,", row[9], row[5], row[6], ",,,,,,,", row[1], ",,,,,,,", excelDateToJSDate(row[2]), "", row[10], row[7], row[8], row[11]].join(",")
        );
      }

      setTotalCount(totalDataCount - 1);
      setSuccessCount(totalSuccessCount);
      setTotalAmount(totalSheetAmount / 2);
      setFileData(filteredData);
    };

    reader.readAsArrayBuffer(file);
  };

  function excelDateToJSDate(serial) {
    const excelEpoch = new Date(1900, 0, 1);
    const date = new Date(excelEpoch.getTime() + (serial - 2) * 86400000);
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const day = ("0" + date.getDate()).slice(-2);
    const year = date.getFullYear().toString();
    return `${day}/${month}/${year}`;
  }

  const saveAsTextFile = () => {
    if (!fileData) return alert("No data to save");

    const isTrue = window.confirm(
      "Total sheet data count: " +
        totalCount +
        "\nTotal success data count:" +
        successCount +
        "\nTotal Amount: " +
        totalAmount
    );

    if (isTrue) {
      const textData = fileData.join("\n");
      const blob = new Blob([textData], { type: "text/plain;charset=utf-8" });
      saveAs(blob, "filtered_data.txt");
      alert("File downloaded successfully");
    }
  };

  return (
    <div className="App">
      <div className="card">
        <div className="card-content">
          <h1>Excel to Text File Converter</h1>

          <input
            type="file"
            accept=".xlsx, .xls, .numbers"
            onChange={(e) => handleFileUpload(e)}
          />

          <div className="buttons-container">
            {fileData && (
              <button onClick={saveAsTextFile}>Save as text file</button>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
