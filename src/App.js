import React, { useState } from "react";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import "./App.css";
import excelFileImage from "./assets/images/excelFile.png";
import arrowImage from "./assets/images/arrow.png";
import textFileImage from "./assets/images/textFile.svg";
import uploadImage from './assets/images/uploadFile.png'
import closeImage from './assets/images/close.png'

function App() {
  const [file, setFile] = useState(null);
  const [fileData, setFileData] = useState(null);
  const [totalCount, setTotalCount] = useState(0);
  const [successCount, setSuccessCount] = useState(0);
  const [totalAmount, setTotalAmount] = useState(0);

  const handleFileUpload = (e) => {
    const selectedFile = e.target.files[0];

    if (!selectedFile) return alert("No file selected");

    setFile(selectedFile);

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

        const isOnlyColumn5NonEmpty = row.every((cell, index) => {
          if (index === 5) {

            return cell != null && cell !== "";
          }

          return cell == null || cell === "";
        });

        if (isOnlyColumn5NonEmpty) {
          break;
        }


        if (row[5] == null || row[5] === "") {
          continue;
        }

        totalSuccessCount = totalSuccessCount + 1;
        totalSheetAmount = totalSheetAmount + row[5];

        filteredData.push(
          ["N,",
            row[9] ,
            row[5],
            row[6]?.trim("")  ,
            ",,,,,,,",
            row[1]?.trim(""),
            ",,,,,,,",
            excelDateToJSDate(row[2]),
            "", row[10]?.trim(""),
            row[7]?.trim(""),
            row[8]?.trim(""),
            row[11]?.trim("")]
            .join(",")
        );
        filteredData.push("\n");
      }
      setTotalCount(totalDataCount - 1);
      setSuccessCount(totalSuccessCount);
      setTotalAmount(totalSheetAmount);
      setFileData(filteredData);
    };

    

    reader.readAsArrayBuffer(selectedFile);
  };

  function excelDateToJSDate(serial) {
    const excelEpoch = new Date(1900, 0, 1);
    const date = new Date(excelEpoch.getTime() + (serial - 2) * 86400000);
    const month = ("0" + (date.getMonth() + 1)).slice(-2);
    const day = ("0" + date.getDate()).slice(-2);
    const year = date.getFullYear().toString();
    return `${day}/${month}/${year}`;
  }

  console.log(fileData);
  
  const saveAsTextFile = () => {
    if (!fileData) return alert("No data to save");

    const currentDate  = new Date()
    const Day = currentDate.getDate()
    const Month = currentDate.getMonth() +1
    
    console.log(`Plato SCX ${Day}${Month}.2.txt`);
    
    const isTrue = window.confirm(
      "Total sheet data count: " +
      totalCount +
      "\nTotal success data count:" +
      successCount +
      "\nTotal Amount: " +
      totalAmount.toFixed(2) +
      "\n\nDo you want to save this data ?"
    );

    

    if (isTrue) {
      const textData = fileData.join("");
      const blob = new Blob([textData], { type: "text/plain;charset=utf-8" });
      saveAs(blob, `Plato SCX ${Day}${Month}.2.txt`);
      alert("File downloaded successfully");
    }
  };

  const handelFileClose = () => {
    setFile(null);
    setFileData(null);
    setTotalCount(0);
    setSuccessCount(0);
    setTotalAmount(0);
  }

  return (
    <div className="App">
      <div className="card">
        <div className="card-content">
          <h1>Excel to Text File Converter</h1>


          <div className="images-container" >
            <img className="image excel-img" src={excelFileImage} alt="Excel file image" />
            <img className="image arrow-img" src={arrowImage} alt="Arrow image" />
            <img className="image text-img" src={textFileImage} alt="Text file image" />
          </div>
          {
            !file ? <div className="upload-file-container" >
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={(e) => handleFileUpload(e)}
              />
              <img className="upload-img" src={uploadImage} alt="Excel file" />
              <h3>Drop your excel file here or Browser</h3>
              <p>Supports: xlsx,xls</p>
            </div> :
              <div className="file-info" >
                <h3>{file.name}</h3>
                <p>{(file.size / 1024).toFixed(2)} KB</p>
                <img className="icon close" src={closeImage} alt="Close" onClick={handelFileClose} />
              </div>
          }

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
