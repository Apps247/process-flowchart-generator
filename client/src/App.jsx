import { useState, useRef } from 'react'
import Spreadsheet from "react-spreadsheet";
import * as XLSX from 'xlsx/xlsx.mjs';
import JSZip from 'jszip';
import './App.css'

function App() {
  const columnLabels = ['Step ID', 'Description', 'Responsible', 'Decision'];
  const initialData = [
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }],
    [{ value: '' }, { value: '' }, { value: '' }, { value: '' }]
  ];

  const [data, setData] = useState(initialData)
  const dataRef = useRef(initialData)

  const handleFileUpload = (e) => {
    const file = e.target.files[0]
    console.log(file)
    setSpreadsheetDataFromFile(file)
  }

  const handleSpreadsheetChange = (updatedData) => {
    dataRef.current = updatedData;
  }

  // convert file to react spreadsheet format
  function setSpreadsheetDataFromFile(file) {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = e.target.result
      const workbook = XLSX.read(data, { type: 'binary' })
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      const sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 })
      const spreadsheetData = sheetData.map(row => row.map(cell => ({ value: cell })))
      dataRef.current = spreadsheetData;
      setData(dataRef.current)
    }
    reader.readAsArrayBuffer(file)
  }

  const convertExcelToVisio = () => {
    document.getElementById("download-link")?.remove();
    document.getElementById("output-preview").style.display = 'none';
    document.getElementById("output-preview").src = "";
    const cleanedData = dataRef.current.map(row => ({
      'Step ID': row[0].value,
      Description: row[1].value,
      Responsible: row[2].value,
      Decision: row[3].value
    }));
    console.log(cleanedData)
    fetch('http://localhost:5000/generate-flowchart', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(cleanedData),
    })
      .then(response => response.blob())
      .then(JSZip.loadAsync)
      .then(zip => {
        // Extract the Visio file
        const visioFile = zip.files['output.vsdx'];
        if (visioFile) {
          visioFile.async('blob').then(blob => {
            const url = window.URL.createObjectURL(blob);
            const link = document.createElement('a');
            link.textContent = 'Download Visio File';
            link.href = url;
            link.setAttribute('download', 'output.vsdx');
            document.getElementById("output").appendChild(link);
            console.log("Visio file ready to download")
          });
        }

        // Extract the image file
        const imageFile = zip.files['output.png'];
        if (imageFile) {
          imageFile.async('blob').then(blob => {
            const url = window.URL.createObjectURL(blob);
            // const link = document.createElement('a');
            // link.textContent = 'Download Image File';
            // link.href = url;
            // link.setAttribute('download', 'output.png');
            // document.getElementById("output").appendChild(link);
            document.getElementById("output-preview").src = url;
            document.getElementById("output-preview").style.display = 'block';
          });
        }
      })
      .catch(error => {
        console.error('Error:', error);
      });
  }

  return (
    <>
      <div className="container" style={{ display: 'flex', alignItems: 'center'}}>
        <div className="column" style={{ display: "flex", flexDirection: "column" }}>
          <input type="file" onChange={handleFileUpload} />
          <Spreadsheet data={data} columnLabels={columnLabels} onChange={handleSpreadsheetChange} />
        </div>
        <div className="column" style={{ marginLeft: '20px', marginRight: '20px', }}>
          <button onClick={convertExcelToVisio}>Convert</button>
        </div>
        <div id="output" className="column" style={{ display: "flex", flexDirection: "column"}}>
          <div style={{ width: '500px', height: '400px', backgroundColor: 'white' }}>
            <img id='output-preview' src="" alt="output" style={{display: 'none', maxWidth: '500px'}} />
          </div>
        </div>
      </div>
    </>
  )
}

export default App
