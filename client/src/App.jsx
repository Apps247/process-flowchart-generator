import { useState, useRef } from 'react'
import Spreadsheet from "react-spreadsheet";
import * as XLSX from 'xlsx/xlsx.mjs';
import './App.css'

function App() {
  const initialData = [
    [
      { value: 'Step', style: { fontWeight: 'bold' } },
      { value: 'Description', style: { fontWeight: 'bold' } },
      { value: 'Responsible', style: { fontWeight: 'bold' } },
      { value: 'Decision', style: { fontWeight: 'bold' } }
    ],
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
    // console.log(dataRef.current)
    // console.log(dataRef.current[0][0].value)
    console.log(dataRef.current.map(row => row[0].value));
    // const cleanedData = dataRef.current.map(row => ({
    //   step: row[0].value,
    //   description: row[1].value,
    //   responsible: row[2].value,
    //   decision: row[3].value
    // }));
    // fetch('http://localhost:5000/generate-flowchart', {
    //   method: 'POST',
    //   headers: {
    //     'Content-Type': 'application/json'
    //   },
    //   body: JSON.stringify(cleanedData),
    // })
    //   .then(response => response.json())
    //   .then(data => {
    //     // Handle the response data here
    //   })
    //   .catch(error => {
    //     // Handle any errors here
    //   });
  }

  return (
    <>
      <div className="container" style={{ display: 'flex', alignItems: 'center' }}>
        <div className="column" style={{ display: "flex", flexDirection: "column" }}>
          <input type="file" onChange={handleFileUpload} />
          <Spreadsheet data={data} onChange={handleSpreadsheetChange} />
        </div>
        <div className="column" style={{ marginLeft: '20px', marginRight: '20px', }}>
          <button onClick={convertExcelToVisio}>Convert</button>
        </div>
        <div className="column">

        </div>
      </div>
    </>
  )
}

export default App
