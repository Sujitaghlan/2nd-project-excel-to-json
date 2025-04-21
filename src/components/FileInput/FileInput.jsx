import { useRef, useState } from "react"
import * as XLSX from 'xlsx'
import './FileInput.css'

const FileInput = () => {
  const [data, setData] = useState([]);
  const [fileName, setFileName] = useState('');
  const [sheetName, setSheetName] = useState([]);
  const [selectedSheet, setSelectedSheet] = useState('');
  const workbookRef = useRef(null);
  const fileRef = useRef();

  const checkExcelFile = (file) => {
    const allowedExtensions = ['.xlsx', '.xls'];
    const fileName = file.name;
    const fileExtension = fileName.slice(fileName.lastIndexOf('.')).toLowerCase();
    return allowedExtensions.includes(fileExtension);
  }
  const handleFileUpload = (e) => {
    const file = e.target.files[0];
    const MAX_FILE_SIZE = 5 * 1024 * 1024;
    if(!file) return;

    if(checkExcelFile(file) && file.size <= MAX_FILE_SIZE){
      const reader = new FileReader();
      reader.readAsArrayBuffer(file);
      reader.onload = (e) => {
      const data1 = e.target.result;
      const workbook = XLSX.read(data1, {type: "buffer"});
      workbookRef.current = workbook;
      setSheetName(workbook.SheetNames);
      console.log(sheetName[0]);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const parsedData = XLSX.utils.sheet_to_json(sheet, {defval:""})
      const isAllRowsFilled = parsedData.every(row => 
        Object.values(row).every(value => value !== "" && value !== null)
      );

      if(isAllRowsFilled){
        setData(parsedData);
      }
      else{
        alert("Fill all the rows in excel File");
      }

    }
    setFileName(file.name);
    }
    else{
      alert("File must be in .xlsx or .xls format");
    }
  }
  //console.log(data);

  const handleSelectSheet = (e) => {
    const selectedSheetName = e.target.value;
    setSelectedSheet(selectedSheetName);

    if (workbookRef.current) {
      const sheet = workbookRef.current.Sheets[selectedSheetName];
      const parsedData = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const isAllRowsFilled = parsedData.every((row) =>
        Object.values(row).every((value) => value !== "" && value !== null)
      );

      if (isAllRowsFilled) {
        setData(parsedData);
      } else {
        alert("Please ensure all rows in the selected sheet are filled.");
      }
    }
  };

  const handleRemoveFile = () => {
    setFileName("");
    setSheetName([]);
    setSelectedSheet("");
    setData([]);
    workbookRef.current = null;
    if (fileRef.current) fileRef.current.value = "";
  };

  return (
    <div className="input-container">
          <h1>Excel to Json Converter</h1>
      <div class="input-group mb-3">
        <input type="file" class="form-control" id="inputGroupFile02"  onChange={handleFileUpload} ref={fileRef}/>
      </div>
      {sheetName.length > 0 && (
  <select className="form-select" multiple aria-label="Multiple select example" onChange={handleSelectSheet}>
      <option selected>Select Sheets</option>
    {sheetName.map((sheet, index) => (
      <option value={sheet} key={index}>
        {sheet}
      </option>
    ))}
  </select>
)}
        {
          fileName && (
            <div className="remove-container">
            <span>{fileName}</span>
            <button className="removeFile" onClick={handleRemoveFile}>X</button>
            </div>
          )
        }

      <br />

      {
        data && data.length > 0 && (
          <div>
      <table class="table table-striped">
  <thead>
    <tr>
      {
        Object.keys(data[0]).map((key) => (
          <th scope="col" key={key}>{key}</th>  
        ))
      }

    </tr>
  </thead>
  <tbody>
     {
      data.map((row, index) => (
        <tr key={index}>
          {Object.values(row).map((value, index) => (
            <td key={index}>{value}</td>
          ))}
        </tr>
      ))
     }
      
  </tbody>
</table>
      </div>
        )
      }
      
    </div>
  )
}

export default FileInput