import React, { useState } from 'react';
import ExcelJS from 'exceljs';
import { camelCase } from 'lodash-es';
import './style.css';

function readToStream(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();

    reader.readAsArrayBuffer(file);
    reader.onload = () => {
      const buffer = reader.result;
      resolve(buffer);
    };
  });
}

const SINGLE_SHIPMENT_SHEET_NAME = 'Shipment List - Single Pick Up';
const DATA_START_ROW = 3;

export default function App() {
  const [selectedFile, setSelectedFile] = useState();
  console.log({ selectedFile });

  const extractExcelData = async () => {
    const workbook = new ExcelJS.Workbook();
    const fileStream = await readToStream(selectedFile);
    await workbook.xlsx.load(fileStream);

    const sheet = workbook.getWorksheet(SINGLE_SHIPMENT_SHEET_NAME);
    const rowCount = sheet.actualRowCount - DATA_START_ROW;
    const columns = sheet.getRow(1)?.values?.map((value) => camelCase(value));
    const columnIndexMap = columns.reduce(
      (finalMap, columnName, columnIndex) => ({
        ...finalMap,
        [columnName]: columnIndex,
      }),
      {}
    );
    const rowData = sheet.getRows(DATA_START_ROW, rowCount)?.map((row) => {
      return columns.reduce(
        (finalRowData, columnKey) => ({
          ...finalRowData,
          [columnKey]: row.getCell(columnIndexMap[columnKey])?.value,
        }),
        {}
      );
    });

    console.log({ columns, columnIndexMap, rowData });
  };

  const submit = async () => {
    await extractExcelData();
  };

  return (
    <div>
      <h1>Hello StackBlitz!</h1>
      <p>Start editing to see some magic happen :)</p>
      <div style={{ marginBottom: 32 }}>
        <a
          href="https://docs.google.com/spreadsheets/d/1a0rL23stbV6eYBEZ1nyVMt0ydRl02F-uJ2TFeNpEHe0/edit#gid=791627080"
          target="_blank"
        >
          Download Template
        </a>
      </div>
      <input
        type="file"
        onChange={(event) => setSelectedFile(event?.target?.files?.[0])}
      />
      <div>
        <button onClick={submit}>Submit</button>
      </div>
    </div>
  );
}
