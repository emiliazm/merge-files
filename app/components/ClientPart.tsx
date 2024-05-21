"use client";
import { useState } from "react";
import * as XLSX from "xlsx";

function downloadExcelFile(data: any[], fileName: string) {
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Merged Data");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
}

export default function ClientPart() {
  const [excelFile, setExcelFile] = useState<File | null>(null);
  const [jsonText, setJsonText] = useState<string>("");
  // const [mergedData, setMergedData] = useState<any[]>([]);

  const handleExcelFileChange = (event: any) => {
    setExcelFile(event.target.files[0]);
  };

  const handleJsonTextChange = (event: any) => {
    setJsonText(event.target.value);
  };

  const handleSubmit = async () => {
    if (!excelFile || !jsonText) {
      alert("Please upload an Excel file and paste JSON text.");
      return;
    }

    // Read Excel file
    const data = await excelFile.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheetName = workbook.SheetNames[0];
    const options = {
      header: 1,
      range: 5, // Start reading from the fourth row (index starts at 0) because of excel file format
    };
    const rawData: any[][] = XLSX.utils.sheet_to_json(
      workbook.Sheets[sheetName],
      options
    );

    const headers = rawData[0];
    const excelData = rawData.slice(1).map((row: any[]) => {
      let obj: { [key: string]: any } = {};
      row.forEach((cell, index) => {
        obj[headers[index]] = cell; // Map each cell to the header
      });
      return obj;
    });

    // Parse JSON text to compare
    const jsonData = JSON.parse(jsonText);

    // Convert the id to string for easy comparison
    const jsonDataMap = new Map(
      jsonData.map((item: any) => [item.id.toString(), item])
    );

    // Merge the data
    const merged = excelData.map((item: any) => {
      const matchedItem = jsonDataMap.get(item["RSP Reference"]);
      const suspendNotes =
        item["Suspend  Notes"] !== ""
          ? item["Suspend  Notes"]
          : item["Suspend Reason"];

      if (matchedItem) {
        return {
          "RSP Reference": item["RSP Reference"],
          "Suspend Notes": suspendNotes,
          "Company Name": (matchedItem as { [key: string]: any })[
            "company_name"
          ],
        };
      } else {
        return {
          "RSP Reference": item["RSP Reference"],
          "Suspend Notes": suspendNotes,
          "Company Name": "",
        };
      }
    });

    // setMergedData(merged);
    downloadExcelFile(merged, "MergedData");
  };

  return (
    <div className="flex flex-col items-end justify-start gap-4 w-full">
      <div className="flex flex-col items-center justify-center gap-4 w-full">
        <div className="w-full flex flex-col">
          <label htmlFor="jsonText" className="text-sm font-bold">
            Excel File
          </label>
          <input
            type="file"
            name="file"
            onChange={handleExcelFileChange}
            accept=".xlsx, .xls"
            className=""
          />
        </div>
        <div className="w-full">
          <label htmlFor="jsonText" className="text-sm font-bold">JSON</label>
          <textarea
            value={jsonText}
            onChange={handleJsonTextChange}
            placeholder='[
                          {
                            "id": 2000908000,
                            "company_name": "Contact Energy"
                          },
                          {
                            "id": 2001150000,
                            "company_name": "Contact Energy"
                          }
                         ]'
            className="w-full h-80 p-2 border border-gray-300 rounded-md resize-none"
          ></textarea>
        </div>
      </div>

      <button
        onClick={handleSubmit}
        className="hover:bg-orange-700 text-white py-2 px-4 rounded mt-4"
      >
        Merge and Process
      </button>
    </div>
  );
}
