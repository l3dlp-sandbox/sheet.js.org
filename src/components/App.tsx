/* sheet.js.org (C) 2022-present SheetJS LLC -- https://sheetjs.com */
import React, { useState, ChangeEvent, useRef, useCallback } from "react";
import DataGrid, { TextEditor } from "react-data-grid";
import * as _XLSX from "xlsx";
import { WorkBook, WorkSheet, utils, writeFile } from "xlsx";
import { BeatLoader } from "react-spinners";
import { useWorker } from "@koale/useworker";
import Swal from 'sweetalert2'
import withReactContent from 'sweetalert2-react-content'
import { useDropzone } from "react-dropzone";

import "../styles/App.css";

const XLSX = _XLSX;
const parseAB = (ab: ArrayBuffer): WorkBook | Error => {
  try {
    return ((globalThis as any).XLSX as typeof XLSX).read(ab, { WTF: true, dense: true });
  } catch(e) { return e instanceof Error ? e : new Error(e as any); }
}

const TIMEOUT = 10_000; // 10 seconds

const MySwal = withReactContent(Swal)

type Row = any[]; /*{
  [index: string]: string | number;
};*/

type Column = {
  key: string;
  name: string;
  editor: typeof TextEditor;
};

type DataSet = {
  [index: string]: WorkSheet;
};

function getRowsCols(
  data: DataSet,
  sheetName: string
): {
  rows: Row[];
  columns: Column[];
} {
  const rows: Row[] = utils.sheet_to_json(data[sheetName], {header:1});
  let columns: Column[] = [];

  for (let row of rows) {
    const keys: string[] = Object.keys(row);

    if (keys.length > columns.length) {
      columns = keys.map((key) => {
        return { key, name: utils.encode_col(+key), editor: TextEditor };
      });
    }
  }

  return { rows, columns };
}

const exportTypes = ["xlsx", "xlsb", "csv", "html"];

export default function App() {
  const [rows, setRows] = useState<Row[]>([]);
  const [columns, setColumns] = useState<Column[]>([]);
  const [workBook, setWorkBook] = useState<DataSet>({} as DataSet);
  const [sheets, setSheets] = useState<string[]>([]);
  const [current, setCurrent] = useState<string>("");
  const [loading, setLoading] = useState<boolean>(false);
  const fileInput = useRef<HTMLInputElement>(null);

  const [parseWorker, controller] = useWorker(parseAB, {
    remoteDependencies: [
      "https://cdn.sheetjs.com/xlsx-latest/package/dist/xlsx.full.min.js",
    ],
  });

  const onDrop = useCallback(async acceptedFiles => {
    if(loading) return;
    await setLoading(true);
    handleF(acceptedFiles[0]);
  }, [handleF, loading]);
  const { getRootProps, getInputProps, isDragActive } = useDropzone({onDrop});

  async function selectSheet(name: string, reset = true) {
    await setLoading(true);
    if(reset) workBook[current] = utils.json_to_sheet(rows, {
      header: columns.map((col: Column) => col.key),
      skipHeader: true
    });

    const { rows: new_rows, columns: new_columns } = getRowsCols(workBook, name);

    await setRows(new_rows);
    await setColumns(new_columns);
    await setCurrent(name);
    await setLoading(false);
  }

  async function handleFile(ev: ChangeEvent<HTMLInputElement>): Promise<void> {
    await setLoading(true);
    const f = ev.target.files?.[0];
    if(!f) {
      if(fileInput?.current) fileInput.current.value = "";
      return await setLoading(false);
    }
    await handleF(f);
  }
  async function handleF(f: File): Promise<void> {
    if(f.size > 1_048_576) {
      const res = await MySwal.fire({
        icon: "warning",
        title: "Large File",
        text: `File is ${(f.size/1_048_576)>>>0} MB and reading may be slow.  Should we proceed?`,
        confirmButtonText: 'Continue',
        showCancelButton: true,
        cancelButtonText: 'Stop'
      });
      if(!res.isConfirmed) {
        if(fileInput?.current) fileInput.current.value = "";
        return await setLoading(false);
      }
    }
    let end = setTimeout(async() => {
      await controller.kill();
      await MySwal.fire({
        icon: "error",
        title: "Timeout",
        text: `Stopped reading after ${TIMEOUT/1000} seconds`,
        footer: <a href={`mailto:oss@sheetjs.com?subject=Public Demo Error&body=Timeout on file of size ${f.size} bytes`}>We would appreciate the feedback</a>
      });
      if(fileInput?.current) fileInput.current.value = "";
      return await setLoading(false);
    }, TIMEOUT);
    const file = await f.arrayBuffer();
    let data: ReturnType<typeof parseAB> = new Error("");
    data = await parseWorker(file);

    if(data instanceof Error) {
      console.log(data);
      await MySwal.fire({
        icon: "error",
        title: "This file does not appear to be a valid spreadsheet",
        text: `Library Error: ${data.message || data}`,
        footer: <a href={`mailto:oss@sheetjs.com?subject=Public Demo Error&body=${encodeURIComponent(data.message)}`}>We would appreciate the feedback</a>
      });
      if(fileInput?.current) fileInput.current.value = "";
      return await setLoading(false);
    }

    await setWorkBook(data.Sheets);
    await setSheets(data.SheetNames);

    /* repeated from selectSheet since workBook will be stale */
    const name = data.SheetNames[0];
    const { rows: new_rows, columns: new_columns } = getRowsCols(data.Sheets, name);

    await setRows(new_rows);
    await setColumns(new_columns);
    await setCurrent(name);

    clearTimeout(end);

    if(fileInput?.current) fileInput.current.value = "";
    await setLoading(false);
  }

  async function saveFile(ext: string) {
    await setLoading(true);
    const wb = utils.book_new();

    sheets.forEach((n) => {
      utils.book_append_sheet(wb, workBook[n], n);
    });

    writeFile(wb, "sheet." + ext);
    await setLoading(false);
  }

  async function onSelect(e: ChangeEvent) {
    const idx = parseInt(((e as ChangeEvent).target as HTMLSelectElement).value, 10);
    await selectSheet(sheets[idx]);
  }

  return (
    <>
      <input type="file" onChangeCapture={handleFile} ref={fileInput} disabled={loading} />
      <BeatLoader loading={loading} size={10}/>
      {!loading && (<div className='dropzone' {...getRootProps()}>
        <input className='dropinput' {...getInputProps}/>
        {loading ? (<p>Loading file ...</p>) : isDragActive ? (<p>Drop files here!</p>) : (<p>... or drag and drop files here.</p>)}
      </div>)}
      {sheets.length > 0 && (
        <>
          <p>Use the dropdown to switch to a worksheet:&nbsp;
            <select id="wselect" onChange={onSelect}>
              {sheets.map((sheet, idx) => (<option key={sheet} value={idx}>{sheet}</option>))}
            </select>
          </p>
          <div className="flex-cont">
            <b>Current Sheet: {current}</b>
          </div>
          <DataGrid columns={columns} rows={rows} onRowsChange={setRows} />
          <p>Click one of the buttons to create a new file with the modified data</p>
          <div className="flex-cont">
            {exportTypes.map((ext) => (
              <button key={ext} onClick={() => saveFile(ext)}>
                export [.{ext}]
              </button>
            ))}
          </div>
        </>
      )}
    </>
  );
}
