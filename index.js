/* eslint-disable */
import XLSX from 'xlsx';

const json_to_array = (key, jsonData) => jsonData.map(v => key.map(j => v[j]));

const get_header_row = (sheet) => {
  const headers = [];
  const range = XLSX.utils.decode_range(sheet['!ref']);
  let C;
  const R = range.s.r; /* start in the first row */
  for (C = range.s.c; C <= range.e.c; ++C) { /* walk every column in the range */
    let cell = sheet[XLSX.utils.encode_cell({c: C, r: R})]; /* find the cell in the first row */
    let hdr = 'UNKNOWN ' + C; // <-- replace with your desired default
    if (cell && cell.t) hdr = XLSX.utils.format_cell(cell);
    headers.push(hdr);
  }
  return headers;
};

export default {
  /**
   * 读取
   * @param data
   * @param type
   * @returns {{header: Array, results: (Array|*)}}
   */
  read(data, type) {
    const workbook = XLSX.read(data, {type: type});
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];
    const header = get_header_row(worksheet);
    const results = XLSX.utils.sheet_to_json(worksheet);
    return {header, results};
  },
  /**
   * 导出
   * @param key
   * @param data
   * @param title
   * @param filename
   * @returns {*|*}
   */
  exPort({key, data, title, filename}) {
    const wb = XLSX.utils.book_new();
    const arr = json_to_array(key, data);
    arr.unshift(title);
    const ws = XLSX.utils.aoa_to_sheet(arr);
    XLSX.utils.book_append_sheet(wb, ws, filename);
    return XLSX.writeFile(wb, filename + '.xlsx');
  }
};
