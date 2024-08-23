import "./App.css";
import * as XLSX from "xlsx";
import { Button, Table } from "antd";

function createExcel() {
  const data = getDatasource();

  const workSheet = XLSX.utils.aoa_to_sheet([
    ["Item", "1st Half 2024", "", "", "", "", "", "2nd Half 2024"],
  ]);

  const header = Object.keys(data[0]).map((key) => capitalizeFirstLetter(key));
  XLSX.utils.sheet_add_aoa(workSheet, [header], { origin: "A2" });
  XLSX.utils.sheet_add_json(workSheet, data, { origin: "A3", skipHeader: true });

  workSheet["!merges"] = [
    { s: { r: 0, c: 0 }, e: { r: 1, c: 0 } },
    { s: { r: 0, c: 1 }, e: { r: 0, c: 6 } },
    { s: { r: 0, c: 7 }, e: { r: 0, c: 12 } },
  ];
  workSheet["!cols"] = [{ wch: 10 }, ...Array(12).fill({ wpx: 30 })];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, workSheet, "2024 Sales");
  XLSX.writeFile(workbook, "Sales.xlsx");
}

const months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
const items = ["Monitors", "Speakers", "Keyboards", "Mouses", "Laptops"];

function generateSalesData() {
  return months.reduce(
    (prev, month) => ({ ...prev, [month]: Math.floor(Math.random() * 100) }),
    {}
  );
}

function getDatasource() {
  return items.map((item) => ({ item, ...generateSalesData() }));
}

function capitalizeFirstLetter(string: string) {
  return string.charAt(0).toUpperCase() + string.slice(1);
}

function App() {
  const datasource = getDatasource();
  const columns = [
    {
      title: "Item",
      dataIndex: "item",
      key: "item",
    },
    {
      title: "1st Half 2024",
      children: months.slice(0, 6).map((month) => ({
        title: capitalizeFirstLetter(month),
        dataIndex: month,
        key: month,
      })),
    },
    {
      title: "2nd Half 2024",
      children: months.slice(6).map((month) => ({
        title: capitalizeFirstLetter(month),
        dataIndex: month,
        key: month,
      })),
    },
  ];

  return (
    <main>
      <Table size="small" columns={columns} dataSource={datasource} pagination={false} bordered />
      <Button
        onClick={() => createExcel()}
        style={{ float: "right", marginTop: "12px" }}
        type="primary"
      >
        Export to Excel
      </Button>
    </main>
  );
}

export default App;
