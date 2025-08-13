<!--
/**
 * @fileoverview
 * # ExcelParser PCF Documentation
 *
 * ## Description
 * This PowerApps Component Framework (PCF) control enables users to upload, parse, and process Excel (.xlsx) files directly within Power Apps. It extracts worksheet data, performs basic validation, and outputs the parsed content for use in your app.
 *
 * ## Features
 * - Upload and parse Excel (.xlsx) files
 * - Extract data from selected worksheets
 * - Display parsed data in Power Apps
 * - Supports basic data validation
 *
 * ## Inputs and Outputs
 *
 * | Property Name      | Type     | Direction | Description                                      |
 * |--------------------|----------|-----------|--------------------------------------------------|
 * | fileUpload         | File     | Input     | The Excel (.xlsx) file to be uploaded and parsed. |
 * | worksheetSelection | String   | Input     | Name of the worksheet to extract data from.       |
 * | parsedData         | JSON     | Output    | Parsed worksheet data as a JSON object.           |
 * | parsedTable        | Table    | Output    | Parsed worksheet data as a table (tabular format).|
 *
 * ## Usage
 * 1. Add the ExcelParser control to your Power App.
 * 2. Configure input properties (file upload, worksheet selection).
 * 3. Use the output properties (parsedData, parsedTable) in your app logic.
 *
 * ## Author
 * Your Name
 *
 * @see [Repository](https://github.com/your-org/ExcelParser-PCF)
 * @license MIT
 */
-->
# ExcelParser PCF 📊

A PowerApps Component Framework (PCF) control for parsing and processing Excel files within Power Apps.

## ✨ Features

- 📁 Upload and parse Excel (.xlsx) files
- 📄 Extract data from worksheets
- 👀 Display parsed data in Power Apps
- ✅ Supports basic data validation

## 🚀 Getting Started

### 🛠 Prerequisites

- 🟢 Node.js (v14+)
- 🛠 Power Apps CLI (`pac`)
- 🌐 Microsoft Power Platform environment

### 📦 Installation

1. Clone the repository:
  ```bash
  git clone https://github.com/your-org/ExcelParser-PCF.git
  cd ExcelParser-PCF
  ```

2. Install dependencies:
  ```bash
  npm install
  ```

3. Build the PCF control:
  ```bash
  npm run build
  ```

4. Test locally:
  ```bash
  npm start
  ```

5. Deploy to Power Apps:
  ```bash
  pac pcf push --publisher-prefix <your-prefix>
  ```

## 📝 Usage

1. ➕ Add the ExcelParser control to your Power App.
2. ⚙️ Configure properties as needed.
3. 📤 Upload an Excel file to parse and display its contents.

## ⚙️ Configuration

- **Input Properties:** 📁 File upload, 📝 worksheet selection, etc.
- **Output Properties:** 🧾 Parsed data as JSON or table.

## 🔄 Inputs and Outputs

| Property Name      | Type     | Direction | Description                                      |
|--------------------|----------|-----------|--------------------------------------------------|
| 📁 `fileUpload`         | File     | Input     | The Excel (.xlsx) file to be uploaded and parsed. |
| 📝 `worksheetSelection` | String   | Input     | Name of the worksheet to extract data from.       |
| 🧾 `parsedData`         | JSON     | Output    | Parsed worksheet data as a JSON object.           |
| 📊 `parsedTable`        | Table    | Output    | Parsed worksheet data as a table (tabular format).|

## 🤝 Contributing

Contributions are welcome! Please open issues or submit pull requests.

## 📄 License

[MIT](LICENSE)

---

**Author:** Your Name  
**Repository:** [GitHub Link](https://github.com/your-org/ExcelParser-PCF)