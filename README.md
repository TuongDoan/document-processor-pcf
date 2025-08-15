# Document Processor PCF ⚡️📄

A Power Apps Component Framework (PCF) control for **fast, local Excel parsing** and **multi-file upload**—no slow SharePoint handoffs or long Power Automate loops. 🚀



## ✨ Why this exists

A common Excel parsing workflow: Upload the file to SharePoint → wait for Power Automate to trigger → iterate through a formatted table.

That approach works, but **slow**, **inefficient**, and **large tables often stall**. It doesn’t have to be like that. 

This **Document Processor PCF** extracts tables locally, lightning fast and returns clean JSON you can immediately use in your app or send to downstream Power Automate flows, AI Builder, agents, or other services. ⚙️


---


## 🧠 What it does

### 📊 Excel parsing modes

**📍 Fixed Range Mode**
  - Parse a specific, predefined cell range (e.g., `Sheet1!A1:B2`).
  - Optional **AutoHeader**: promote the first row to headers when your data has none.

**🧱 Fixed Column Mode**
  - Provide a starting header range (e.g., `Sheet1!B2:Z2`). The parser reads columns **B → Z** starting at **row 2**, and continues **until the first blank row**.
  - Ideal for standardized templates (e.g., purchase orders with variable row counts).

---

**🔎 Dynamic Search Mode**
  - Scans **all worksheets** (across multiple files) to detect **every table**.
  - A “table” = any **continuous** block of cells with **no blank rows or columns** interrupting it.
    > Originally a tool of my MCP server, this lets you extract all tables and hand them to AI Builder, or your remote agent for deeper analysis and richer context. 🤝

<img width="1101" height="454" alt="dynamic" src="https://github.com/user-attachments/assets/1b00bdbe-ac4a-4a60-9f71-6703dc8eaffb" />
Dynamic mode parses all tables across all sheets

---


## 🧰 Other features

- 🗂️ **Multiple file upload** with drag & drop  
- 🖱️ **Enhanced button UX** (multi-line caption, icon options)  
- 🧩 **Extension limiting** to control allowed file types  
- 🎨 **Theme inheritance** from Power Apps  
- 🌀 **Progress spinner** and optional secondary text for better feedback

<img width="698" height="168" alt="button" src="https://github.com/user-attachments/assets/e5f9081d-9754-4905-bdbb-d563421f4f6c" />


---


## 🔌 Params

### 🔧 Inputs (key properties)

- **Text**: Button label.  
- **ButtonIcon**: `Attach | Add document | ArrowUpload`.  
- **IconStyle**: `Outline | Filled`.  
- **ParsingMode**: `Dynamic | Fixed range | Fixed column`.  
- **AutoHeader**: `true | false` (use with Fixed Range/Fixed Column when the dataset lacks headers).  
- **targetRange**:
  - **Fixed Range**: e.g., `Sheet1!A1:B2`.  
  - **Fixed Column**: header row span, e.g., `Sheet1!B2:Z2`.  
  - **Dynamic**: **ignored**.  
- **AllowMultipleFiles**: Enable multi-file selection.  
- **AllowedFileTypes**: Semicolon-separated extensions (e.g., `xlsx;xls`).  
- **AllowDropFiles** & **AllowDropFilesText**: Drag & drop and its hint text.  
- **ShowActionSpinner**: Show a spinner while processing.  
- **ShowSecondaryContent** & **SecondaryContent**: Helper text under the label.  
- **Appearance**: `Primary | Secondary | Outline | Subtle | Transparent`.  
- **Align**: `Left | Center | Right | Justify`.  
- **FontWeight**: `Bold | Lighter | Normal | Semibold`.  
- **IconPosition**: `Before | After`.  
- **Shape**: `Rounded | Circular | Square`.  
- **ButtonSize**: `Small | Medium | Large`.  
- **Width / Height / Visible**: Layout controls.

### 📤 Outputs

- **FilesAsJSON**: JSON array of selected files (Base64 and metadata).  
- **ExcelOutput**: JSON array of parsed table objects for Power Fx or downstream APIs/Agents.

> **Dependencies & notes**  
> - `targetRange` is **required** in **Fixed Range** and **Fixed Column** modes (format differs).  
> - `targetRange` is **not used** in **Dynamic** mode.  

---


## 🖼️ Base64 output (non-Excel too)

When you upload other file types, the control also returns an **array of Base64 strings** (one object per file) via `FilesAsJSON`.  
Great for:
- Sending images to **AI Vision** models. 👁️  
- Uploading to **Blob storage** or APIs without extra conversions. ☁️


---

## 🔐 Security & privacy

- Parsing happens **client-side** in the app session. 
- Although it was designed to dodge Formula Injection Attacks, you should validate and sanitize sensitive outputs before forwarding them to external services.
- Use **AllowedFileTypes** to restrict uploads in sensitive apps.


---

## 🏁 Quick start

1. **Import** the PCF control into your solution. 📦  
2. **Add** the control to your canvas app screen. 🧱  
3. **Configure**:
   - Choose **ParsingMode**.
   - **Fixed Range**: set `targetRange` like `Sheet1!A1:H50`.  
   - **Fixed Column**: set `targetRange` like `Sheet1!B2:Z2`.  
   - Toggle **AutoHeader** if needed.  
4. **Use outputs**:
   - Read **`ExcelOutput`** for structured tables.  
   - Read **`FilesAsJSON`** for Base64 content (e.g., AI Vision or storage).  

