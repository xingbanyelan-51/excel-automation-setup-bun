# Excel Automation Setup with Bun

[![GitHub license](https://img.shields.io/github/license/xingbanyelan-51/excel-automation-setup-bun)](https://github.com/xingbanyelan-51/excel-automation-setup-bun/blob/main/LICENSE)
[![GitHub stars](https://img.shields.io/github/stars/xingbanyelan-51/excel-automation-setup-bun)](https://github.com/xingbanyelan-51/excel-automation-setup-bun/stargazers)
[![GitHub issues](https://img.shields.io/github/issues/xingbanyelan-51/excel-automation-setup-bun)](https://github.com/xingbanyelan-51/excel-automation-setup-bun/issues)

A user-friendly PowerShell script to bootstrap a Bun-based environment for automating Excel file processing. It installs necessary dependencies, generates a template `main.js` script (only if it doesn't exist), and allows interactive processing of Excel files via drag-and-drop in a loop.

This setup is perfect for developers or data enthusiasts who want to read, manipulate, and process Excel files using JavaScript libraries in a command-line workflow.

## Features
- **Automatic Setup**: Installs Bun (if missing), initializes a project, and adds essential Excel libraries (`xlsx`, `exceljs`, `xlsx-populate`).
- **Template Generation**: Creates a starter `main.js` file with import examples (won't overwrite existing files).
- **Interactive Loop**: Prompts for Excel file paths repeatedly, processes them with `bun run main.js`, and suggests rewriting the script for custom logic.
- **Error Handling**: Graceful error catching and console clearing for a clean user experience.
- **Cross-Compatible**: Works on Windows with PowerShell 7.5.4+; focuses on local, offline execution.

## Prerequisites
- Windows OS (tested on Windows 10/11).
- PowerShell 7.5.4 or higher (install from [Microsoft's official site](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows)).
- Internet access for initial Bun installation and npm package downloads (offline after setup).

## Installation
1. **Clone the Repository**:
   ```
   git clone https://github.com/xingbanyelan-51/excel-automation-setup-bun.git
   cd excel-automation-setup-bun
   ```

2. **Run the Script**:
   - Double-click `Excel Automate.ps1` (or run via PowerShell: `.\Excel Automate.ps1`).
   The script will handle everything: install Bun, dependencies, and generate `main.js` if needed.

3. **Customize**:
   - Edit `main.js` to add your Excel processing logic (e.g., reading sheets, modifying data).
   - Re-run the script to test with your Excel files.

## Usage
1. Double-click or run the `.ps1` script.
2. If it's the first run:
   - Bun and dependencies are installed automatically.
   - A template `main.js` is created.
3. Drag-and-drop an Excel file (or paste its path) when prompted.
4. The script runs `bun run main.js` on the file and outputs results to the console.
5. After processing, it prompts: "处理完成，请手动重写main.js，并再次拖入EXCEL文件：" (Process complete, please manually rewrite main.js and drag in another Excel file).
6. Leave the path blank to exit the loop.
7. Press Enter to close the window when done.

### Example Output
![Script Running Screenshot](https://via.placeholder.com/800x400?text=Script+Running+Screenshot)  
*(Replace with an actual screenshot of the script in action.)*

### Customizing main.js
The generated `main.js` template looks like this:
```javascript
import XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import XlsxPopulate from 'xlsx-populate';

console.log('All libraries imported successfully!');

const filePath = process.argv[2]?.trim();
if (!filePath) {
    console.log('Usage: bun run main.js "C:\\path\\to\\your\\file.xlsx"');
    process.exit(1);
}

try {
    const workbook = XLSX.readFile(filePath);
    console.log(`SheetJS success! Sheets: ${workbook.SheetNames.join(', ')}`);
    
    // Example with ExcelJS: Load workbook
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.readFile(filePath);
    console.log('ExcelJS loaded successfully.');
} catch (err) {
    console.error('Failed to process Excel file:', err.message);
}

console.log('请重写main.js脚本');
```

Add your logic here, e.g., data extraction, modifications, or exports.

## Dependencies
- [Bun](https://bun.sh/) - Fast JavaScript runtime.
- [xlsx](https://www.npmjs.com/package/xlsx) - Excel parser/writer (SheetJS).
- [exceljs](https://www.npmjs.com/package/exceljs) - Advanced Excel read/write with styling.
- [xlsx-populate](https://www.npmjs.com/package/xlsx-populate) - Template-based Excel population.

All installed automatically by the script.

## Contributing
Contributions are welcome! Please follow these steps:
1. Fork the repository.
2. Create a feature branch (`git checkout -b feature/YourFeature`).
3. Commit your changes (`git commit -m 'Add YourFeature'`).
4. Push to the branch (`git push origin feature/YourFeature`).
5. Open a Pull Request.

Report issues or suggest improvements via [GitHub Issues](https://github.com/xingbanyelan-51/excel-automation-setup-bun/issues).

Note: This project is licensed under GPLv3, so any contributions must comply with the license terms.

## License
This project is licensed under the GNU General Public License v3.0 - see the [LICENSE](LICENSE) file for details.

## Acknowledgments
- Inspired by community discussions on Excel automation with JavaScript.
- Thanks to the maintainers of Bun, xlsx, exceljs, and xlsx-populate.

---

If you find this useful, star the repo or share it! Questions? Open an issue.
