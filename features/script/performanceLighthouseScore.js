const { Builder } = require('selenium-webdriver');
const { By } = require('selenium-webdriver');
const { exec } = require('child_process');
const fs = require('fs');
const xlsx = require('xlsx');
const path = require('path');
const ExcelJS = require('exceljs');

// Function to run tests
async function runTests(environment) {
    const urls = readURLsFromExcel('urls.xlsx', environment);

    if (urls.length === 0) {
        console.log(`No URLs found for environment ${environment}. Skipping test execution.`);
        return;
    }

    const driver = await new Builder().forBrowser('chrome').build();

    const devices = ['desktop', 'mobile', 'tablet'];
    const currentDateTime = new Date();
    const formattedDate = currentDateTime.toLocaleDateString('en-US', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).replace(/\//g, '-');
    const formattedTime = currentDateTime.toLocaleTimeString('en-US', {
        hour: '2-digit',
        minute: '2-digit',
        hour12: true
    }).replace(/:/g, '.').replace(' ', '');
    const mainDir = path.resolve(__dirname, '../../');
    const baseDir = path.join(mainDir, 'Performance Test Result');
    const environmentDir = path.join(baseDir, environment);
    const testDir = path.join(environmentDir, `${formattedDate} ${formattedTime} ${environment} Performance Testing`);

    fs.mkdirSync(testDir, { recursive: true });

    const jsonBaseDir = path.join(testDir, 'JSON');
    const htmlBaseDir = path.join(testDir, 'HTML');

    ['desktop', 'mobile', 'tablet'].forEach(deviceType => {
        fs.mkdirSync(path.join(htmlBaseDir, deviceType), { recursive: true });
        fs.mkdirSync(path.join(jsonBaseDir, deviceType), { recursive: true });
    });

    try {
        for (const url of urls) {
            const sanitizedURL = sanitizeURL(url);
            for (const device of devices) {
                const jsonReportPath = path.join(jsonBaseDir, device, `${sanitizedURL}.json`);
                const htmlDeviceDir = path.join(htmlBaseDir, device);
                try {
                    await driver.get(url);
                    const { htmlReportPath } = await runLighthouseAudit(url, jsonReportPath, htmlDeviceDir, device);
                    const scores = await extractScoresFromHTML(htmlReportPath, driver);
                    if (scores) {
                        await updateExcelSummary(url, environment, device, scores, testDir);
                    }
                } catch (auditError) {
                    console.error(`Lighthouse audit failed for URL: ${url} on ${device}, continuing with next URL.`);
                }
            }
        }
    } catch (error) {
        console.error('Error during test execution:', error);
    } finally {
        await driver.quit();
    }
}

async function updateExcelSummary(url, environment, deviceType, scores, testDir) {
    try {
        const currentDateTime = new Date();
        const formattedDate = currentDateTime.toLocaleDateString('en-US', {
            year: 'numeric',
            month: '2-digit',
            day: '2-digit',
        }).replace(/\//g, '-');
        const formattedTime = currentDateTime.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit',
            hour12: true
        }).replace(/:/g, '.').replace(' ', '');
        const fileName = `${environment} ${formattedDate} LH Audit report.xlsx`;
        const filePath = path.join(testDir, fileName);
        console.log(`Excel file path: ${filePath}`);

        let workbook = new ExcelJS.Workbook();
        let worksheet;

        if (fs.existsSync(filePath)) {
            console.log('Reading existing workbook');
            workbook = await workbook.xlsx.readFile(filePath);
            worksheet = workbook.getWorksheet(`${environment} ${formattedDate} LH report`);
        } else {
            console.log('Creating new workbook');
            worksheet = workbook.addWorksheet(`${environment} ${formattedDate} LH report`);

            // Merging cells for environment and title
            worksheet.mergeCells('A1:M1');
            worksheet.mergeCells('A2:M2');
            worksheet.getCell('A1').value = `Environment: ${environment}`;
            worksheet.getCell('A2').value = `Lighthouse ${formattedDate}`;

            // Styling merged cells
            ['A1', 'A2'].forEach(cell => {
                worksheet.getCell(cell).font = { size: 14, bold: true };
                worksheet.getCell(cell).alignment = { horizontal: 'left', vertical: 'middle' };
            });

            // Adding headers
            worksheet.getRow(4).values = ['URL', 'Desktop', '', '', '', 'Mobile', '', '', '', 'Tablet', '', '', ''];
            worksheet.getRow(5).values = ['', 'Performance', 'Accessibility', 'Best Practices', 'SEO', 'Performance', 'Accessibility', 'Best Practices', 'SEO', 'Performance', 'Accessibility', 'Best Practices', 'SEO'];

            // Styling headers
            const headerRow = worksheet.getRow(4);
            const subheaderRow = worksheet.getRow(5);

            // Set background color and text alignment
            worksheet.getCell('A4').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: '000000' } };
            worksheet.getCell('A4').font = { size: 12, bold: true, color: { argb: 'FFFFFF' } };
            worksheet.getCell('A4').alignment = { horizontal: 'center', vertical: 'middle' };

            ['B4', 'F4', 'J4'].forEach(cell => {
                worksheet.getCell(cell).alignment = { horizontal: 'center', vertical: 'middle' };
                worksheet.getCell(cell).font = { size: 12, bold: true, color: { argb: 'FFFFFF' } };
            });

            headerRow.eachCell((cell, colNumber) => {
                if (colNumber > 1) {
                    cell.font = { size: 12, bold: true, color: { argb: 'FFFFFF' } };
                    cell.alignment = { horizontal: 'center', vertical: 'middle' };
                    cell.fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: colNumber < 6 ? '4F81BD' : colNumber < 10 ? '92D050' : 'FFC000' }
                    };
                    cell.border = {
                        top: { style: 'thin' },
                        left: { style: 'thin' },
                        bottom: { style: 'thin' },
                        right: { style: 'thin' }
                    };
                }
            });

            subheaderRow.eachCell((cell, colNumber) => {
                cell.font = { size: 12, bold: true };
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'D9EAD3' }
                };
                cell.border = {
                    top: { style: 'thin' },
                    left: { style: 'thin' },
                    bottom: { style: 'thin' },
                    right: { style: 'thin' }
                };
            });

            // Adjust column widths
            worksheet.getColumn(1).width = 50; // URL
            worksheet.getColumn(2).width = 15; // Desktop - Performance
            worksheet.getColumn(3).width = 15; // Desktop - Accessibility
            worksheet.getColumn(4).width = 15; // Desktop - Best Practices
            worksheet.getColumn(5).width = 15; // Desktop - SEO
            worksheet.getColumn(6).width = 15; // Mobile - Performance
            worksheet.getColumn(7).width = 15; // Mobile - Accessibility
            worksheet.getColumn(8).width = 15; // Mobile - Best Practices
            worksheet.getColumn(9).width = 15; // Mobile - SEO
            worksheet.getColumn(10).width = 15; // Tablet - Performance
            worksheet.getColumn(11).width = 15; // Tablet - Accessibility
            worksheet.getColumn(12).width = 15; // Tablet - Best Practices
            worksheet.getColumn(13).width = 15; // Tablet - SEO
        }

        // Ensure cells are not already merged before merging
        const mergeIfNotMerged = (cellRange) => {
            const [start, end] = cellRange.split(':');
            const startCell = worksheet.getCell(start);
            const endCell = worksheet.getCell(end);
            if (!(startCell.isMerged && endCell.isMerged)) {  
                worksheet.mergeCells(cellRange);
            }
        };

        // Merge cells for URL header
        mergeIfNotMerged('A4:A5');

        // Function to merge columns based on their background color
        const mergeColumnsByColor = (startCol, endCol) => {
            let startMerge = false;
            let mergeStartCol = '';
            let mergeEndCol = '';
            for (let col = startCol; col <= endCol; col++) {
                const cell = worksheet.getCell(getExcelCellRef(4, col));
                if (cell.fill.fgColor.argb === worksheet.getCell(getExcelCellRef(4, startCol)).fill.fgColor.argb) {
                    if (!startMerge) {
                        startMerge = true;
                        mergeStartCol = getExcelCellRef(4, col);
                    }
                } else {
                    if (startMerge) {
                        mergeEndCol = getExcelCellRef(4, col - 1);
                        mergeIfNotMerged(`${mergeStartCol}:${mergeEndCol}`);
                    }
                    startMerge = false;
                    mergeStartCol = '';
                    mergeEndCol = '';
                }
            }
            if (startMerge) {
                mergeEndCol = getExcelCellRef(4, endCol);
                mergeIfNotMerged(`${mergeStartCol}:${mergeEndCol}`);
            }
        };

        // Helper function to get Excel cell reference
        const getExcelCellRef = (row, col) => {
            return `${String.fromCharCode(64 + col)}${row}`;
        };

        // Merge columns based on their background color
        mergeColumnsByColor(2, 5);   // Desktop columns (Performance to SEO)
        mergeColumnsByColor(6, 9);   // Mobile columns (Performance to SEO)
        mergeColumnsByColor(10, 13); // Tablet columns (Performance to SEO)

        // Find if the URL already exists
        let rowIndex = 6;
        let existingRow;
        while (worksheet.getRow(rowIndex).getCell(1).value) {
            if (worksheet.getRow(rowIndex).getCell(1).text === url) {
                existingRow = worksheet.getRow(rowIndex);
                break;
            }
            rowIndex++;
        }

        // Initialize or update row values
        let rowValues = existingRow ? existingRow.values : [{ text: url, hyperlink: url }, '', '', '', '', '', '', '', '', '', '', '', ''];

        // Set scores based on the device type
        if (deviceType === 'desktop') {
            rowValues[1] = scores.performanceScore;      // Performance
            rowValues[2] = scores.accessibilityScore;    // Accessibility
            rowValues[3] = scores.bestPracticesScore;    // Best Practices
            rowValues[4] = scores.seoScore;              // SEO
        } else if (deviceType === 'mobile') {
            rowValues[6] = scores.performanceScore;
            rowValues[7] = scores.accessibilityScore;
            rowValues[8] = scores.bestPracticesScore;
            rowValues[9] = scores.seoScore;
        } else if (deviceType === 'tablet') {
            rowValues[10] = scores.performanceScore;
            rowValues[11] = scores.accessibilityScore;
            rowValues[12] = scores.bestPracticesScore;
            rowValues[13] = scores.seoScore;
        }

        // Add or update data in worksheet
        if (existingRow) {
            existingRow.values = rowValues;
        } else {
            worksheet.getRow(rowIndex).values = rowValues;
        }

        // Remove extra rows after URL link if any
        if (worksheet.getRow(rowIndex + 1).getCell(1).value === null) {
            worksheet.spliceRows(rowIndex + 1, 1);
        }

        // Style the new or updated row
        const newRow = worksheet.getRow(rowIndex);
        newRow.height = 20; // Adjust row height for better spacing
        newRow.eachCell((cell, colNumber) => {
            // Setting cell borders
            cell.border = {
                top: { style: 'thin' },
                left: { style: 'thin' },
                bottom: { style: 'thin' },
                right: { style: 'thin' }
            };

            // Setting font styles
            cell.font = { size: 10 };

            // Aligning cell text
            cell.alignment = { horizontal: 'center', vertical: 'middle' };

            // Color coding for the scores based on thresholds
            if (colNumber > 1 && cell.value) {
                let score = cell.value;
                if (score >= 90) {
                    cell.font.color = { argb: '008000' }; // Green
                } else if (score >= 50) {
                    cell.font.color = { argb: 'FFA500' }; // Orange
                } else {
                    cell.font.color = { argb: 'FF0000' }; // Red
                }
            }
        });

        // Set URL cell to hyperlink with blue color and underline
        worksheet.getCell(`A${rowIndex}`).alignment = { horizontal: 'left', vertical: 'middle' };
        worksheet.getCell(`A${rowIndex}`).font = { color: { argb: '0000FF' }, underline: true };

        await workbook.xlsx.writeFile(filePath);
        console.log(`Excel file has been updated: ${filePath}`);
    } catch (error) {
        console.error('Error updating Excel summary:', error);
    }
}

async function extractScoresFromHTML(htmlReportPath, driver) {
    try {
        await driver.get(`file://${htmlReportPath}`);
        
        const performanceScore = await driver.findElement(By.xpath('/html/body/article/div[2]/div[2]/div/div/div/div[2]/a[1]/div[2]')).getText();
        const accessibilityScore = await driver.findElement(By.xpath('/html/body/article/div[2]/div[2]/div/div/div/div[2]/a[2]/div[2]')).getText();
        const bestPracticesScore = await driver.findElement(By.xpath('/html/body/article/div[2]/div[2]/div/div/div/div[2]/a[3]/div[2]')).getText();
        const seoScore = await driver.findElement(By.xpath('/html/body/article/div[2]/div[2]/div/div/div/div[2]/a[4]/div[2]')).getText();

        return {
            performanceScore: parseFloat(performanceScore),
            accessibilityScore: parseFloat(accessibilityScore),
            bestPracticesScore: parseFloat(bestPracticesScore),
            seoScore: parseFloat(seoScore)
        };
    } catch (error) {
        console.error(`Error extracting scores from HTML report: ${error}`);
        return null;
    }
}

// Function to read URLs from the Excel file based on the environment sheet
function readURLsFromExcel(filePath, sheetName) {
    const workbook = xlsx.readFile(filePath);
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
        throw new Error(`Sheet ${sheetName} does not exist`);
    }
    const data = xlsx.utils.sheet_to_json(sheet);
    const urls = data.map(row => row['URL Link']);

    if (urls.length === 0) {
        console.log(`Missing URL link on the environment ${sheetName}. Please add to proceed.`);
        // You might choose to throw an error here to stop execution or handle it as needed.
    }

    return urls;
}

// Function to sanitize URL to create a valid filename
function sanitizeURL(url) {
    return url.replace(/[^a-z0-9]/gi, '_').toLowerCase();
}

// Function to run Lighthouse audit
function runLighthouseAudit(url, jsonReportPath, htmlBaseDir, deviceType, retries = 3) {
    return new Promise((resolve, reject) => {
        const absoluteJsonReportPath = path.resolve(jsonReportPath);

        let config;
        if (deviceType === 'desktop') {
            config = '--preset=desktop';
        } else if (deviceType === 'mobile') {
            config = '--emulated-form-factor=mobile';
        } else if (deviceType === 'tablet') {
            config = '--emulated-form-factor=mobile --screenEmulation.width=768 --screenEmulation.height=1024 --screenEmulation.deviceScaleFactor=2';
        }

        const sanitizedURL = sanitizeURL(url);

        const runAudit = (retryCount) => {
            console.log(`Running Lighthouse for URL: ${url} on ${deviceType} (Retry ${retryCount})`);
            exec(`npx lighthouse ${url} --output=json --output-path="${absoluteJsonReportPath}" ${config} --max-wait-for-load 45000`, (error, stdout, stderr) => {
                if (error) {
                    console.error(`Error executing Lighthouse JSON report for URL ${url}: ${error.message}`);
                    console.error(`stderr: ${stderr}`);
                    if (retryCount < retries) {
                        runAudit(retryCount + 1);
                    } else {
                        reject(error);
                        return;
                    }
                } else {
                    const htmlReportPath = path.join(htmlBaseDir, `${sanitizedURL}.html`);
                    exec(`npx lighthouse ${url} --output=html --output-path="${htmlReportPath}" ${config} --max-wait-for-load 45000`, (error, stdout, stderr) => {
                        if (error) {
                            console.error(`Error executing Lighthouse HTML report for URL ${url}: ${error.message}`);
                            console.error(`stderr: ${stderr}`);
                            if (retryCount < retries) {
                                runAudit(retryCount + 1);
                            } else {
                                reject(error);
                                return;
                            }
                        } else {
                            console.log(`Lighthouse audit completed for URL ${url} on ${deviceType}`);
                            console.log(`stdout: ${stdout}`);
                            resolve({ jsonReportPath: absoluteJsonReportPath, htmlReportPath });
                        }
                    });
                }
            });
        };

        runAudit(0);
    });
}

// Determine the environment from the command line argument
const environment = process.argv[2]; // Pass environment as command line argument
if (!environment) {
    console.error('Please provide an environment (DEV, SIT, UAT, PROD) as an argument');
    process.exit(1);
}

runTests(environment);