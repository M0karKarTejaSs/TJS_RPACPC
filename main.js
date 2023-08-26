const { app, BrowserWindow, ipcMain } = require('electron');
const jetpack = require('fs-jetpack');
const path = require('path');
const axios = require('axios'); // Import the axios library
process.env.NODE_TLS_REJECT_UNAUTHORIZED = "0";
app.allowRendererProcessReuse = true;
app.commandLine.appendSwitch('ignore-certificate-errors');
app.commandLine.appendSwitch('allow-insecure-localhost');
let mainWindow;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 800,
        height: 600,
        webPreferences: {
            nodeIntegration: true
        }
    });

    mainWindow.loadFile('src/index.html');
    mainWindow.webContents.openDevTools();
}

app.whenReady().then(createWindow);

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
    }
});

ipcMain.on('file-selected', async (event, filePath) => {
    try {
        const fileName = path.basename(filePath);
        const baseName = path.basename(filePath, path.extname(filePath));
        const extension = path.extname(filePath);

        const uploadFolderPath = jetpack.path(__dirname, 'uploads');
        jetpack.dir(uploadFolderPath);

        const uniqueFileName = generateUniqueFileName(uploadFolderPath, baseName, extension);

        const destinationPath = jetpack.path(uploadFolderPath, uniqueFileName);
        jetpack.copy(filePath, destinationPath);

        console.log('File saved successfully');

        // Process the Excel data and fetch PAN details from API
        await processAndFetchPanData(destinationPath);
    } catch (err) {
        console.error(err);
    }
});

async function processAndFetchPanData(filePath) {
    const xlsx = require('xlsx');

    const workbook = xlsx.readFile(filePath);
    const firstSheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[firstSheetName];

    const rows = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

    for (let row of rows) {
        const panNumber = row[0]; // Update index based on your Excel layout

        if (panNumber) {
            console.log('Extracted PAN Number:', panNumber);

            try {
                const panDetails = await fetchPanDetailsFromApi(panNumber);
                console.log('PAN Details:', panDetails);
            } catch (error) {
                console.error('Error fetching PAN details:', error.message);
            }
        }
    }
}


async function fetchPanDetailsFromApi(panNumber) {
    const apiUrl = 'https://api.rpacpc.com/services/get-pan-details';
    const headers = {
        secretkey: 'e5b45d31-7a63-4dfa-a7f1-c8fc89192fe8',
        token: 'HZqJwTTU+6SnoILGiwfD2h6Lgpp977mCfFJ4+XrnVvUDKENPJ0WjgRGO0uv9NODrf7KjCl6d34LQJOvn8w/aih79BZHUU6zKzfcoQDLBHkAHaUceuj1AUFRwD6kdoXZLSaZofXaeNXH2P7bcfGvVjM0kW7VS3bmljOlKz0wC2K5lhXs5eeXuKK7IAIGPNoXeXqU8UTaJtdQk4B3N4sM9v/R/6zuvMSz2t6oJQRTj4geWs9nKW6StVxZk2JzwGR1bw2cqWh00lwXmOCKmOxNhdDmMfQBQVXtH6qrBX2FykV162zMzzFMIoBOxdqBdCq0abjZH+hzQpIBUlmFDAIJFS/XL3I3/Or5fD3wNvt4Il5MhZqYCIwIFg2yH9hNvbPQ7gaNvz1zsLf0CBrFqUiw9P2JV3laBgkKHj26ooq9cj8mEy2EEn4YduF3wNcnuuLrl'
    };

    try {
        const response = await axios.post(apiUrl, { pancard: panNumber }, {
            headers: {
                'Content-Type': 'application/json',
                ...headers
            }
        });

        if (response.status !== 200) {
            console.error('API Error Response:', {
                status: response.status,
                statusText: response.statusText,
                headers: response.headers,
                body: response.data // Response data
            });
            throw new Error('API Error');
        }

        return response.data;
    } catch (error) {
        console.error('Fetch Error:', error);
        throw new Error('API Error');
    }
}

function generateUniqueFileName(folderPath, baseName, extension) {
    let uniqueName = `${baseName}${extension}`;
    let count = 1;

    while (jetpack.exists(jetpack.path(folderPath, uniqueName))) {
        uniqueName = `${baseName}_${count}${extension}`;
        count++;
    }

    return uniqueName;
}
