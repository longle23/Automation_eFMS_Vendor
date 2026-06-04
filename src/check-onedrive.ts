import { loadConfig } from './config.js';
import { checkOneDriveFile } from './onedrive.js';

const config = loadConfig();

if (!config.oneDrive) {
  throw new Error('OneDrive is not configured in .env');
}

const fileName = await checkOneDriveFile(config.oneDrive);
console.log(`OneDrive connection OK. File accessible: ${fileName}`);
