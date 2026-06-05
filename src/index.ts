import { mkdir, readFile, writeFile, copyFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import ExcelJS from 'exceljs';
import { EfmsClient } from './client.js';
import { getAccessToken } from './auth.js';
import { loadConfig } from './config.js';
import {
  downloadFileFromOneDrive,
  isOneDriveNotFoundError,
  uploadFileToOneDrive,
  uploadFileToOneDrivePath,
} from './onedrive.js';

export { EfmsClient } from './client.js';
export { getAccessToken } from './auth.js';
export { loadConfig } from './config.js';

const isDirectRun = import.meta.url === pathToFileURL(process.argv[1] ?? '').href;
const dataDir = join(process.cwd(), 'data');
const api1OutputPath = join(dataDir, 'api1-response.json');
const api2OutputPath = join(dataDir, 'api2-response.json');
const templateWorkbookPath = join(dataDir, 'Vendor_Payment_Template.xlsx');
const outputWorkbookPath = join(dataDir, 'Vendor_Payment_Output.xlsx');
const envPath = join(process.cwd(), '.env');
const intervalMs = 30 * 60 * 1000;
const worksheetName = 'VENDOR_PAYMENT';
const remoteWorkbookName = 'VENDOR_PAYMENT.xlsx';

type Api1ResponseFile = {
  meta: {
    lastRunAt: string;
    accessTokenPreview: string;
  };
  response: unknown;
};

type Api2ResponseFile = {
  meta: {
    lastRunAt: string;
    nextRunAt: string;
    pageNumber: number;
    pageSize: number;
  };
  response: unknown;
};

type SettlementPayment = {
  id?: string;
  settlementNo?: string | null;
  payeeName?: string | null;
  requestDate?: string | null;
  dueDate?: string | null;
  datetimeModified?: string | null;
  [key: string]: unknown;
};

async function ensureDataDir() {
  await mkdir(dataDir, { recursive: true });
}

async function saveApi1Response(raw: unknown, token: string) {
  await ensureDataDir();
  const payload: Api1ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      accessTokenPreview: `${token.slice(0, 6)}...${token.slice(-4)}`,
    },
    response: raw,
  };

  await writeFile(api1OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi2Response(data: unknown) {
  await ensureDataDir();

  const now = new Date();
  const payload: Api2ResponseFile = {
    meta: {
      lastRunAt: now.toISOString(),
      nextRunAt: new Date(now.getTime() + intervalMs).toISOString(),
      pageNumber: 1,
      pageSize: 1000,
    },
    response: data,
  };

  await writeFile(api2OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

function iteratePaymentsFromLastToFirst(payments: SettlementPayment[]): SettlementPayment[] {
  return [...payments].reverse();
}

function getVietnamTimestamp() {
  return new Intl.DateTimeFormat('vi-VN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
    timeZone: 'Asia/Ho_Chi_Minh',
  }).format(new Date());
}

function logStep(step: string, status: string) {
  console.log(`${getVietnamTimestamp()} | [${step}] ${status}`);
}

async function ensureWorkbookExists() {
  try {
    await readFile(outputWorkbookPath);
  } catch {
    await copyFile(templateWorkbookPath, outputWorkbookPath);
  }
}

async function resetWorkbookFromTemplate() {
  await ensureDataDir();
  await copyFile(templateWorkbookPath, outputWorkbookPath);
}

async function updateEnvValue(key: string, value: string) {
  let contents = '';

  try {
    contents = await readFile(envPath, 'utf8');
  } catch {
    contents = '';
  }

  const line = `${key}=${value}`;
  const pattern = new RegExp(`^${key}=.*$`, 'm');
  const nextContents = pattern.test(contents)
    ? contents.replace(pattern, line)
    : `${contents.trimEnd()}\n${line}\n`;

  await writeFile(envPath, nextContents, 'utf8');
}

async function upsertPaymentsInWorkbook(payments: SettlementPayment[]) {
  await ensureDataDir();
  await ensureWorkbookExists();

  const rows = payments
    .map((payment) => ({
      i: payment.settlementNo ?? '',
    }))
    .filter((row) => row.i);

  if (!rows.length) {
    return { added: 0, updated: 0 };
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbookPath);
  const worksheet = workbook.getWorksheet(worksheetName);

  if (!worksheet) {
    throw new Error(`Worksheet not found: ${worksheetName}`);
  }

  const rowsBySettlementNo = new Map<string, ExcelJS.Row>();
  worksheet.eachRow((row) => {
    const settlementNo = String(row.getCell(9).value ?? '').trim();
    if (settlementNo) {
      rowsBySettlementNo.set(settlementNo, row);
    }
  });

  let added = 0;
  let updated = 0;

  for (const row of rows) {
    const settlementNo = row.i.trim();
    const existingRow = rowsBySettlementNo.get(settlementNo);

    if (!existingRow) {
      worksheet.addRow([null, null, null, null, null, null, null, null, row.i]);
      rowsBySettlementNo.set(settlementNo, worksheet.lastRow!);
      added += 1;
      continue;
    }

    // Existing rows are matched by settlementNo only. No other columns are managed.
  }

  if (added || updated) {
    await workbook.xlsx.writeFile(outputWorkbookPath);
  }

  return { added, updated };
}

async function prepareWorkbookFromOneDrive(config: NonNullable<ReturnType<typeof loadConfig>['oneDrive']>) {
  await ensureDataDir();

  try {
    await downloadFileFromOneDrive(outputWorkbookPath, config);
    return { createdFromTemplate: false };
  } catch (error) {
    if (!isOneDriveNotFoundError(error)) {
      throw error;
    }

    await resetWorkbookFromTemplate();
    return { createdFromTemplate: true };
  }
}

async function runOnce() {
  logStep('start', 'bắt đầu');
  const config = loadConfig();
  logStep('config', 'đã tải cấu hình');

  const { token, raw } = await getAccessToken();
  logStep('access-token', 'lấy thành công');
  await saveApi1Response(raw, token);
  logStep('api1', 'đã lưu API 1');

  const client = new EfmsClient(config.baseUrl);
  logStep('request', 'đã bắt đầu gửi yêu cầu');

  const data = await client.postJson<{ requester: string }, { data: SettlementPayment[] }>(
    '/Accounting/api/v1/en-US/AcctSettlementPayment/paging?pageNumber=1&pageSize=1000',
    {
      requester: config.settlementPaymentRequester,
    },
    {
      auth: false,
      headers: {
        Authorization: `Bearer ${token}`,
      },
    },
  );

  await saveApi2Response(data);
  logStep('api2', 'đã lưu API 2');

  const payments = iteratePaymentsFromLastToFirst(data.data ?? []);

  if (payments.length) {
    let createdWorkbookFromTemplate = false;

    if (config.oneDrive) {
      const workbook = await prepareWorkbookFromOneDrive(config.oneDrive);
      createdWorkbookFromTemplate = workbook.createdFromTemplate;
      logStep(
        'onedrive',
        createdWorkbookFromTemplate
          ? 'không tìm thấy workbook, đã tạo mới từ template'
          : 'đã tải workbook hiện tại',
      );
    }

    const result = await upsertPaymentsInWorkbook(payments);
    logStep('workbook', `thêm ${result.added} dòng, cập nhật ${result.updated} dòng`);

    if (config.oneDrive && (result.added || result.updated)) {
      if (createdWorkbookFromTemplate) {
        const uploadedFile = await uploadFileToOneDrivePath(outputWorkbookPath, config.oneDrive, remoteWorkbookName);

        if (!uploadedFile.id) {
          throw new Error('OneDrive path upload did not return a new file id');
        }

        await updateEnvValue('ONEDRIVE_FILE_ID', uploadedFile.id);
        logStep('onedrive', `đã cập nhật ONEDRIVE_FILE_ID=${uploadedFile.id}`);
      } else {
        await uploadFileToOneDrive(outputWorkbookPath, config.oneDrive);
      }
      logStep('onedrive', 'đã cập nhật workbook');
    }
  } else {
    logStep('workbook', 'không có dòng mới');
  }

  logStep('request', 'hoàn tất');
}

async function startScheduler() {
  while (true) {
    try {
      await runOnce();
    } catch (error) {
      logStep('run-failed', 'thất bại');
      console.error(error);
      process.exitCode = 1;
    }

    logStep('waiting-next-run', `đang chờ ${intervalMs / 60000} phút để chạy lại`);
    await new Promise((resolve) => setTimeout(resolve, intervalMs));
  }
}

if (isDirectRun) {
  if (process.argv.includes('--once')) {
    void runOnce().catch((error) => {
      logStep('run-failed', 'thất bại');
      console.error(error);
      process.exitCode = 1;
    });
  } else {
    void startScheduler();
  }
}
