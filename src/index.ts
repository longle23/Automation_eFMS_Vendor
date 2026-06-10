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
const api3OutputPath = join(dataDir, 'api3-response.json');
const api4OutputPath = join(dataDir, 'api4-response.json');
const api2StatePath = join(dataDir, 'api2-state.json');
const templateWorkbookPath = join(dataDir, 'Vendor_Payment_Template.xlsx');
const outputWorkbookPath = join(dataDir, 'Vendor_Payment_Output.xlsx');
const envPath = join(process.cwd(), '.env');
const intervalMs = 30 * 60 * 1000;
const worksheetName = 'VENDOR_PAYMENT';
const legacyWorksheetName = 'Vendor_Payment';
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

type Api3ResponseFile = {
  meta: {
    lastRunAt: string;
    settlementCount: number;
  };
  response: Record<string, unknown>;
};

type Api4ResponseFile = {
  meta: {
    lastRunAt: string;
    settlementCount: number;
    settlementIds: string[];
  };
  response: Record<string, unknown>;
};

type SettlementPayment = {
  id?: string;
  settlementNo?: string | null;
  payeeAccountNo?: string | null;
  payeeName?: string | null;
  requestDate?: string | null;
  dueDate?: string | null;
  invoiceDate?: string | null;
  datetimeModified?: string | null;
  [key: string]: unknown;
};

type SettlementApprovalInfo = {
  requesterAprDate?: string | null;
  managerAprDate?: string | null;
  accountantAprDate?: string | null;
  [key: string]: unknown;
};

type Api2StateFile = {
  meta: {
    lastRunAt: string;
    itemCount: number;
  };
  items: Record<string, string>;
};

type SyncResult = {
  added: number;
  updated: number;
  unchanged: number;
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

async function saveApi3Response(data: Record<string, unknown>) {
  await ensureDataDir();

  const payload: Api3ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      settlementCount: Object.keys(data).length,
    },
    response: data,
  };

  await writeFile(api3OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi4Response(data: Record<string, unknown>, settlementIds: string[]) {
  await ensureDataDir();

  const payload: Api4ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      settlementCount: Object.keys(data).length,
      settlementIds,
    },
    response: data,
  };

  await writeFile(api4OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

function extractChargeNoGrpSettlement(response: unknown) {
  if (!response || typeof response !== 'object') {
    return null;
  }

  const record = response as {
    data?: { chargeNoGrpSettlement?: unknown };
    result?: { chargeNoGrpSettlement?: unknown };
    chargeNoGrpSettlement?: unknown;
  };

  const chargeNoGrpSettlement =
    record.data?.chargeNoGrpSettlement ??
    record.result?.chargeNoGrpSettlement ??
    record.chargeNoGrpSettlement ??
    [];

  if (!Array.isArray(chargeNoGrpSettlement)) {
    return chargeNoGrpSettlement;
  }

  const firstItem = chargeNoGrpSettlement[0] ?? null;

  if (!firstItem || typeof firstItem !== 'object') {
    return firstItem;
  }

  return firstItem;
}

function iteratePaymentsFromLastToFirst(payments: SettlementPayment[]): SettlementPayment[] {
  return [...payments].reverse();
}

function stringifyPaymentSignature(payment: SettlementPayment) {
  const normalized = Object.fromEntries(
    Object.entries(payment)
      .filter(([, value]) => value !== undefined)
      .sort(([a], [b]) => a.localeCompare(b)),
  );

  return JSON.stringify(normalized);
}

function getPaymentId(payment: SettlementPayment) {
  return String(payment.id ?? '').trim();
}

function getPaymentSettlementNo(payment: SettlementPayment) {
  return String(payment.settlementNo ?? '').trim();
}

function getUniqueSettlementNos(payments: SettlementPayment[]) {
  return [...new Set(payments.map((payment) => getPaymentSettlementNo(payment)).filter(Boolean))];
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

async function loadApi2State(): Promise<Api2StateFile> {
  try {
    const raw = await readFile(api2StatePath, 'utf8');
    return JSON.parse(raw) as Api2StateFile;
  } catch {
    return {
      meta: {
        lastRunAt: '',
        itemCount: 0,
      },
      items: {},
    };
  }
}

async function saveApi2State(state: Api2StateFile) {
  await ensureDataDir();
  await writeFile(api2StatePath, `${JSON.stringify(state, null, 2)}\n`, 'utf8');
}

function getWorksheetFieldByHeader(worksheet: ExcelJS.Worksheet, headerName: string) {
  const headerRow = worksheet.getRow(1);

  for (let column = 1; column <= worksheet.columnCount; column += 1) {
    const value = String(headerRow.getCell(column).value ?? '').trim();
    if (value === headerName) {
      return column;
    }
  }

  return null;
}

function formatApprovalDate(value?: string | null) {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat('vi-VN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false,
  }).format(date);
}

function formatInvoiceDate(value?: string | null) {
  if (!value) {
    return '';
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return value;
  }

  return new Intl.DateTimeFormat('vi-VN', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour12: false,
  }).format(date);
}

function getSettlementApprovalRowValues(
  payment: SettlementPayment,
  approvalInfo?: SettlementApprovalInfo,
  invoiceDate?: string | null,
) {
  const rowValues = getPaymentRowValues(payment);
  rowValues[10] = formatInvoiceDate(invoiceDate ?? payment.invoiceDate);
  rowValues[12] = formatApprovalDate(approvalInfo?.requesterAprDate);
  rowValues[13] = formatApprovalDate(approvalInfo?.managerAprDate);
  rowValues[14] = formatApprovalDate(approvalInfo?.accountantAprDate);
  return rowValues;
}

function extractInvoiceOrStatementNo(note?: string | null) {
  const text = String(note ?? '').trim();
  const match = text.match(/\b(?:HD|HĐ)\s*:?(?:\s*)([0-9]+(?:-[0-9]+)*)\b/i);
  return match?.[1] ?? '';
}

function extractVendorCode(payeeAccountNo?: string | null) {
  const text = String(payeeAccountNo ?? '').trim();
  const match = text.match(/^(VD\d+)(?:-|$)/i);
  return match?.[1] ?? '';
}

function extractMst(payeeAccountNo?: string | null) {
  const vendorCode = extractVendorCode(payeeAccountNo);
  return vendorCode.replace(/^VD/i, '');
}

function formatAmount(value?: string | number | null) {
  if (value === null || value === undefined || value === '') {
    return '';
  }

  const text = String(value).trim();
  const numeric = Number(text.replace(/,/g, ''));

  if (Number.isNaN(numeric)) {
    return text;
  }

  return new Intl.NumberFormat('en-US').format(numeric);
}

function getServiceCode(note?: string | null) {
  return String(note ?? '')
    .trim()
    .slice(0, 3);
}

function getPaymentRowValues(payment: SettlementPayment) {
  const rowValues = new Array(22).fill('');
  rowValues[0] = extractMst(payment.payeeAccountNo as string | null | undefined);
  rowValues[1] = extractVendorCode(payment.payeeAccountNo as string | null | undefined);
  rowValues[6] = getServiceCode(payment.note as string | null | undefined);
  rowValues[7] = payment.payeeName ?? '';
  rowValues[8] = payment.settlementNo ?? '';
  rowValues[9] = extractInvoiceOrStatementNo(payment.note as string | null | undefined);
  rowValues[11] = formatAmount(payment.amount as string | number | null | undefined);
  return rowValues;
}

async function syncPaymentsToWorkbook(
  payments: SettlementPayment[],
  approvalInfoBySettlementNo: Record<string, SettlementApprovalInfo>,
  invoiceDateBySettlementId: Record<string, string | null | undefined>,
  previousState: Api2StateFile,
): Promise<SyncResult> {
  await ensureDataDir();
  await ensureWorkbookExists();

  if (!payments.length) {
    return { added: 0, updated: 0, unchanged: 0 };
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(outputWorkbookPath);
  const worksheet = workbook.getWorksheet(worksheetName) ?? workbook.getWorksheet(legacyWorksheetName);

  if (!worksheet) {
    throw new Error(`Worksheet not found: ${worksheetName} or ${legacyWorksheetName}`);
  }

  const settlementNoColumn = getWorksheetFieldByHeader(worksheet, 'SỐ ĐNTT-FMS');

  const rowsBySettlementNo = new Map<string, ExcelJS.Row>();

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) {
      return;
    }

    const settlementNo = settlementNoColumn ? String(row.getCell(settlementNoColumn).value ?? '').trim() : '';

    if (settlementNo) {
      rowsBySettlementNo.set(settlementNo, row);
    }
  });

  let added = 0;
  let updated = 0;
  let unchanged = 0;

  for (const payment of payments) {
    const id = getPaymentId(payment);
    const settlementNo = getPaymentSettlementNo(payment);

    if (!id) {
      continue;
    }

    const signature = stringifyPaymentSignature(payment);
    const previousSignature = previousState.items[id];
    const existingRow = settlementNo ? rowsBySettlementNo.get(settlementNo) : undefined;
    const approvalInfo = settlementNo ? approvalInfoBySettlementNo[settlementNo] : undefined;
    const invoiceDate = id ? invoiceDateBySettlementId[id] : undefined;

    if (previousSignature === signature && !approvalInfo) {
      unchanged += 1;
      continue;
    }

    if (previousSignature === signature && approvalInfo && existingRow) {
      const nextValues = getSettlementApprovalRowValues(payment, approvalInfo, invoiceDate);
      nextValues.forEach((value, index) => {
        existingRow.getCell(index + 1).value = value as ExcelJS.CellValue;
      });
      updated += 1;
      continue;
    }

    if (existingRow) {
      const nextValues = getSettlementApprovalRowValues(payment, approvalInfo, invoiceDate);
      nextValues.forEach((value, index) => {
        existingRow.getCell(index + 1).value = value as ExcelJS.CellValue;
      });
      updated += 1;
      continue;
    }

    if (previousSignature) {
      unchanged += 1;
      continue;
    }

    const nextValues = getSettlementApprovalRowValues(payment, approvalInfo, invoiceDate);
    worksheet.addRow(nextValues);
    const newRow = worksheet.lastRow!;
    if (settlementNo) {
      rowsBySettlementNo.set(settlementNo, newRow);
    }
    added += 1;
  }

  if (added || updated) {
    await workbook.xlsx.writeFile(outputWorkbookPath);
  }

  return { added, updated, unchanged };
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
  const settlementNos = getUniqueSettlementNos(payments);
  const api3Responses: Record<string, unknown> = {};

  for (const settlementNo of settlementNos) {
    const response = await client.getJson<unknown>(
      '/Accounting/api/v1/en-US/AcctSettlementPayment/GetInfoApproveSettlementBySettlementNo',
      {
        auth: false,
        headers: {
          Authorization: `Bearer ${token}`,
        },
      },
      {
        settlementNo,
      },
    );

    api3Responses[settlementNo] = response;
  }

  await saveApi3Response(api3Responses);
  logStep('api3', `đã lưu API 3 cho ${settlementNos.length} settlementNo`);

  const settlementIds = [...new Set(payments.map((payment) => getPaymentId(payment)).filter(Boolean))];
  const api4Responses: Record<string, unknown> = {};
  const invoiceDateBySettlementId: Record<string, string | null | undefined> = {};

  for (const settlementId of settlementIds) {
    const response = await client.getJson<unknown>(
      '/Accounting/api/v1/en-US/AcctSettlementPayment/GetDetailSettlementPaymentById',
      {
        auth: false,
        headers: {
          Authorization: `Bearer ${token}`,
        },
      },
      {
        settlementId,
        view: 'LIST',
      },
    );

    const chargeNoGrpSettlement = extractChargeNoGrpSettlement(response) as { invoiceDate?: string | null } | null;
    api4Responses[settlementId] = chargeNoGrpSettlement;
    invoiceDateBySettlementId[settlementId] = chargeNoGrpSettlement?.invoiceDate;
  }

  await saveApi4Response(api4Responses, settlementIds);
  logStep('api4', `đã lưu API 4 cho ${settlementIds.length} settlementId`);

  const approvalInfoBySettlementNo = Object.fromEntries(
    Object.entries(api3Responses).map(([settlementNo, response]) => [
      settlementNo,
      (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.data ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.result ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.approvalInfo ??
        (response as SettlementApprovalInfo),
    ]),
  );

  const previousState = await loadApi2State();

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

    const result = await syncPaymentsToWorkbook(
      payments,
      approvalInfoBySettlementNo,
      invoiceDateBySettlementId,
      previousState,
    );
    logStep('workbook', `thêm ${result.added} dòng, cập nhật ${result.updated} dòng, giữ nguyên ${result.unchanged} dòng`);

    const stateItems = Object.fromEntries(
      payments
        .map((payment) => {
          const id = getPaymentId(payment);
          return id ? [id, stringifyPaymentSignature(payment)] : null;
        })
        .filter((entry): entry is [string, string] => entry !== null),
    );

    await saveApi2State({
      meta: {
        lastRunAt: new Date().toISOString(),
        itemCount: Object.keys(stateItems).length,
      },
      items: stateItems,
    });

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
