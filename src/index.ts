import { mkdir, readFile, writeFile, copyFile } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import ExcelJS from 'exceljs';
import { EfmsClient } from './client.js';
import { getAccessToken, getApi5TokenResponse, getApi9TokenResponse } from './auth.js';
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
const api5OutputPath = join(dataDir, 'api5-response.json');
const api6OutputPath = join(dataDir, 'api6-response.json');
const api7OutputPath = join(dataDir, 'api7-response.json');
const api8OutputPath = join(dataDir, 'api8-response.json');
const api9OutputPath = join(dataDir, 'api9-response.json');
const api10OutputPath = join(dataDir, 'api10-response.json');
const api11OutputPath = join(dataDir, 'api11-response.json');
const api12OutputPath = join(dataDir, 'api12-response.json');
const api2StatePath = join(dataDir, 'api2-state.json');
const api6StatePath = join(dataDir, 'api6-state.json');
const api10StatePath = join(dataDir, 'api10-state.json');
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

type Api5ResponseFile = {
  meta: {
    lastRunAt: string;
    status: number;
    headers: Record<string, string>;
  };
  response: unknown;
};

type Api6ResponseFile = {
  meta: {
    lastRunAt: string;
    status: number;
    headers: Record<string, string>;
    sourceToken: 'api5';
  };
  response: unknown;
};

type Api7ResponseFile = {
  meta: {
    lastRunAt: string;
    settlementCount: number;
    source: 'api6';
  };
  response: Record<string, unknown>;
};

type Api8ResponseFile = {
  meta: {
    lastRunAt: string;
    settlementCount: number;
    source: 'api6';
  };
  response: Record<string, unknown>;
};

type Api9ResponseFile = {
  meta: {
    lastRunAt: string;
    status: number;
    headers: Record<string, string>;
  };
  response: unknown;
};

type Api10ResponseFile = {
  meta: {
    lastRunAt: string;
    status: number;
    headers: Record<string, string>;
    sourceToken: 'api9';
  };
  response: unknown;
};

type Api11ResponseFile = {
  meta: {
    lastRunAt: string;
    settlementCount: number;
    source: 'api10';
  };
  response: Record<string, unknown>;
};

type Api12ResponseFile = {
  meta: {
    lastRunAt: string;
    settlementCount: number;
    source: 'api10';
  };
  response: Record<string, unknown>;
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

type ApiStateFile = {
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

async function saveApi5Response(response: Response, text: string) {
  await ensureDataDir();

  let parsed: unknown = text;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = text;
  }

  const headers = Object.fromEntries(response.headers.entries());
  const payload: Api5ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      status: response.status,
      headers,
    },
    response: parsed,
  };

  await writeFile(api5OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi6Response(response: Response, text: string) {
  await ensureDataDir();

  let parsed: unknown = text;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = text;
  }

  const headers = Object.fromEntries(response.headers.entries());
  const payload: Api6ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      status: response.status,
      headers,
      sourceToken: 'api5',
    },
    response: parsed,
  };

  await writeFile(api6OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
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

async function saveApi7Response(data: Record<string, unknown>, settlementNos: string[]) {
  await ensureDataDir();

  const payload: Api7ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      settlementCount: settlementNos.length,
      source: 'api6',
    },
    response: data,
  };

  await writeFile(api7OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi8Response(data: Record<string, unknown>, settlementIds: string[]) {
  await ensureDataDir();

  const payload: Api8ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      settlementCount: settlementIds.length,
      source: 'api6',
    },
    response: data,
  };

  await writeFile(api8OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi9Response(response: Response, text: string) {
  await ensureDataDir();

  let parsed: unknown = text;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = text;
  }

  const headers = Object.fromEntries(response.headers.entries());
  const payload: Api9ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      status: response.status,
      headers,
    },
    response: parsed,
  };

  await writeFile(api9OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi10Response(response: Response, text: string) {
  await ensureDataDir();

  let parsed: unknown = text;
  try {
    parsed = JSON.parse(text);
  } catch {
    parsed = text;
  }

  const headers = Object.fromEntries(response.headers.entries());
  const payload: Api10ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      status: response.status,
      headers,
      sourceToken: 'api9',
    },
    response: parsed,
  };

  await writeFile(api10OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi11Response(data: Record<string, unknown>, settlementNos: string[]) {
  await ensureDataDir();

  const payload: Api11ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      settlementCount: settlementNos.length,
      source: 'api10',
    },
    response: data,
  };

  await writeFile(api11OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

async function saveApi12Response(data: Record<string, unknown>, settlementIds: string[]) {
  await ensureDataDir();

  const payload: Api12ResponseFile = {
    meta: {
      lastRunAt: new Date().toISOString(),
      settlementCount: settlementIds.length,
      source: 'api10',
    },
    response: data,
  };

  await writeFile(api12OutputPath, `${JSON.stringify(payload, null, 2)}\n`, 'utf8');
}

function normalizeApi10Payments(response: unknown) {
  return normalizeApiPayments(response).map((payment) => ({
    ...payment,
    amount: payment.amount,
    payeeName: payment.payeeName,
    payeeAccountNo: payment.payeeAccountNo,
    settlementNo: payment.settlementNo,
    note: payment.note,
    invoiceDate: payment.invoiceDate ?? null,
    invoiceNo: payment.invoiceNo ?? null,
  }));
}

function enrichPaymentsWithDetails(
  payments: SettlementPayment[],
  paymentDetailsById: Record<string, unknown>,
) {
  return payments.map((payment) => {
    const detail = payment.id ? (paymentDetailsById[payment.id] as Record<string, unknown> | undefined) : undefined;
    return {
      ...payment,
      invoiceDate: payment.invoiceDate ?? (detail?.invoiceDate as string | null | undefined) ?? null,
      invoiceNo: payment.invoiceNo ?? (detail?.invoiceNo as string | null | undefined) ?? null,
      note: payment.note ?? (detail?.notes as string | null | undefined) ?? null,
    };
  });
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

function getRequesterAprDateValue(value?: string | null) {
  if (!value) {
    return Number.POSITIVE_INFINITY;
  }

  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return Number.POSITIVE_INFINITY;
  }

  return date.getTime();
}

function sortPaymentsByRequesterAprDate(
  payments: SettlementPayment[],
  approvalInfoBySettlementNo: Record<string, SettlementApprovalInfo>,
) {
  return [...payments].sort((a, b) => {
    const aSettlementNo = getPaymentSettlementNo(a);
    const bSettlementNo = getPaymentSettlementNo(b);
    const aRequesterAprDate = getRequesterAprDateValue(approvalInfoBySettlementNo[aSettlementNo]?.requesterAprDate);
    const bRequesterAprDate = getRequesterAprDateValue(approvalInfoBySettlementNo[bSettlementNo]?.requesterAprDate);

    if (aRequesterAprDate !== bRequesterAprDate) {
      return aRequesterAprDate - bRequesterAprDate;
    }

    return aSettlementNo.localeCompare(bSettlementNo);
  });
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

async function loadApiState(path: string): Promise<ApiStateFile> {
  try {
    const raw = await readFile(path, 'utf8');
    return JSON.parse(raw) as ApiStateFile;
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

async function saveApiState(path: string, state: ApiStateFile) {
  await ensureDataDir();
  await writeFile(path, `${JSON.stringify(state, null, 2)}\n`, 'utf8');
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
  groupEfms = '',
) {
  const rowValues = getPaymentRowValues(payment, groupEfms);
  rowValues[10] = formatInvoiceDate(invoiceDate ?? payment.invoiceDate);
  rowValues[12] = formatApprovalDate(approvalInfo?.requesterAprDate);
  rowValues[13] = formatApprovalDate(approvalInfo?.managerAprDate);
  rowValues[14] = formatApprovalDate(approvalInfo?.accountantAprDate);
  return rowValues;
}

function setPaymentRowValuesByColumnHeaders(
  row: ExcelJS.Row,
  payment: SettlementPayment,
  columnHeaders: Map<string, number>,
  groupEfms: string,
  invoiceDate?: string | null,
  approvalInfo?: SettlementApprovalInfo,
) {
  const mapping = buildColumnMapping(payment, groupEfms, invoiceDate, approvalInfo);
  
  for (const [headerName, value] of Object.entries(mapping)) {
    const colNum = columnHeaders.get(headerName);
    if (colNum !== undefined && value !== undefined) {
      row.getCell(colNum).value = value as ExcelJS.CellValue;
    }
  }
}

function extractInvoiceOrStatementNo(note?: string | null, invoiceNo?: string | null) {
  const invoiceText = String(invoiceNo ?? '').trim();
  if (invoiceText) {
    return invoiceText;
  }

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

type ColumnMapping = Record<string, string | number | undefined>;

function buildColumnMapping(payment: SettlementPayment, groupEfms: string, invoiceDate?: string | null, approvalInfo?: SettlementApprovalInfo): ColumnMapping {
  return {
    'MST': extractMst(payment.payeeAccountNo as string | null | undefined),
    'VENDOR CODE': extractVendorCode(payment.payeeAccountNo as string | null | undefined),
    'DỊCH VỤ': getServiceCode(payment.note as string | null | undefined),
    'NCC': payment.payeeName ?? '',
    'SỐ ĐNTT-FMS': payment.settlementNo ?? '',
    'SỐ HOÁ ĐƠN/BANG KÊ': extractInvoiceOrStatementNo(payment.note as string | null | undefined, payment.invoiceNo as string | null | undefined),
    'NGÀY HOÁ ĐƠN': formatInvoiceDate(invoiceDate ?? payment.invoiceDate),
    'SỐ TIỀN99': formatAmount(payment.amount as string | number | null | undefined),
    'NGÀY LẬP - FMS': formatApprovalDate(approvalInfo?.requesterAprDate),
    'NGÀY DUYỆT - HOD': formatApprovalDate(approvalInfo?.managerAprDate),
    'NGÀY KẾ TOÁN ĐÃ KIỂM TRA': formatApprovalDate(approvalInfo?.accountantAprDate),
    'GROUP EFMS': groupEfms,
    'GHI CHÚ (NẾU BỊ THIẾU CHỨNG)': payment.note ?? '',
  };
}

function getPaymentRowValues(payment: SettlementPayment, groupEfms: string) {
  const rowValues = new Array(22).fill('');
  rowValues[0] = extractMst(payment.payeeAccountNo as string | null | undefined);
  rowValues[1] = extractVendorCode(payment.payeeAccountNo as string | null | undefined);
  rowValues[6] = getServiceCode(payment.note as string | null | undefined);
  rowValues[7] = payment.payeeName ?? '';
  rowValues[8] = payment.settlementNo ?? '';
  rowValues[9] = extractInvoiceOrStatementNo(payment.note as string | null | undefined, payment.invoiceNo as string | null | undefined);
  rowValues[11] = formatAmount(payment.amount as string | number | null | undefined);
  rowValues[15] = groupEfms;
  return rowValues;
}

function buildColumnHeadersMap(worksheet: ExcelJS.Worksheet): Map<string, number> {
  const headerRow = worksheet.getRow(1);
  const headers = new Map<string, number>();
  
  for (let column = 1; column <= worksheet.columnCount; column += 1) {
    const value = String(headerRow.getCell(column).value ?? '').trim();
    if (value) {
      headers.set(value, column);
    }
  }
  
  return headers;
}

async function syncPaymentsToWorkbook(
  payments: SettlementPayment[],
  approvalInfoBySettlementNo: Record<string, SettlementApprovalInfo>,
  invoiceDateBySettlementId: Record<string, string | null | undefined>,
  previousState: ApiStateFile,
  groupEfms: string,
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

  const columnHeaders = buildColumnHeadersMap(worksheet);
  const settlementNoColumn = columnHeaders.get('SỐ ĐNTT-FMS');

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

    if (existingRow) {
      if (previousSignature === signature && !approvalInfo) {
        unchanged += 1;
        continue;
      }

      setPaymentRowValuesByColumnHeaders(existingRow, payment, columnHeaders, groupEfms, invoiceDate, approvalInfo);
      updated += 1;
      continue;
    }

    const newRow = worksheet.addRow([]);
    setPaymentRowValuesByColumnHeaders(newRow, payment, columnHeaders, groupEfms, invoiceDate, approvalInfo);
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

function normalizeApiPayments(response: unknown) {
  if (!response || typeof response !== 'object') {
    return [] as SettlementPayment[];
  }

  const record = response as { data?: unknown; result?: unknown; items?: unknown };
  const payments = record.data ?? record.result ?? record.items ?? [];
  return Array.isArray(payments) ? (payments as SettlementPayment[]) : [];
}

function shouldUseApi6StateCheck(payment: SettlementPayment) {
  return Boolean(getPaymentId(payment));
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

  logStep('api5', 'bắt đầu gọi API 5');
  const api5 = await getApi5TokenResponse();
  await saveApi5Response(api5.response, api5.text);
  logStep('api5', `đã call và lưu response API 5 (HTTP ${api5.response.status})`);

  const api5TokenResponse = JSON.parse(api5.text) as { access_token?: string };
  const api5AccessToken = api5TokenResponse.access_token;

  if (!api5AccessToken) {
    throw new Error('API 5 response did not include access_token for API 6');
  }

  const previousApi6State = await loadApiState(api6StatePath);
  logStep('api6', 'bắt đầu gọi API 6');
  const api6RawResponse = await client.request(
    '/Accounting/api/v1/en-US/AcctSettlementPayment/paging?pageNumber=1&pageSize=1000',
    {
      method: 'POST',
      auth: false,
      headers: {
        Authorization: `Bearer ${api5AccessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        requester: '606b7384-2589-43f3-a4d7-95a94ac026f4',
      }),
    },
  );
  const api6Text = await api6RawResponse.text();
  await saveApi6Response(api6RawResponse, api6Text);
  logStep('api6', `đã call và lưu response API 6 (HTTP ${api6RawResponse.status})`);

  const api6Parsed = JSON.parse(api6Text) as unknown;
  const api6Payments = normalizeApiPayments(api6Parsed);
  const api6SettlementNos = getUniqueSettlementNos(api6Payments);
  const api6PaymentsWithState = api6Payments.filter(shouldUseApi6StateCheck);
  const api6StateItems = Object.fromEntries(
    api6PaymentsWithState
      .map((payment) => {
        const id = getPaymentId(payment);
        return id ? [id, stringifyPaymentSignature(payment)] : null;
      })
      .filter((entry): entry is [string, string] => entry !== null),
  );

  const api7Responses: Record<string, unknown> = {};

  for (const settlementNo of api6SettlementNos) {
    const response = await client.getJson<unknown>(
      '/Accounting/api/v1/en-US/AcctSettlementPayment/GetInfoApproveSettlementBySettlementNo',
      {
        auth: false,
        headers: {
          Authorization: `Bearer ${api5AccessToken}`,
        },
      },
      {
        settlementNo,
      },
    );

    api7Responses[settlementNo] = response;
  }

  await saveApi7Response(api7Responses, api6SettlementNos);
  logStep('api7', `đã lưu API 7 cho ${api6SettlementNos.length} settlementNo từ API 6`);

  const api6SettlementIds = [...new Set(api6Payments.map((payment) => getPaymentId(payment)).filter(Boolean))];
  const api8Responses: Record<string, unknown> = {};

  for (const settlementId of api6SettlementIds) {
    const response = await client.getJson<unknown>(
      '/Accounting/api/v1/en-US/AcctSettlementPayment/GetDetailSettlementPaymentById',
      {
        auth: false,
        headers: {
          Authorization: `Bearer ${api5AccessToken}`,
        },
      },
      {
        settlementId,
        view: 'LIST',
      },
    );

    const chargeNoGrpSettlement = extractChargeNoGrpSettlement(response) as { invoiceDate?: string | null } | null;
    api8Responses[settlementId] = chargeNoGrpSettlement;
    invoiceDateBySettlementId[settlementId] = chargeNoGrpSettlement?.invoiceDate;
  }

  await saveApi8Response(api8Responses, api6SettlementIds);
  logStep('api8', `đã lưu API 8 cho ${api6SettlementIds.length} settlementId từ API 6`);

  const api6ChangedCount = Object.entries(api6StateItems).filter(([id, signature]) => previousApi6State.items[id] !== signature).length;

  await saveApiState(api6StatePath, {
    meta: {
      lastRunAt: new Date().toISOString(),
      itemCount: Object.keys(api6StateItems).length,
    },
    items: api6StateItems,
  });
  logStep('api6-state', `đã lưu state cho ${Object.keys(api6StateItems).length} bản ghi, thay đổi ${api6ChangedCount} bản ghi`);

  const api6PaymentsForWorkbook = api6Payments.length ? api6Payments : payments;

  const approvalInfoBySettlementNo = Object.fromEntries(
    Object.entries(api7Responses).map(([settlementNo, response]) => [
      settlementNo,
      (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.data ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.result ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.approvalInfo ??
        (response as SettlementApprovalInfo),
    ]),
  );

  const previousState = await loadApiState(api2StatePath);

  if (api6PaymentsForWorkbook.length) {
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

    const sortedPayments = sortPaymentsByRequesterAprDate(api6PaymentsForWorkbook, approvalInfoBySettlementNo);

    const groupEfms = api6Payments.length > 0 ? 'STL_TKI' : 'OPS_MANAGEMENT';
    const result = await syncPaymentsToWorkbook(
      sortedPayments,
      approvalInfoBySettlementNo,
      invoiceDateBySettlementId,
      previousState,
      groupEfms,
    );
    logStep('workbook', `thêm ${result.added} dòng, cập nhật ${result.updated} dòng, giữ nguyên ${result.unchanged} dòng`);

    const stateItems = Object.fromEntries(
      api6PaymentsForWorkbook
        .map((payment) => {
          const id = getPaymentId(payment);
          return id ? [id, stringifyPaymentSignature(payment)] : null;
        })
        .filter((entry): entry is [string, string] => entry !== null),
    );

    const api2ChangedCount = Object.entries(stateItems).filter(([id, signature]) => previousState.items[id] !== signature).length;

    await saveApiState(api2StatePath, {
      meta: {
        lastRunAt: new Date().toISOString(),
        itemCount: Object.keys(stateItems).length,
      },
      items: stateItems,
    });

    await saveApiState(api6StatePath, {
      meta: {
        lastRunAt: new Date().toISOString(),
        itemCount: Object.keys(stateItems).length,
      },
      items: stateItems,
    });
    logStep('state', `API 2 thay đổi ${api2ChangedCount} bản ghi, API 6 thay đổi ${api6ChangedCount} bản ghi`);

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

  logStep('api9', 'bắt đầu gọi API 9');
  const api9 = await getApi9TokenResponse();
  await saveApi9Response(api9.response, api9.text);
  logStep('api9', `đã call và lưu response API 9 (HTTP ${api9.response.status})`);

  const api9TokenResponse = JSON.parse(api9.text) as { access_token?: string };
  const api9AccessToken = api9TokenResponse.access_token;

  if (!api9AccessToken) {
    throw new Error('API 9 response did not include access_token for API 10');
  }

  logStep('api10', 'bắt đầu gọi API 10');
  const api10RawResponse = await client.request(
    '/Accounting/api/v1/en-US/AcctSettlementPayment/paging?pageNumber=1&pageSize=1000',
    {
      method: 'POST',
      auth: false,
      headers: {
        Authorization: `Bearer ${api9AccessToken}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        requester: '606b7384-2589-43f3-a4d7-95a94ac026f4',
      }),
    },
  );
  const api10Text = await api10RawResponse.text();
  await saveApi10Response(api10RawResponse, api10Text);
  logStep('api10', `đã call và lưu response API 10 (HTTP ${api10RawResponse.status})`);

  const previousApi10State = await loadApiState(api10StatePath);
  const api10Parsed = JSON.parse(api10Text) as unknown;
  const api10Payments = normalizeApi10Payments(api10Parsed);
  const api10StateItems = Object.fromEntries(
    api10Payments
      .map((payment) => {
        const id = getPaymentId(payment);
        return id ? [id, stringifyPaymentSignature(payment)] : null;
      })
      .filter((entry): entry is [string, string] => entry !== null),
  );

  const api10ChangedCount = Object.entries(api10StateItems).filter(
    ([id, signature]) => previousApi10State.items[id] !== signature,
  ).length;

  await saveApiState(api10StatePath, {
    meta: {
      lastRunAt: new Date().toISOString(),
      itemCount: Object.keys(api10StateItems).length,
    },
    items: api10StateItems,
  });

  const api10SettlementIds = [...new Set(api10Payments.map((payment) => getPaymentId(payment)).filter(Boolean))];
  const api10SettlementNos = getUniqueSettlementNos(api10Payments);

  const api12Responses: Record<string, unknown> = {};
  if (api10SettlementIds.length) {
    logStep('api12', 'bắt đầu gọi API 12');

    for (const settlementId of api10SettlementIds) {
      try {
        const response = await client.getJson<unknown>(
          '/Accounting/api/v1/en-US/AcctSettlementPayment/GetDetailSettlementPaymentById',
          {
            auth: false,
            headers: {
              Authorization: `Bearer ${api9AccessToken}`,
            },
          },
          {
            settlementId,
            view: 'LIST',
          },
        );

        const chargeNoGrpSettlement = extractChargeNoGrpSettlement(response);
        api12Responses[settlementId] = chargeNoGrpSettlement;
        invoiceDateBySettlementId[settlementId] = chargeNoGrpSettlement?.invoiceDate;
      } catch (error) {
        logStep('api12', `lỗi khi gọi API 12 cho settlementId ${settlementId}: ${error}`);
        api12Responses[settlementId] = { error: String(error) };
      }
    }

    await saveApi12Response(api12Responses, api10SettlementIds);
    logStep('api12', `đã lưu API 12 cho ${api10SettlementIds.length} settlementId từ API 10`);
  }

  const api11Responses: Record<string, unknown> = {};
  if (api10SettlementNos.length) {
    logStep('api11', 'bắt đầu gọi API 11');

    for (const settlementNo of api10SettlementNos) {
      try {
        const response = await client.getJson<unknown>(
          '/Accounting/api/v1/en-US/AcctSettlementPayment/GetInfoApproveSettlementBySettlementNo',
          {
            auth: false,
            headers: {
              Authorization: `Bearer ${api9AccessToken}`,
            },
          },
          {
            settlementNo,
          },
        );

        api11Responses[settlementNo] = response;
      } catch (error) {
        logStep('api11', `lỗi khi gọi API 11 cho settlementNo ${settlementNo}: ${error}`);
        api11Responses[settlementNo] = { error: String(error) };
      }
    }

    await saveApi11Response(api11Responses, api10SettlementNos);
    logStep('api11', `đã lưu API 11 cho ${api10SettlementNos.length} settlementNo từ API 10`);
  }

  const api10ApprovalInfoBySettlementNo = Object.fromEntries(
    Object.entries(api11Responses).map(([settlementNo, response]) => [
      settlementNo,
      (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.data ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.result ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.approvalInfo ??
        (response as SettlementApprovalInfo),
    ]),
  );

  const mergedApprovalInfoBySettlementNo = {
    ...approvalInfoBySettlementNo,
    ...api10ApprovalInfoBySettlementNo,
  };

  const api10PaymentsForWorkbook = api10Payments.length ? enrichPaymentsWithDetails(api10Payments, api12Responses) : payments;
  if (api10PaymentsForWorkbook.length) {
    const previousApi10WorkbookState = await loadApiState(api2StatePath);
    const sortedApi10Payments = sortPaymentsByRequesterAprDate(api10PaymentsForWorkbook, mergedApprovalInfoBySettlementNo);
    const resultApi10 = await syncPaymentsToWorkbook(
      sortedApi10Payments,
      mergedApprovalInfoBySettlementNo,
      invoiceDateBySettlementId,
      previousApi10WorkbookState,
      'STL_TKI',
    );

    const api10WorkbookStateItems = Object.fromEntries(
      api10PaymentsForWorkbook
        .map((payment) => {
          const id = getPaymentId(payment);
          return id ? [id, stringifyPaymentSignature(payment)] : null;
        })
        .filter((entry): entry is [string, string] => entry !== null),
    );

    await saveApiState(api2StatePath, {
      meta: {
        lastRunAt: new Date().toISOString(),
        itemCount: Object.keys(api10WorkbookStateItems).length,
      },
      items: api10WorkbookStateItems,
    });

    logStep('api10-workbook', `thêm ${resultApi10.added} dòng, cập nhật ${resultApi10.updated} dòng, giữ nguyên ${resultApi10.unchanged} dòng`);
    logStep('api10-state', `đã lưu state cho ${Object.keys(api10StateItems).length} bản ghi, thay đổi ${api10ChangedCount} bản ghi`);
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
