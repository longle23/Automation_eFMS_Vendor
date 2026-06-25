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
  departmentName?: string | null;
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

function sortPaymentsBySettlementNoAndRequesterAprDate(
  payments: SettlementPayment[],
  approvalInfoBySettlementNo: Record<string, SettlementApprovalInfo>,
) {
  return [...payments].sort((a, b) => {
    const aSettlementNo = getPaymentSettlementNo(a);
    const bSettlementNo = getPaymentSettlementNo(b);

    const settlementNoCompare = aSettlementNo.localeCompare(bSettlementNo);
    if (settlementNoCompare !== 0) {
      return settlementNoCompare;
    }

    const aRequesterAprDate = getRequesterAprDateValue(approvalInfoBySettlementNo[aSettlementNo]?.requesterAprDate);
    const bRequesterAprDate = getRequesterAprDateValue(approvalInfoBySettlementNo[bSettlementNo]?.requesterAprDate);

    if (aRequesterAprDate !== bRequesterAprDate) {
      return aRequesterAprDate - bRequesterAprDate;
    }

    return stringifyPaymentSignature(a).localeCompare(stringifyPaymentSignature(b));
  });
}

function stringifyPaymentSignature(payment: SettlementPayment) {
  const normalized = {
    id: getPaymentId(payment),
    settlementNo: getPaymentSettlementNo(payment),
    payeeName: payment.payeeName ?? '',
    payeeAccountNo: payment.payeeAccountNo ?? '',
    invoiceNo: payment.invoiceNo ?? '',
    invoiceDate: payment.invoiceDate ?? '',
    amount: payment.amount ?? '',
    note: payment.note ?? '',
  };

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
) {
  const rowValues = getPaymentRowValues(payment);
  rowValues[10] = formatInvoiceDate(invoiceDate ?? payment.invoiceDate);
  rowValues[12] = formatApprovalDate(approvalInfo?.requesterAprDate);
  rowValues[13] = formatApprovalDate(approvalInfo?.managerAprDate);
  rowValues[14] = formatApprovalDate(approvalInfo?.accountantAprDate);
  return rowValues;
}

type WorkbookColumnSpec = {
  key: string;
  index: number;
  value: (payment: SettlementPayment, invoiceDate?: string | null, approvalInfo?: SettlementApprovalInfo) => ExcelJS.CellValue;
};

const WORKBOOK_COLUMN_SPECS: WorkbookColumnSpec[] = [
  { key: 'mst', index: 1, value: (payment) => extractMst(payment.payeeAccountNo as string | null | undefined) },
  { key: 'vendorCode', index: 2, value: (payment) => extractVendorCode(payment.payeeAccountNo as string | null | undefined) },
  { key: 'service', index: 7, value: (payment) => getServiceCode(payment.note as string | null | undefined) },
  { key: 'vendorName', index: 8, value: (payment) => payment.payeeName ?? '' },
  { key: 'settlementNo', index: 9, value: (payment) => payment.settlementNo ?? '' },
  {
    key: 'invoiceOrStatementNo',
    index: 10,
    value: (payment) => extractInvoiceOrStatementNo(payment.note as string | null | undefined, payment.invoiceNo as string | null | undefined),
  },
  { key: 'invoiceDate', index: 11, value: (payment, invoiceDate) => formatInvoiceDate(invoiceDate ?? payment.invoiceDate) },
  { key: 'amount', index: 12, value: (payment) => formatAmount(payment.amount as string | number | null | undefined) },
  { key: 'departmentName', index: 16, value: (payment) => getDepartmentName(payment) },
  { key: 'requesterApprovedAt', index: 13, value: (_payment, _invoiceDate, approvalInfo) => formatApprovalDate(approvalInfo?.requesterAprDate) },
  { key: 'managerApprovedAt', index: 14, value: (_payment, _invoiceDate, approvalInfo) => formatApprovalDate(approvalInfo?.managerAprDate) },
  { key: 'accountantApprovedAt', index: 15, value: (_payment, _invoiceDate, approvalInfo) => formatApprovalDate(approvalInfo?.accountantAprDate) },
];

function setPaymentRowValuesFromTemplate(
  row: ExcelJS.Row,
  payment: SettlementPayment,
  invoiceDate?: string | null,
  approvalInfo?: SettlementApprovalInfo,
) {
  for (const spec of WORKBOOK_COLUMN_SPECS) {
    const value = spec.value(payment, invoiceDate, approvalInfo);
    if (value !== undefined) {
      row.getCell(spec.index).value = value;
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

function getDepartmentName(payment: SettlementPayment) {
  return String(payment.departmentName ?? '').trim();
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
  const rowValues = new Array(23).fill('');
  rowValues[0] = extractMst(payment.payeeAccountNo as string | null | undefined);
  rowValues[1] = extractVendorCode(payment.payeeAccountNo as string | null | undefined);
  rowValues[6] = getServiceCode(payment.note as string | null | undefined);
  rowValues[7] = payment.payeeName ?? '';
  rowValues[8] = payment.settlementNo ?? '';
  rowValues[9] = extractInvoiceOrStatementNo(payment.note as string | null | undefined, payment.invoiceNo as string | null | undefined);
  rowValues[11] = formatAmount(payment.amount as string | number | null | undefined);
  rowValues[15] = getDepartmentName(payment);
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

function findColumnByKeywords(headers: Map<string, number>, keywords: string[]) {
  for (const [headerName, column] of headers.entries()) {
    const normalized = headerName.toLowerCase();
    if (keywords.every((keyword) => normalized.includes(keyword.toLowerCase()))) {
      return column;
    }
  }

  return null;
}

function getWorkbookExistingIds(worksheet: ExcelJS.Worksheet) {
  const existingIds = new Set<string>();

  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const id = String(worksheet.getRow(rowNumber).getCell(1).value ?? '').trim();
    if (id) {
      existingIds.add(id);
    }
  }

  return existingIds;
}

async function syncPaymentsToWorkbook(
  payments: SettlementPayment[],
  approvalInfoBySettlementNo: Record<string, SettlementApprovalInfo>,
  invoiceDateBySettlementId: Record<string, string | null | undefined>,
  previousState: ApiStateFile,
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
  const settlementNoColumn = findColumnByKeywords(columnHeaders, ['đntt', 'fms']) ?? findColumnByKeywords(columnHeaders, ['settlement', 'no']);

  const existingIds = getWorkbookExistingIds(worksheet);
  const rowsToWrite: Array<{ payment: SettlementPayment; invoiceDate?: string | null; approvalInfo?: SettlementApprovalInfo }> = [];
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
    const approvalInfo = settlementNo ? approvalInfoBySettlementNo[settlementNo] : undefined;
    const invoiceDate = invoiceDateBySettlementId[id];

    if (existingIds.has(id)) {
      unchanged += 1;
      continue;
    }

    if (previousSignature === signature) {
      unchanged += 1;
      continue;
    }

    rowsToWrite.push({ payment, invoiceDate, approvalInfo });
    added += 1;
  }

  if (!rowsToWrite.length) {
    return { added: 0, updated, unchanged };
  }

  const sortedRowsToWrite = settlementNoColumn
    ? [...rowsToWrite].sort((a, b) => {
        const aSettlementNo = getPaymentSettlementNo(a.payment);
        const bSettlementNo = getPaymentSettlementNo(b.payment);
        return aSettlementNo.localeCompare(bSettlementNo);
      })
    : rowsToWrite;

  const lastDataRow = Math.max(worksheet.lastRow?.number ?? 1, worksheet.rowCount);
  if (lastDataRow > 1) {
    worksheet.spliceRows(2, lastDataRow - 1);
  }

  let nextRowNumber = 2;
  for (const { payment, invoiceDate, approvalInfo } of sortedRowsToWrite) {
    const row = worksheet.getRow(nextRowNumber);
    setPaymentRowValuesFromTemplate(row, payment, invoiceDate, approvalInfo);
    row.commit();
    nextRowNumber += 1;
  }

  await workbook.xlsx.writeFile(outputWorkbookPath);

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

function removeNoteField(payments: SettlementPayment[]) {
  return payments.map((payment) => {
    const { note, ...rest } = payment as Record<string, unknown>;
    return rest as SettlementPayment;
  });
}

function mergePaymentsById(...paymentGroups: SettlementPayment[][]) {
  const merged = new Map<string, SettlementPayment>();

  for (const group of paymentGroups) {
    for (const payment of group) {
      const id = getPaymentId(payment);
      if (!id) {
        continue;
      }
      merged.set(id, payment);
    }
  }

  return [...merged.values()];
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

  const api2Payments = removeNoteField(normalizeApiPayments(data));
  const api2StateItems = Object.fromEntries(
    api2Payments
      .map((payment) => {
        const id = getPaymentId(payment);
        return id ? [id, stringifyPaymentSignature(payment)] : null;
      })
      .filter((entry): entry is [string, string] => entry !== null),
  );

  await saveApiState(api2StatePath, {
    meta: {
      lastRunAt: new Date().toISOString(),
      itemCount: Object.keys(api2StateItems).length,
    },
    items: api2StateItems,
  });
  logStep('api2-state', `đã lưu state cho ${Object.keys(api2StateItems).length} bản ghi`);

  const payments = removeNoteField(iteratePaymentsFromLastToFirst(api2Payments));
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
  const api6Payments = removeNoteField(normalizeApiPayments(api6Parsed));
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

  const approvalInfoBySettlementNo = Object.fromEntries(
    Object.entries(api7Responses).map(([settlementNo, response]) => [
      settlementNo,
      (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.data ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.result ??
        (response as { data?: SettlementApprovalInfo; result?: SettlementApprovalInfo; approvalInfo?: SettlementApprovalInfo })?.approvalInfo ??
        (response as SettlementApprovalInfo),
    ]),
  );

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
  const api10Payments = removeNoteField(normalizeApi10Payments(api10Parsed));
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
  logStep('api10-state', `đã lưu state cho ${Object.keys(api10StateItems).length} bản ghi, thay đổi ${api10ChangedCount} bản ghi`);

  const api10SettlementIds = [...new Set(api10Payments.map((payment) => getPaymentId(payment)).filter(Boolean))];
  const api10SettlementNos = getUniqueSettlementNos(api10Payments);

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

  const api10PaymentsForWorkbook = api10Payments.length ? removeNoteField(enrichPaymentsWithDetails(api10Payments, api12Responses)) : [];
  const api2PaymentsForWorkbook = api2Payments.length ? api2Payments : [];
  const api6PaymentsForWorkbook = api6Payments.length ? api6Payments : [];
  const combinedPaymentsForWorkbook = mergePaymentsById(api2PaymentsForWorkbook, api6PaymentsForWorkbook, api10PaymentsForWorkbook);
  const combinedApprovalInfoBySettlementNo = {
    ...approvalInfoBySettlementNo,
    ...mergedApprovalInfoBySettlementNo,
  };

  if (combinedPaymentsForWorkbook.length) {
    const previousApi10WorkbookState = await loadApiState(api10StatePath);

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
    } else {
      await ensureWorkbookExists();
    }

    const sortedCombinedPayments = sortPaymentsBySettlementNoAndRequesterAprDate(
      combinedPaymentsForWorkbook,
      combinedApprovalInfoBySettlementNo,
    );
    const resultCombined = await syncPaymentsToWorkbook(
      sortedCombinedPayments,
      combinedApprovalInfoBySettlementNo,
      invoiceDateBySettlementId,
      previousApi10WorkbookState,
    );
    logStep('workbook', `dùng state từ ${Object.keys(previousApi10WorkbookState.items).length} bản ghi trước đó`);

    const combinedStateItems = Object.fromEntries(
      api10Payments
        .map((payment) => {
          const id = getPaymentId(payment);
          return id ? [id, stringifyPaymentSignature(payment)] : null;
        })
        .filter((entry): entry is [string, string] => entry !== null),
    );

    await saveApiState(api10StatePath, {
      meta: {
        lastRunAt: new Date().toISOString(),
        itemCount: Object.keys(combinedStateItems).length,
      },
      items: combinedStateItems,
    });

    logStep('api10-workbook', `thêm ${resultCombined.added} dòng, cập nhật ${resultCombined.updated} dòng, giữ nguyên ${resultCombined.unchanged} dòng`);
    logStep('api10-state', `đã lưu state riêng cho API 10 với ${Object.keys(combinedStateItems).length} bản ghi, thay đổi ${api10ChangedCount} bản ghi`);

    if (config.oneDrive) {
      if (createdWorkbookFromTemplate) {
        const uploadedFile = await uploadFileToOneDrivePath(outputWorkbookPath, config.oneDrive, remoteWorkbookName);
        if (!uploadedFile?.id) {
          throw new Error('OneDrive path upload did not return a file id');
        }
        await updateEnvValue('ONEDRIVE_FILE_ID', uploadedFile.id);
        logStep('onedrive', `đã cập nhật ONEDRIVE_FILE_ID=${uploadedFile.id}`);
      } else {
        const uploadedFile = await uploadFileToOneDrive(outputWorkbookPath, config.oneDrive);
        if (!uploadedFile?.id) {
          throw new Error('OneDrive upload did not return a file id');
        }
        logStep('onedrive', `đã upload workbook mới với id=${uploadedFile.id}`);
      }
      logStep('onedrive', 'đã cập nhật workbook');
    }
  } else {
    logStep('workbook', 'không có dữ liệu để tổng hợp');
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
