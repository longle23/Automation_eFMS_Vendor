import { mkdir, readFile, writeFile, copyFile, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { spawn } from 'node:child_process';
import { EfmsClient } from './client.js';
import { getAccessToken } from './auth.js';
import { loadConfig } from './config.js';

export { EfmsClient } from './client.js';
export { getAccessToken } from './auth.js';
export { loadConfig } from './config.js';

const isDirectRun = import.meta.url === pathToFileURL(process.argv[1] ?? '').href;
const dataDir = join(process.cwd(), 'data');
const api1OutputPath = join(dataDir, 'api1-response.json');
const api2OutputPath = join(dataDir, 'api2-response.json');
const templateWorkbookPath = join(dataDir, 'Vendor_Payment_Template.xlsx');
const outputWorkbookPath = join(dataDir, 'Vendor_Payment_Output.xlsx');
const statePath = join(dataDir, 'Vendor_Payment_Output.state.json');
const intervalMs = 30 * 60 * 1000;
const worksheetName = 'Vendor_Payment';

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

type OutputState = {
  seenIds: string[];
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

async function loadState(): Promise<OutputState> {
  try {
    const contents = await readFile(statePath, 'utf8');
    const parsed = JSON.parse(contents) as Partial<OutputState>;
    return {
      seenIds: Array.isArray(parsed.seenIds)
        ? parsed.seenIds.filter((id): id is string => typeof id === 'string')
        : [],
    };
  } catch {
    return { seenIds: [] };
  }
}

async function saveState(state: OutputState) {
  await ensureDataDir();
  await writeFile(statePath, `${JSON.stringify(state, null, 2)}\n`, 'utf8');
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

async function appendPaymentsToWorkbook(payments: SettlementPayment[]) {
  await ensureDataDir();
  await ensureWorkbookExists();

  const rows = payments
    .map((payment) => ({
      h: payment.payeeName ?? '',
      i: payment.settlementNo ?? '',
    }))
    .filter((row) => row.h || row.i);

  if (!rows.length) {
    return;
  }

  const tempScriptPath = join(dataDir, 'append-payments.py');
  const script = `from openpyxl import load_workbook\npath = r'''${outputWorkbookPath}'''\nwb = load_workbook(path)\nws = wb['${worksheetName}']\nrows = ${JSON.stringify(rows, null, 2)}\nfor row in rows:\n    ws.append([None, None, None, None, None, None, None, row['h'], row['i']])\nwb.save(path)\n`;
  await writeFile(tempScriptPath, script, 'utf8');

  try {
    await new Promise<void>((resolve, reject) => {
      const child = spawn('python', [tempScriptPath], { stdio: 'inherit' });
      child.on('exit', (code) => {
        if (code === 0) resolve();
        else reject(new Error(`Failed to append workbook rows (exit code ${code ?? 'unknown'})`));
      });
    });
  } finally {
    await unlink(tempScriptPath).catch(() => undefined);
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

  const state = await loadState();
  const seenIds = new Set(state.seenIds);
  const incomingRows = iteratePaymentsFromLastToFirst(
    (data.data ?? []).filter((item) => item.id && !seenIds.has(item.id)),
  );

  if (incomingRows.length) {
    await appendPaymentsToWorkbook(incomingRows);
    for (const row of incomingRows) {
      if (row.id) seenIds.add(row.id);
    }
    await saveState({ seenIds: [...seenIds] });
    logStep('workbook', `cập nhật ${incomingRows.length} dòng mới`);
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
  void startScheduler();
}
