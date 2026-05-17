/**
 * Service Worker - 专利查询助手
 * 职责：状态管理、消息路由、PDF 保存协调
 */

// ========== 状态 ==========
let state = {
  patents: [],
  currentIndex: 0,
  columnName: '',
  searchMode: 'appNo', // 'appNo' | 'pubNo'
  accumulatedFees: {},
  printedPatents: {}
};

// Service Worker 启动时恢复状态
chrome.storage.local.get('patentQueryState', (result) => {
  if (result.patentQueryState) {
    state = { ...state, ...result.patentQueryState };
  }
});

function saveState() {
  chrome.storage.local.set({ patentQueryState: state });
}

// ========== 消息处理 ==========
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  (async () => {
    try {
      switch (message.action) {
        case 'LOAD_PATENTS':
          state.patents = message.data.patents;
          state.columnName = message.data.columnName;
          state.currentIndex = 0;
          saveState();
          sendResponse({ success: true, total: state.patents.length });
          break;

        case 'GET_STATE':
          sendResponse({ success: true, data: state });
          break;

        case 'SET_INDEX':
          state.currentIndex = message.data.index;
          saveState();
          sendResponse({ success: true });
          break;

        case 'SET_SEARCH_MODE':
          state.searchMode = message.data.mode;
          saveState();
          sendResponse({ success: true });
          break;

        case 'FILL_PATENT':
          await fillPatentNumber(message.data?.index);
          sendResponse({ success: true });
          break;

        case 'FILL_AND_ADVANCE':
          await fillPatentNumber(state.currentIndex);
          sendResponse({ success: true, index: state.currentIndex });
          break;

        case 'NEXT_PATENT':
          if (state.currentIndex < state.patents.length) {
            state.currentIndex++;
          }
          saveState();
          sendResponse({ success: true, currentIndex: state.currentIndex });
          break;

        case 'PRINT_PAGE':
          await triggerPrint();
          sendResponse({ success: true });
          break;

        case 'EXTRACT_FEES':
          const feeResult = await extractFees();
          if (feeResult.success && feeResult.data) {
            const d = feeResult.data;
            const hasBasicInfo = d.basicInfo && Object.keys(d.basicInfo).length > 0;
            const hasFees = (d.feesDue && d.feesDue.length > 0) || (d.feesPaid && d.feesPaid.length > 0);
            if (!hasBasicInfo && !hasFees) {
              feeResult.error = '未识别到费用信息，请确保已打开费用查询页面';
              delete feeResult.success;
            }
          }
          sendResponse(feeResult);
          break;

        case 'EXTRACT_AND_ACCUMULATE':
          const accResult = await extractFees();
          if (accResult.success && accResult.data) {
            const data = accResult.data;
            const hasBasicInfo = data.basicInfo && Object.keys(data.basicInfo).length > 0;
            const hasFees = (data.feesDue && data.feesDue.length > 0) || (data.feesPaid && data.feesPaid.length > 0);

            if (!hasBasicInfo && !hasFees) {
              sendResponse({ error: '未识别到费用信息，请确保已打开费用查询页面' });
              break;
            }

            const patentNo = data.basicInfo?.patentNumber
              || state.patents[state.currentIndex]
              || null;
            if (!patentNo) {
              sendResponse({ error: '无法确定专利号，无法累积' });
              break;
            }

            state.accumulatedFees[patentNo] = data;
            saveState();
            accResult.accumulatedCount = Object.keys(state.accumulatedFees).length;
          }
          sendResponse(accResult);
          break;

        case 'GET_ACCUMULATED':
          sendResponse({ success: true, data: state.accumulatedFees });
          break;

        case 'CLEAR_ACCUMULATED':
          state.accumulatedFees = {};
          saveState();
          sendResponse({ success: true });
          break;

        case 'MARK_PRINTED':
          state.printedPatents[message.data.patent] = true;
          saveState();
          sendResponse({ success: true });
          break;

        case 'UNMARK_PRINTED':
          delete state.printedPatents[message.data.patent];
          saveState();
          sendResponse({ success: true });
          break;

        case 'RESET_ALL':
          state = {
            patents: [],
            currentIndex: 0,
            columnName: '',
            searchMode: 'appNo',
            accumulatedFees: {},
            printedPatents: {}
          };
          saveState();
          sendResponse({ success: true });
          break;

        default:
          sendResponse({ error: 'Unknown action' });
      }
    } catch (err) {
      console.error('[Background] Error:', err);
      sendResponse({ error: err.message });
    }
  })();
  return true;
});

// ========== 填入专利号 ==========
async function fillPatentNumber(index) {
  if (index === undefined) index = state.currentIndex;
  if (index < 0 || index >= state.patents.length) {
    throw new Error('索引超出范围');
  }

  const patent = state.patents[index];
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  if (!tab?.id) throw new Error('未找到活跃标签页');

  try {
    await chrome.tabs.sendMessage(tab.id, {
      action: 'FILL_PATENT',
      data: { patent, index, searchMode: state.searchMode }
    });
  } catch (err) {
    if (err.message.includes('Receiving end does not exist')) {
      try {
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files: ['content.js']
        });
        await chrome.tabs.sendMessage(tab.id, {
          action: 'FILL_PATENT',
          data: { patent, index, searchMode: state.searchMode }
        });
      } catch (retryErr) {
        throw new Error('无法在此页面填入专利号，请确保已打开专利查询页面');
      }
    } else {
      throw err;
    }
  }
}

// ========== 提取费用信息 ==========
async function extractFees() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return { error: '未找到活跃标签页' };

  try {
    const result = await chrome.tabs.sendMessage(tab.id, { action: 'EXTRACT_FEES' });
    return { success: true, data: result };
  } catch (err) {
    if (err.message.includes('Receiving end does not exist')) {
      try {
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          files: ['content.js']
        });
        const result = await chrome.tabs.sendMessage(tab.id, { action: 'EXTRACT_FEES' });
        return { success: true, data: result };
      } catch (retryErr) {
        return { error: '无法在此页面提取费用信息，请确保已打开费用查询页面' };
      }
    }
    return { error: err.message };
  }
}

// ========== 原生打印 ==========
async function triggerPrint() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) throw new Error('未找到活跃标签页');

  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    func: () => { window.print(); }
  });
}

// ========== 保存 PDF ==========
// ========== Side Panel 行为配置 ==========
chrome.runtime.onInstalled.addListener(() => {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true });
});

console.log('[专利费用查询助手] Service Worker 已启动');
