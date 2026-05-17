/**
 * Content Script - 注入到专利查询页面
 * 职责：接收专利号并自动填入页面输入框
 */

(function () {
  'use strict';

  if (window.__PATENT_QUERY_INJECTED__) return;
  window.__PATENT_QUERY_INJECTED__ = true;

  // 策略1：通过 Quasar q-item__label 标签文本，找同级 label[for] 指向的 input
  function findByQuasarLabel(labelText) {
    const labels = document.querySelectorAll('.q-item__label');
    for (const label of labels) {
      if (label.textContent.includes(labelText)) {
        const parent = label.closest('.row');
        if (!parent) continue;
        const forLabel = parent.querySelector('label[for]');
        if (!forLabel) continue;
        const input = document.getElementById(forLabel.getAttribute('for'));
        if (input && input.offsetParent !== null) return input;
      }
    }
    return null;
  }

  // 策略2：通过任意元素的文本内容，找附近可见的 text input
  function findByNearbyText(text) {
    const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
    while (walker.nextNode()) {
      if (!walker.currentNode.textContent.includes(text)) continue;
      const container = walker.currentNode.parentElement.closest('.row, .col, [class*="item"]');
      if (!container) continue;
      const input = container.querySelector('input[type="text"]');
      if (input && input.offsetParent !== null) return input;
    }
    return null;
  }

  // 策略3：通过 placeholder 特征匹配
  function findByPlaceholderPattern(pattern) {
    const inputs = document.querySelectorAll('input[type="text"]');
    for (const input of inputs) {
      if (input.placeholder && input.placeholder.includes(pattern) && input.offsetParent !== null) {
        return input;
      }
    }
    return null;
  }

  function findPatentInput(searchMode) {
    if (searchMode === 'pubNo') {
      return findByQuasarLabel('公开(公告)号：')
        || findByNearbyText('公开(公告)号')
        || findByPlaceholderPattern('CN1086302B')
        || findByPlaceholderPattern('公告号');
    } else {
      return findByQuasarLabel('申请号/专利号：')
        || findByNearbyText('申请号')
        || findByPlaceholderPattern('2010101995057');
    }
  }

  function fillPatentNumber(patent, searchMode) {
    const input = findPatentInput(searchMode);
    if (!input) {
      return { success: false, error: '未找到专利号输入框，请确保页面已加载完成' };
    }

    input.focus();
    input.click();

    input.value = '';
    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));

    input.value = patent;

    input.dispatchEvent(new Event('input', { bubbles: true }));
    input.dispatchEvent(new Event('change', { bubbles: true }));
    input.dispatchEvent(new Event('blur', { bubbles: true }));

    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, bubbles: true }));
    input.dispatchEvent(new KeyboardEvent('keyup', { key: 'Enter', keyCode: 13, bubbles: true }));

    return { success: true, patent, field: input.name || input.id || input.placeholder };
  }

  // ========== 提取费用信息 ==========
  function extractFeeInfo() {
    const result = { basicInfo: {}, feesDue: [], feesPaid: [] };

    // 1. 提取专利基本信息（表格形式：左列标题，右列值）
    const allTables = document.querySelectorAll('table');
    for (const table of allTables) {
      const rows = table.querySelectorAll('tr');
      for (const row of rows) {
        const cells = row.querySelectorAll('td, th');
        if (cells.length >= 2) {
          const label = cells[0].textContent.trim();
          const value = cells[1].textContent.trim();
          if (label && value) {
            const fieldMap = {
              '专利号': 'patentNumber',
              '专利申请号': 'patentNumber',
              '专利号/专利申请号': 'patentNumber',
              '专利名称': 'patentName',
              '专利类型': 'patentType',
              '申请日': 'filingDate',
              '授权公告日': 'grantDate',
              '国际分类号': 'ipcClass',
              '申请人': 'applicant',
              '发明人': 'inventor',
            };
            if (fieldMap[label]) {
              result.basicInfo[fieldMap[label]] = value;
            }
          }
        }
      }
    }

    // 2. 通过"应缴费用""已缴费用"标签定位对应的表格
    //    策略：找到标签元素，然后找其后的第一个 table
    const allElements = document.querySelectorAll('*');
    let dueTable = null;
    let paidTable = null;

    for (const el of allElements) {
      const text = el.textContent.trim();
      // 精确匹配标签文本（避免子元素干扰）
      if (el.children.length === 0 || el.childNodes.length === 1) {
        if (text === '应缴费用') {
          dueTable = findNextTable(el);
        } else if (text === '已缴费用') {
          paidTable = findNextTable(el);
        }
      }
    }

    // 3. 如果标签方式没找到，回退到表头识别
    if (!dueTable && !paidTable) {
      for (const table of allTables) {
        const headers = getTableHeaders(table);
        if (headers.includes('缴费截止日') || headers.includes('期限')) {
          dueTable = table;
        }
        if (headers.includes('缴费人') || headers.includes('收据号')) {
          paidTable = table;
        }
      }
    }

    // 4. 解析应缴费用表
    if (dueTable) {
      const rows = dueTable.querySelectorAll('tr');
      for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].querySelectorAll('td');
        if (cells.length >= 2) {
          const feeType = cells[0].textContent.trim();
          const amount = cells[1].textContent.trim();
          const deadline = cells.length >= 3 ? cells[2].textContent.trim() : '';
          if (feeType && amount) {
            result.feesDue.push({ feeType, amount, deadline });
          }
        }
      }
    }

    // 5. 解析已缴费用表
    if (paidTable) {
      const rows = paidTable.querySelectorAll('tr');
      for (let i = 1; i < rows.length; i++) {
        const cells = rows[i].querySelectorAll('td');
        if (cells.length >= 2) {
          result.feesPaid.push({
            feeType: cells[0].textContent.trim(),
            amount: cells[1].textContent.trim(),
            payer: cells.length >= 3 ? cells[2].textContent.trim() : '',
            payDate: cells.length >= 4 ? cells[3].textContent.trim() : '',
            receiptNo: cells.length >= 5 ? cells[4].textContent.trim() : ''
          });
        }
      }
    }

    return result;
  }

  // 查找元素之后的第一个 table
  function findNextTable(el) {
    let node = el;
    while (node) {
      node = node.nextElementSibling;
      if (!node) {
        // 向上找父元素的兄弟
        node = el.parentElement;
        if (!node) break;
        node = node.nextElementSibling;
        if (!node) break;
      }
      // 在当前元素及其子元素中找 table
      const table = node.querySelector('table') || (node.tagName === 'TABLE' ? node : null);
      if (table) return table;
      // 最多向上找 5 层
    }
    return null;
  }

  function getTableHeaders(table) {
    const headerRow = table.querySelector('tr');
    if (!headerRow) return [];
    return Array.from(headerRow.querySelectorAll('th, td')).map(c => c.textContent.trim());
  }

  // ========== 消息监听 ==========
  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'FILL_PATENT') {
      sendResponse(fillPatentNumber(message.data.patent, message.data.searchMode));
    }
    if (message.action === 'EXTRACT_FEES') {
      sendResponse(extractFeeInfo());
    }
    if (message.action === 'PING') {
      sendResponse({ success: true });
    }
    return true;
  });

  console.log('[专利查询助手] Content Script 已加载');
})();
