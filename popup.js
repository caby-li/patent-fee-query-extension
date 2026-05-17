/**
 * Popup 弹出窗逻辑
 */

async function checkStatus() {
  try {
    const resp = await chrome.runtime.sendMessage({ action: 'GET_STATE' });
    if (resp?.success) {
      const { patents, currentIndex } = resp.data;
      document.getElementById('totalCount').textContent = patents.length;
      document.getElementById('doneCount').textContent = Math.min(currentIndex, patents.length);
    }
  } catch (e) {
    // 静默
  }
}

document.getElementById('btnOpenPanel').addEventListener('click', async () => {
  await chrome.sidePanel.open({ windowId: (await chrome.windows.getCurrent()).id });
  window.close();
});

checkStatus();
