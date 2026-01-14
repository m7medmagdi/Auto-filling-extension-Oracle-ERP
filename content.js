window.addEventListener('ExcelDataLoaded', () => {
  chrome.storage.local.get(['excelData', 'currentRowIndex'], (result) => {
    if (!result.excelData || result.currentRowIndex === undefined) return;
    
    const row = result.excelData[result.currentRowIndex];
    if (!row) return;

    const [department, costCenter, future1, future2, barcode, checkbox] = row;

    document.querySelector("input[name*='Department']")?.value = department || "";
    document.querySelector("input[name*='costCenter']")?.value = costCenter || "";
    document.querySelector("input[name*='future1']")?.value = future1 || "";
    document.querySelector("input[name*='future2']")?.value = future2 || "";
    document.querySelector("input[name*='locationBarcode']")?.value = barcode || "";

    const cb = document.querySelector("input[name*='sbc1']");
    if (cb && !cb.checked && checkbox) cb.click();
  });
});

chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'resetForm') {
    // Clear all relevant input fields
    const inputs = document.querySelectorAll('input[type="text"], input[type="checkbox"]');
    inputs.forEach(input => {
      if(input.type === 'checkbox') {
        input.checked = false;
      } else {
        input.value = '';
      }
    });
    sendResponse({ status: 'formReset' });
  }
});
