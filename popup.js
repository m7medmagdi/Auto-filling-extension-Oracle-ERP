document.addEventListener('DOMContentLoaded', () => {
  const fileInput = document.getElementById('excelFile');
  const fillBtn = document.getElementById('fillBtn');
  const submitBtn = document.getElementById('submitBtn');
  const submitFullBtn = document.getElementById('submitFullBtn');
  const resetBtn = document.getElementById('resetBtn');
  const currentRowDisplay = document.getElementById('currentRow');

  // Load saved state
  chrome.storage.local.get(['excelData', 'currentRowIndex'], (result) => {
    if (result.currentRowIndex !== undefined) {
      currentRowDisplay.textContent = result.currentRowIndex + 1;
    }
    if (result.excelData) {
      fileInput.disabled = true; // Disable file input if data is already loaded
    }
  });

  // Handle file upload
  fileInput.addEventListener('change', () => {
    if (!fileInput.files.length) return alert("Please select an Excel file");
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      chrome.storage.local.set({ excelData: rows, currentRowIndex: 1 }, () => {
        currentRowDisplay.textContent = 1;
        fileInput.disabled = true;
      });
    };
    reader.readAsArrayBuffer(fileInput.files[0]);
  });

  // Function to fill form and optionally click Add Row
  const fillForm = (rowData, tabId, clickAddRow, callback) => {
    chrome.scripting.executeScript({
      target: { tabId: tabId },
      args: [rowData, clickAddRow],
      func: (data, clickAddRow) => {
        // Fill form fields first
        const fields = document.querySelectorAll('input');
        fields.forEach(input => {
          const name = input.name || "";
          if (name.includes("Department")) input.value = data[0] || "";
          if (name.includes("costCenter")) input.value = data[1] || "";
          if (name.includes("future1")) input.value = data[2] || "";
          if (name.includes("future2")) input.value = data[3] || "";
          if (name.includes("locationBarcode")) input.value = data[4] || "";
          if (name.includes("sbc1")) input.checked = !!data[5];
        });

        // Click Add Row button if requested (after filling)
        if (clickAddRow) {
          setTimeout(() => {
            const addRowButton = document.querySelector("img[title='Add Row']");
            if (addRowButton) {
              addRowButton.click();
            } else {
              alert("Add Row button not found");
              throw new Error("Add Row button not found"); // Stop execution on error
            }
          }, 500); // 500ms delay after filling fields
        }
      }
    }, (results) => {
      // Check for errors in script execution
      if (chrome.runtime.lastError) {
        console.error(chrome.runtime.lastError.message);
        alert("Error executing script: " + chrome.runtime.lastError.message);
        return;
      }
      callback();
    });
  };

  // Handle fill button click (single row, fill only)
  fillBtn.addEventListener('click', () => {
    chrome.storage.local.get(['excelData', 'currentRowIndex'], (result) => {
      if (!result.excelData) return alert("Please select an Excel file first");
      
      const rows = result.excelData;
      let currentRowIndex = result.currentRowIndex || 1;
      if (currentRowIndex >= rows.length) {
        alert("Reached the end of the data");
        return;
      }

      const rowData = rows[currentRowIndex];
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        fillForm(rowData, tabs[0].id, false, () => {
          chrome.storage.local.set({ currentRowIndex: currentRowIndex + 1 }, () => {
            currentRowDisplay.textContent = currentRowIndex + 1;
          });
        });
      });
    });
  });

  // Handle submit button click (single row, fill and add)
  submitBtn.addEventListener('click', () => {
    chrome.storage.local.get(['excelData', 'currentRowIndex'], (result) => {
      if (!result.excelData) return alert("Please select an Excel file first");

      const rows = result.excelData;
      let currentRowIndex = result.currentRowIndex || 1;
      if (currentRowIndex >= rows.length) {
        alert("Reached the end of the data");
        return;
      }

      const rowData = rows[currentRowIndex];
      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        fillForm(rowData, tabs[0].id, true, () => {
          chrome.storage.local.set({ currentRowIndex: currentRowIndex + 1 }, () => {
            currentRowDisplay.textContent = currentRowIndex + 1;
          });
        });
      });
    });
  });

  // Handle submit full button click (all rows)
  submitFullBtn.addEventListener('click', () => {
    chrome.storage.local.get(['excelData', 'currentRowIndex'], (result) => {
      if (!result.excelData) return alert("Please select an Excel file first");

      const rows = result.excelData;
      let currentRowIndex = result.currentRowIndex || 1;
      if (currentRowIndex >= rows.length) {
        alert("Reached the end of the data");
        return;
      }

      chrome.tabs.query({ active: true, currentWindow: true }, (tabs) => {
        const tabId = tabs[0].id;

        // Function to process rows sequentially with delay
        const processRow = (index) => {
          if (index >= rows.length) {
            chrome.storage.local.set({ currentRowIndex: index }, () => {
              currentRowDisplay.textContent = index;
              alert("All rows have been processed");
            });
            return;
          }

          const rowData = rows[index];
          fillForm(rowData, tabId, true, () => {
            chrome.storage.local.set({ currentRowIndex: index + 1 }, () => {
              currentRowDisplay.textContent = index + 1;
              // Process next row after 1500ms delay
              setTimeout(() => {
                processRow(index + 1);
              }, 1500);
            });
          });
        };

        // Start processing from the current row
        processRow(currentRowIndex);
      });
    });
  });

  // Handle reset button click
  resetBtn.addEventListener('click', () => {
    chrome.storage.local.clear(() => {
      fileInput.disabled = false;
      currentRowDisplay.textContent = 1;
      alert("Extension reset to default state");
    });
  });
});