let cleanButton;
let statusBox;

Office.onReady((info) => {
  cleanButton = document.getElementById("cleanButton");
  statusBox = document.getElementById("status");

  if (info.host !== Office.HostType.Excel) {
    setStatus("این افزونه باید داخل Microsoft Excel باز شود.", true);
    cleanButton.disabled = true;
    return;
  }

  setStatus("متصل شد. یک محدوده از سلول‌ها را انتخاب کن.");
  cleanButton.addEventListener("click", cleanSelectedRange);
});

async function cleanSelectedRange() {
  cleanButton.disabled = true;
  setStatus("در حال پاکسازی داده‌های انتخاب‌شده...");

  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();

      const cleanedValues = range.values.map((row) =>
        row.map((value) => {
          if (typeof value !== "string") return value;

          return value
            .replace(/\u00A0/g, " ")
            .replace(/\s+/g, " ")
            .trim();
        })
      );

      range.values = cleanedValues;
      await context.sync();
    });

    setStatus("پاکسازی با موفقیت انجام شد.");
  } catch (error) {
    console.error(error);
    setStatus(`خطا: ${error.message}`, true);
  } finally {
    cleanButton.disabled = false;
  }
}

function setStatus(message, isError = false) {
  statusBox.textContent = message;
  statusBox.style.background = isError ? "#fff1f2" : "#174ea6";
  statusBox.style.color = isError ? "#b42318" : "#174ea6";
}