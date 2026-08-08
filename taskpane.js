/* global Office, Excel */

Office.onReady(function (info) {
  if (info.host !== Office.HostType.Excel) {
    showMessage(
      "Add-in ini hanya dapat digunakan di Excel.",
      "error"
    );
    return;
  }

  setDefaultDate();

  document
    .getElementById("qty")
    .addEventListener("input", calculateSubtotal);

  document
    .getElementById("harga")
    .addEventListener("input", calculateSubtotal);

  document
    .getElementById("purchaseForm")
    .addEventListener("submit", saveToExcel);

  document
    .getElementById("resetButton")
    .addEventListener("click", resetForm);

  loadMasterData();
});

function setDefaultDate() {
  const dateInput = document.getElementById("tanggal");
  const today = new Date();

  const year = today.getFullYear();
  const month = String(
    today.getMonth() + 1
  ).padStart(2, "0");
  const day = String(
    today.getDate()
  ).padStart(2, "0");

  dateInput.value = `${year}-${month}-${day}`;
}

function calculateSubtotal() {
  const qty = Number(
    document.getElementById("qty").value
  ) || 0;

  const harga = Number(
    document.getElementById("harga").value
  ) || 0;

  document.getElementById("subtotal").value =
    qty * harga;
}

async function loadMasterData() {
  try {
    await Excel.run(async function (context) {
      const supplierSheet = context.workbook.worksheets
        .getItemOrNullObject("Supplier");

      const productSheet = context.workbook.worksheets
        .getItemOrNullObject("Produk");

      supplierSheet.load("isNullObject");
      productSheet.load("isNullObject");

      await context.sync();

      if (!supplierSheet.isNullObject) {
        const supplierRange = supplierSheet
          .getRange("A2:A1000");

        supplierRange.load("values");
        await context.sync();

        fillDatalist(
          "supplierList",
          supplierRange.values
        );
      }

      if (!productSheet.isNullObject) {
        const productRange = productSheet
          .getRange("A2:A1000");

        productRange.load("values");
        await context.sync();

        fillDatalist(
          "produkList",
          productRange.values
        );
      }
    });

    showMessage(
      "Data supplier dan produk berhasil dimuat.",
      "success"
    );
  } catch (error) {
    showMessage(
      "Gagal memuat master: " + error.message,
      "error"
    );
  }
}

function fillDatalist(listId, values) {
  const datalist = document.getElementById(listId);
  datalist.innerHTML = "";

  const uniqueValues = new Set();

  values.forEach(function (row) {
    const value = row[0];

    if (
      value !== null &&
      value !== undefined &&
      String(value).trim() !== ""
    ) {
      uniqueValues.add(String(value).trim());
    }
  });

  uniqueValues.forEach(function (value) {
    const option = document.createElement("option");
    option.value = value;
    datalist.appendChild(option);
  });
}

function getFormData() {
  const data = {
    tanggal: document.getElementById("tanggal").value,
    noBeli: document.getElementById("noBeli").value.trim(),
    supplierId: document
      .getElementById("supplierId")
      .value.trim(),
    produkId: document
      .getElementById("produkId")
      .value.trim(),
    qty: Number(document.getElementById("qty").value),
    harga: Number(
      document.getElementById("harga").value
    ),
    metodeBayar: document
      .getElementById("metodeBayar")
      .value,
    statusBayar: document
      .getElementById("statusBayar")
      .value
  };

  if (!data.tanggal) {
    throw new Error("Tanggal wajib diisi.");
  }

  if (!data.noBeli) {
    throw new Error("No. beli wajib diisi.");
  }

  if (!data.supplierId) {
    throw new Error("Supplier ID wajib diisi.");
  }

  if (!data.produkId) {
    throw new Error("Produk ID wajib diisi.");
  }

  if (data.qty <= 0) {
    throw new Error("Qty harus lebih besar dari nol.");
  }

  if (data.harga <= 0) {
    throw new Error("Harga harus lebih besar dari nol.");
  }

  data.subtotal = data.qty * data.harga;

  return data;
}

async function saveToExcel(event) {
  event.preventDefault();

  try {
    const data = getFormData();

    await Excel.run(async function (context) {
      const sheet = context.workbook.worksheets
        .getActiveWorksheet();

      const usedRange = sheet.getUsedRangeOrNullObject();

      usedRange.load([
        "isNullObject",
        "rowIndex",
        "rowCount"
      ]);

      await context.sync();

      let rowIndex = 1;

      if (!usedRange.isNullObject) {
        rowIndex = Math.max(
          1,
          usedRange.rowIndex + usedRange.rowCount
        );
      }

      const values = [[
        data.tanggal,
        data.noBeli,
        data.supplierId,
        data.produkId,
        data.qty,
        data.harga,
        data.subtotal,
        data.metodeBayar,
        data.statusBayar
      ]];

      const targetRange = sheet.getRangeByIndexes(
        rowIndex,
        0,
        1,
        9
      );

      targetRange.values = values;
      targetRange.numberFormat = [[
        "dd/mm/yyyy",
        "@",
        "@",
        "@",
        "0",
        "#,##0.00",
        "#,##0.00",
        "@",
        "@"
      ]];

      await context.sync();

      showMessage(
        `Data tersimpan pada baris ${rowIndex + 1}.`,
        "success"
      );
    });

    resetForm();
  } catch (error) {
    showMessage(
      "Gagal menyimpan: " + error.message,
      "error"
    );
  }
}

function resetForm() {
  document.getElementById("noBeli").value = "";
  document.getElementById("supplierId").value = "";
  document.getElementById("produkId").value = "";
  document.getElementById("qty").value = 1;
  document.getElementById("harga").value = 0;
  document.getElementById("subtotal").value = 0;
  document.getElementById("metodeBayar").value = "Tunai";
  document.getElementById("statusBayar").value = "Lunas";

  setDefaultDate();
}

function showMessage(text, type) {
  const message = document.getElementById("message");

  message.textContent = text;
  message.className = `show ${type}`;
}