/**
 * ============================================================================
 * 📱 TALHA TECH - COMPLETE ENTERPRISE POS & ERP BACKEND ENGINE
 * 🏢 কারওয়ান বাজার • বসুন্ধরা • পান্থপথ | প্রো: মোঃ আবু তালহা
 * 🔗 Spreadsheet ID: 1DRBYeIUws4sbcfMRrDUfFVITLvEOvUyjnfgE1U6rx3M
 * ============================================================================
 */

const SPREADSHEET_ID = "1DRBYeIUws4sbcfMRrDUfFVITLvEOvUyjnfgE1U6rx3M";

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== "") {
    try {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      return SpreadsheetApp.getActiveSpreadsheet();
    }
  }
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ============================================================================
// 🎨 ১. স্বয়ংক্রিয় ডাটাবেজ সেটআপ ও প্রিমিয়াম স্টাইলিং (এক ক্লিকে সাজিয়ে দেবে)
// ============================================================================
function setupCompleteDatabase() {
  const ss = getSpreadsheet();
  
  // শিটের তালিকা, কলাম হেডার ও রয়্যাল কালার থিম
  const schemas = [
    {
      name: "Products",
      headers: ["Product ID", "Product Name", "Category", "Stock Qty", "Buy Price (৳)", "Sell Price (৳)", "Reorder Level", "Supplier", "Notes", "Last Updated"],
      theme: { bg: "#083344", text: "#38bdf8", accent: "#e0f2fe", widths: [130, 220, 130, 90, 110, 110, 100, 140, 180, 120] }
    },
    {
      name: "Sales",
      headers: ["Invoice ID", "Date", "Time", "Customer Name", "Phone", "Items Summary", "Subtotal (৳)", "Discount (৳)", "Delivery (৳)", "Total Amount (৳)", "Paid (৳)", "Due (৳)", "Profit (৳)", "Payment Method", "Status", "Notes"],
      theme: { bg: "#1e1b4b", text: "#818cf8", accent: "#eef2ff", widths: [140, 100, 90, 160, 120, 260, 100, 100, 90, 120, 100, 100, 110, 110, 100, 160] }
    },
    {
      name: "Expenses",
      headers: ["Expense ID", "Date", "Time", "Expense Name", "Category", "Amount (৳)", "Account", "Type", "Note"],
      theme: { bg: "#4c0519", text: "#fb7185", accent: "#ffe4e6", widths: [130, 100, 90, 180, 150, 120, 110, 110, 200] }
    },
    {
      name: "Damages",
      headers: ["Damage ID", "Date", "Time", "Product ID", "Product Name", "Qty (Pcs)", "Unit Buy Price (৳)", "Total Loss (৳)", "Reason", "Notes"],
      theme: { bg: "#450a0a", text: "#f87171", accent: "#fee2e2", widths: [130, 100, 90, 130, 200, 90, 120, 120, 140, 180] }
    },
    {
      name: "Dues",
      headers: ["Due ID", "Date", "Due Date", "Customer Name", "Phone", "Invoice ID", "Total Due (৳)", "Paid Amount (৳)", "Remaining (৳)", "Status", "Payment History", "Notes"],
      theme: { bg: "#451a03", text: "#fbbf24", accent: "#fef3c7", widths: [130, 100, 100, 160, 120, 130, 110, 110, 110, 100, 250, 180] }
    },
    {
      name: "Supplier_Payables",
      headers: ["Payable ID", "Date", "Due Date", "Supplier Name", "Phone", "Challan / Inv", "Total Amount (৳)", "Paid Amount (৳)", "Balance Due (৳)", "Status", "Payments JSON", "Notes"],
      theme: { bg: "#2e1065", text: "#c084fc", accent: "#f3e8ff", widths: [130, 100, 100, 180, 120, 120, 120, 120, 120, 100, 240, 180] }
    },
    {
      name: "Categories",
      headers: ["Category Name", "Description", "Active Status"],
      theme: { bg: "#064e3b", text: "#34d399", accent: "#d1fae5", widths: [180, 240, 110] }
    },
    {
      name: "Cash_Accounts",
      headers: ["Account ID", "Account Name", "Icon", "Starting Balance (৳)", "Current Balance (৳)", "Type"],
      theme: { bg: "#0f172a", text: "#38bdf8", accent: "#f1f5f9", widths: [120, 160, 80, 140, 140, 120] }
    }
  ];

  schemas.forEach(schema => {
    let sheet = ss.getSheetByName(schema.name);
    if (!sheet) {
      sheet = ss.insertSheet(schema.name);
    }

    // হেডার সেটআপ
    const headerRange = sheet.getRange(1, 1, 1, schema.headers.length);
    headerRange.setValues([schema.headers]);
    headerRange.setBackground(schema.theme.bg);
    headerRange.setFontColor(schema.theme.text);
    headerRange.setFontFamily("Arial");
    headerRange.setFontSize(10);
    headerRange.setFontWeight("bold");
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 40);

    // হেডার ফ্রিজ
    sheet.setFrozenRows(1);

    // কলাম উইডথ সেটআপ
    schema.theme.widths.forEach((w, idx) => {
      sheet.setColumnWidth(idx + 1, w);
    });

    // অল্টারনেটিং রো কালার (Banding)
    const existingBandings = sheet.getBandings();
    if (existingBandings.length > 0) {
      existingBandings.forEach(b => b.remove());
    }
    
    sheet.getRange(2, 1, 100, schema.headers.length)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);

    // ডিফল্ট ক্যাটাগরি ডেটা যোগ (যদি ফাঁকা থাকে)
    if (schema.name === "Categories" && sheet.getLastRow() === 1) {
      sheet.getRange(2, 1, 5, 3).setValues([
        ["চার্জিং কেবল", "মোবাইল চার্জিং ক্যাবলস", "Active"],
        ["চার্জার", "ফাস্ট চার্জার ও এডাপ্টার", "Active"],
        ["গ্লাস/কভার", "স্ক্রিন প্রটেক্টর ও ব্যাক কভার", "Active"],
        ["ইয়ারফোন/TWS", "ব্লুটুথ ও তারের ইয়ারফোন", "Active"],
        ["গ্যাজেট", "অন্যান্য স্মার্ট এক্সেসরিজ", "Active"]
      ]);
    }

    // ডিফল্ট অ্যাকাউন্টস যোগ (যদি ফাঁকা থাকে)
    if (schema.name === "Cash_Accounts" && sheet.getLastRow() === 1) {
      sheet.getRange(2, 1, 5, 6).setValues([
        ["cash", "নগদ ক্যাশ", "💵", 0, 0, "Cash"],
        ["bkash", "বিকাশ", "📱", 0, 0, "MFS"],
        ["nagad", "নগদ (Nagad)", "📲", 0, 0, "MFS"],
        ["rocket", "রকেট", "🚀", 0, 0, "MFS"],
        ["bank", "ব্যাংক অ্যাকাউন্ট", "🏦", 0, 0, "Bank"]
      ]);
    }
  });

  // ডিফল্ট 'Sheet1' ডিলিট (যদি থাকে)
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e){}
  }

  Logger.log("✅ Talha Tech ডাটাবেজ সফলভাবে সেটআপ ও ডিজাইন সম্পন্ন হয়েছে!");
}

// ============================================================================
// ⚡ ২. হাই-স্পিড ডেটা ফেচিং API (doGet - JSONP & JSON Support)
// ============================================================================
function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(5000);
  
  try {
    const ss = getSpreadsheet();
    
    // ১. প্রোডাক্টস লোড
    const pSheet = ss.getSheetByName("Products");
    const pData = pSheet && pSheet.getLastRow() > 1 
      ? pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 10).getValues() : [];
    
    const products = pData.map(r => ({
      id: String(r[0] || ""),
      name: String(r[1] || ""),
      category: String(r[2] || "গ্যাজেট"),
      qty: Number(r[3]) || 0,
      buyPrice: Number(r[4]) || 0,
      sellPrice: Number(r[5]) || 0,
      reorderLevel: Number(r[6]) || 5,
      supplier: String(r[7] || ""),
      notes: String(r[8] || ""),
      lastUpdated: String(r[9] || "")
    })).filter(p => p.name !== "");

    // ২. সেলস লোড
    const sSheet = ss.getSheetByName("Sales");
    const sData = sSheet && sSheet.getLastRow() > 1 
      ? sSheet.getRange(2, 1, sSheet.getLastRow() - 1, 16).getValues() : [];
      
    const sales = sData.map(r => {
      let items = [];
      try { items = JSON.parse(r[5]); } catch(err) { items = [{ name: String(r[5]||""), qty: 1, unitPrice: Number(r[9])||0 }]; }
      return {
        invoiceId: String(r[0] || ""),
        date: formatDate(r[1]),
        time: String(r[2] || ""),
        customerName: String(r[3] || "সাধারণ ক্রেতা"),
        customerPhone: String(r[4] || ""),
        items: Array.isArray(items) ? items : [],
        subtotal: Number(r[6]) || 0,
        discount: { amount: Number(r[7]) || 0 },
        deliveryCharge: Number(r[8]) || 0,
        totalAmount: Number(r[9]) || 0,
        amountPaid: Number(r[10]) || 0,
        dueAmount: Number(r[11]) || 0,
        profit: Number(r[12]) || 0,
        paymentMethod: String(r[13] || "Cash"),
        isReturned: String(r[14] || "").toLowerCase() === "returned",
        notes: String(r[15] || "")
      };
    });

    // ৩. এক্সপেন্সেস লোড
    const eSheet = ss.getSheetByName("Expenses");
    const eData = eSheet && eSheet.getLastRow() > 1 
      ? eSheet.getRange(2, 1, eSheet.getLastRow() - 1, 9).getValues() : [];
      
    const expenses = eData.map(r => ({
      id: String(r[0] || ""),
      date: formatDate(r[1]),
      time: String(r[2] || ""),
      name: String(r[3] || ""),
      category: String(r[4] || "অন্যান্য"),
      amount: Number(r[5]) || 0,
      cashAccount: String(r[6] || "cash"),
      type: String(r[7] || "business"),
      note: String(r[8] || "")
    }));

    // ৪. ড্যামেজ লোড
    const dSheet = ss.getSheetByName("Damages");
    const dData = dSheet && dSheet.getLastRow() > 1 
      ? dSheet.getRange(2, 1, dSheet.getLastRow() - 1, 10).getValues() : [];
      
    const damages = dData.map(r => ({
      id: String(r[0] || ""),
      date: formatDate(r[1]),
      time: String(r[2] || ""),
      productId: String(r[3] || ""),
      productName: String(r[4] || ""),
      qty: Number(r[5]) || 0,
      buyPrice: Number(r[6]) || 0,
      lossAmount: Number(r[7]) || 0,
      reason: String(r[8] || ""),
      notes: String(r[9] || "")
    }));

    // ৫. কাস্টমার বাকি (Dues) লোড
    const dueSheet = ss.getSheetByName("Dues");
    const dueData = dueSheet && dueSheet.getLastRow() > 1 
      ? dueSheet.getRange(2, 1, dueSheet.getLastRow() - 1, 12).getValues() : [];
      
    const dues = dueData.map(r => {
      let pHistory = [];
      try { pHistory = JSON.parse(r[10]); } catch(e){}
      return {
        id: String(r[0] || ""),
        date: formatDate(r[1]),
        dueDate: formatDate(r[2]),
        customer: String(r[3] || ""),
        phone: String(r[4] || ""),
        invoiceId: String(r[5] || ""),
        amount: Number(r[6]) || 0,
        paidAmount: Number(r[7]) || 0,
        payments: Array.isArray(pHistory) ? pHistory : [],
        isPaid: String(r[9] || "") === "পরিশোধিত",
        note: String(r[11] || "")
      };
    });

    // ৬. ক্যাটাগরি তালিকা
    const cSheet = ss.getSheetByName("Categories");
    const cData = cSheet && cSheet.getLastRow() > 1 
      ? cSheet.getRange(2, 1, cSheet.getLastRow() - 1, 1).getValues() : [];
    const categories = cData.map(r => String(r[0])).filter(Boolean);

    const payload = {
      status: "success",
      businessName: "Talha Tech",
      timestamp: new Date().toISOString(),
      products: products,
      sales: sales,
      expenses: expenses,
      damages: damages,
      dues: dues,
      categories: categories.length ? categories : ["চার্জিং কেবল", "চার্জার", "গ্লাস/কভার", "ইয়ারফোন/TWS", "গ্যাজেট"]
    };

    const output = JSON.stringify(payload);
    const callback = e && e.parameter && e.parameter.callback;
    
    if (callback) {
      return ContentService.createTextOutput(`${callback}(${output})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    
    return ContentService.createTextOutput(output)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    const errOutput = JSON.stringify({ status: "error", message: error.toString() });
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService.createTextOutput(`${callback}(${errOutput})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errOutput)
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ============================================================================
// 🚀 ৩. হাই-পারফরম্যান্স অ্যাকশন হ্যান্ডলার (doPost - Realtime Save/Update)
// ============================================================================
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    const postData = JSON.parse(e.postData.contents || "{}");
    const action = postData.action;
    const ss = getSpreadsheet();

    // -------------------------------------------------------------
    // ACTION: বিক্রি সেভ এবং অটো-স্টক মাইনাস (saveSale)
    // -------------------------------------------------------------
    if (action === "saveSale") {
      const sale = postData.sale;
      const salesSheet = ss.getSheetByName("Sales");
      
      const itemsJson = JSON.stringify(sale.items || []);
      salesSheet.appendRow([
        sale.invoiceId,
        sale.date || todayStr(),
        sale.time || "",
        sale.customerName || "সাধারণ ক্রেতা",
        sale.customerPhone || "",
        itemsJson,
        sale.subtotal || 0,
        (sale.discount && sale.discount.amount) || 0,
        sale.deliveryCharge || 0,
        sale.totalAmount || 0,
        sale.amountPaid || 0,
        sale.dueAmount || 0,
        sale.profit || 0,
        sale.paymentMethod || "Cash",
        "Completed",
        sale.notes || ""
      ]);

      // স্টক আপডেট
      if (Array.isArray(sale.items)) {
        const prodSheet = ss.getSheetByName("Products");
        if (prodSheet && prodSheet.getLastRow() > 1) {
          const range = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 4);
          const values = range.getValues();
          let modified = false;

          sale.items.forEach(soldItem => {
            for (let i = 0; i < values.length; i++) {
              if (String(values[i][0]) === String(soldItem.productId) || 
                  String(values[i][1]).toLowerCase() === String(soldItem.name).toLowerCase()) {
                values[i][3] = Math.max(0, (Number(values[i][3]) || 0) - (Number(soldItem.qty) || 0));
                modified = true;
                break;
              }
            }
          });

          if (modified) range.setValues(values);
        }
      }

      // যদি বাকি বিক্রি হয়
      if (sale.paymentMethod === "Due") {
        const duesSheet = ss.getSheetByName("Dues");
        duesSheet.appendRow([
          "due_" + Date.now(),
          sale.date || todayStr(),
          "",
          sale.customerName || "সাধারণ ক্রেতা",
          sale.customerPhone || "",
          sale.invoiceId,
          sale.totalAmount,
          0,
          sale.totalAmount,
          "অপরিশোধিত",
          "[]",
          sale.notes || ""
        ]);
      }

      return jsonResponse({ status: "success", message: "Sale recorded successfully!" });
    }

    // -------------------------------------------------------------
    // ACTION: বাল্ক পণ্য এন্ট্রি (saveBulkProducts)
    // -------------------------------------------------------------
    if (action === "saveBulkProducts") {
      const products = postData.products || [];
      const prodSheet = ss.getSheetByName("Products");
      const existingData = prodSheet.getLastRow() > 1 
        ? prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 10).getValues() : [];

      const map = {};
      existingData.forEach((r, idx) => {
        map[String(r[1]).toLowerCase().trim()] = idx; // By Name
      });

      products.forEach(p => {
        const key = String(p.name).toLowerCase().trim();
        if (map.hasOwnProperty(key)) {
          const rowIdx = map[key];
          existingData[rowIdx][3] = Number(existingData[rowIdx][3] || 0) + Number(p.qty || 0); // স্টক যোগ
          existingData[rowIdx][4] = Number(p.buyPrice || existingData[rowIdx][4]);
          existingData[rowIdx][5] = Number(p.sellPrice || existingData[rowIdx][5]);
          existingData[rowIdx][2] = p.category || existingData[rowIdx][2];
          existingData[rowIdx][9] = todayStr();
        } else {
          existingData.push([
            p.id || "prod_" + Date.now() + "_" + Math.floor(Math.random()*100),
            p.name,
            p.category || "গ্যাজেট",
            Number(p.qty) || 0,
            Number(p.buyPrice) || 0,
            Number(p.sellPrice) || 0,
            Number(p.reorderLevel) || 5,
            p.supplier || "",
            p.notes || "",
            todayStr()
          ]);
        }
      });

      if (existingData.length > 0) {
        prodSheet.getRange(2, 1, existingData.length, 10).setValues(existingData);
      }

      return jsonResponse({ status: "success", count: products.length });
    }

    // -------------------------------------------------------------
    // ACTION: একক পণ্য আপডেট (updateProduct)
    // -------------------------------------------------------------
    if (action === "updateProduct") {
      const p = postData.product;
      const prodSheet = ss.getSheetByName("Products");
      if (prodSheet && prodSheet.getLastRow() > 1) {
        const range = prodSheet.getRange(2, 1, prodSheet.getLastRow() - 1, 10);
        const values = range.getValues();
        let found = false;

        for (let i = 0; i < values.length; i++) {
          if (String(values[i][0]) === String(p.id) || String(values[i][1]).toLowerCase() === String(p.name).toLowerCase()) {
            values[i][1] = p.name;
            values[i][2] = p.category || values[i][2];
            values[i][3] = Number(p.qty) || 0;
            values[i][4] = Number(p.buyPrice) || 0;
            values[i][5] = Number(p.sellPrice) || 0;
            values[i][6] = Number(p.reorderLevel) || 5;
            values[i][7] = p.supplier || "";
            values[i][8] = p.notes || "";
            values[i][9] = todayStr();
            found = true;
            break;
          }
        }

        if (found) {
          range.setValues(values);
        } else {
          prodSheet.appendRow([p.id, p.name, p.category, p.qty, p.buyPrice, p.sellPrice, p.reorderLevel || 5, p.supplier || "", p.notes || "", todayStr()]);
        }
      }
      return jsonResponse({ status: "success" });
    }

    // -------------------------------------------------------------
    // ACTION: পণ্য ডিলিট (deleteProduct & deleteProducts)
    // -------------------------------------------------------------
    if (action === "deleteProduct" || action === "deleteProducts") {
      const prodIds = postData.productIds || [postData.productId];
      const prodSheet = ss.getSheetByName("Products");
      if (prodSheet && prodSheet.getLastRow() > 1) {
        const numRows = prodSheet.getLastRow() - 1;
        const ids = prodSheet.getRange(2, 1, numRows, 1).getValues();

        for (let i = ids.length - 1; i >= 0; i--) {
          if (prodIds.includes(String(ids[i][0]))) {
            prodSheet.deleteRow(i + 2);
          }
        }
      }
      return jsonResponse({ status: "success" });
    }

    // -------------------------------------------------------------
    // ACTION: ড্যামেজ এন্ট্রি (saveDamage)
    // -------------------------------------------------------------
    if (action === "saveDamage") {
      const dmg = postData.damage;
      const dmgSheet = ss.getSheetByName("Damages");
      dmgSheet.appendRow([
        dmg.id,
        dmg.date || todayStr(),
        dmg.time || "",
        dmg.productId || "",
        dmg.productName || "",
        dmg.qty || 1,
        dmg.buyPrice || 0,
        dmg.lossAmount || 0,
        dmg.reason || "",
        dmg.notes || ""
      ]);
      return jsonResponse({ status: "success" });
    }

    // -------------------------------------------------------------
    // ACTION: খরচ এন্ট্রি (saveExpense)
    // -------------------------------------------------------------
    if (action === "saveExpense") {
      const exp = postData.expense;
      const expSheet = ss.getSheetByName("Expenses");
      expSheet.appendRow([
        exp.id || "exp_" + Date.now(),
        exp.date || todayStr(),
        exp.time || "",
        exp.name || exp.category || "খরচ",
        exp.category || "অন্যান্য",
        exp.amount || 0,
        exp.cashAccount || "cash",
        exp.type || "business",
        exp.note || ""
      ]);
      return jsonResponse({ status: "success" });
    }

    return jsonResponse({ status: "error", message: "Action not recognized" });

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
// 🛠️ ইউটিলিটি ফাংশনস
// ============================================================
function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function todayStr() {
  return Utilities.formatDate(new Date(), "Asia/Dhaka", "yyyy-MM-dd");
}

function formatDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Dhaka", "yyyy-MM-dd");
  }
  return String(val).slice(0, 10);
}