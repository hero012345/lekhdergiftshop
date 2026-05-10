// ========================================
// كود Google Apps Script - مع الخزنة والمعاملات
// ========================================

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'error',
          message: 'لا توجد بيانات مُرسلة'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    
    Logger.log('=== بداية المزامنة ===');
    Logger.log('Action: ' + data.action);
    Logger.log('Device ID: ' + (data.deviceId || 'unknown'));
    
    if (data.action === 'sync') {
      // NEW: Incremental sync with device-specific tracking
      var deviceId = data.deviceId || 'default_' + new Date().toISOString().slice(0,10);
      var syncTime = data.timestamp || new Date().toISOString();
      
      // Get or create device-specific sheet for tracking
      var deviceSheetName = 'جهاز_' + deviceId.replace(/[^A-Za-z0-9_]/g, '_').slice(0, 20);
      var deviceSheet = sheet.getSheetByName(deviceSheetName);
      if (!deviceSheet) {
        deviceSheet = sheet.insertSheet(deviceSheetName);
        deviceSheet.getRange(1, 1, 1, 3).setBackground('#6b7280').setFontWeight('bold').setFontColor('#ffffff');
        deviceSheet.appendRow(['التوقيت', 'النوع', 'البيانات']);
      }
      
      // Log this device's sync attempt
      deviceSheet.appendRow([syncTime, 'full_sync', JSON.stringify({
        productsCount: (data.products || []).length,
        salesCount: (data.sales || []).length,
        cashRegister: data.cashRegister || 0
      })]);
      
      // ===== CENTRAL MERGE LOGIC =====
      // Instead of overwriting, we merge data intelligently
      
      // ===== CENTRAL MERGE LOGIC =====
      // Instead of overwriting, we merge data intelligently
      
      // 1. PRODUCTS: Merge by ID, update existing or add new (skip empty products)
      var products = data.products || [];
      var productsSheet = sheet.getSheetByName('المنتجات') || sheet.insertSheet('المنتجات');
      var existingProducts = getSheetAsObjects(productsSheet);
      var productMap = {};
      for (var i = 0; i < existingProducts.length; i++) {
        // Skip empty products (no name or no id)
        if (existingProducts[i].id && existingProducts[i].name && String(existingProducts[i].name).trim() !== '') {
          productMap[existingProducts[i].id] = existingProducts[i];
        }
      }

      // Update with incoming products (skip empty ones)
      for (var i = 0; i < products.length; i++) {
        var p = products[i];
        // Skip products with empty names or invalid IDs
        if (!p.id || !p.name || String(p.name).trim() === '') {
          Logger.log('⚠️ تخطي منتج فارغ أو بدون اسم: ' + JSON.stringify(p));
          continue;
        }
        if (productMap[p.id]) {
          // Update existing - use latest timestamp or incoming data
          productMap[p.id] = Object.assign(productMap[p.id], p);
        } else {
          // Add new product
          productMap[p.id] = p;
        }
      }

      // Rewrite products sheet with merged data
      productsSheet.clear();
      productsSheet.getRange(1, 1, 1, 9).setBackground('#f59e0b').setFontWeight('bold').setFontColor('#ffffff');
      productsSheet.appendRow(['ID', 'الاسم', 'الباركود', 'التكلفة', 'التجزئة', 'الجملة', 'المخزون', 'الفئة', 'الوحدة']);
      for (var id in productMap) {
        var p = productMap[id];
        productsSheet.appendRow([
          p.id || '',
          p.name || '',
          p.barcode || '',
          p.cost || 0,
          p.retail || 0,
          p.wholesale || 0,
          p.stock || 0,
          p.category || '',
          p.unit || 'piece'
        ]);
      }
      Logger.log('✓ تم دمج ' + Object.keys(productMap).length + ' منتج');
      
      // 2. SALES: Append only (never overwrite sales data)
      var sales = data.sales || [];
      var salesSheet = sheet.getSheetByName('المبيعات') || sheet.insertSheet('المبيعات');
      if (salesSheet.getLastRow() === 0) {
        salesSheet.getRange(1, 1, 1, 6).setBackground('#10b981').setFontWeight('bold').setFontColor('#ffffff');
        salesSheet.appendRow(['رقم الفاتورة', 'التاريخ', 'الإجمالي', 'الكاشير', 'البائع', 'عدد المنتجات']);
      }
      var existingSales = getSheetAsObjects(salesSheet, ['id', 'date', 'total', 'cashier', 'sellerId', 'itemsCount']);
      var existingSalesIds = {};
      for (var i = 0; i < existingSales.length; i++) {
        existingSalesIds[existingSales[i].id] = true;
      }
      
      // Only add sales that don't exist yet
      var newSalesCount = 0;
      for (var i = 0; i < sales.length; i++) {
        var s = sales[i];
        if (!existingSalesIds[s.id]) {
          salesSheet.appendRow([
            s.id || '',
            s.date || '',
            s.total || 0,
            s.cashier || '',
            s.sellerId || '',
            s.items ? s.items.length : 0
          ]);
          newSalesCount++;
        }
      }
      Logger.log('✓ تمت إضافة ' + newSalesCount + ' عملية بيع جديدة');
      
      // 3. DEBTS: Merge by ID
      var debts = data.debts || [];
      var debtsSheet = sheet.getSheetByName('الديون') || sheet.insertSheet('الديون');
      var existingDebts = getSheetAsObjects(debtsSheet);
      var debtMap = {};
      for (var i = 0; i < existingDebts.length; i++) {
        debtMap[existingDebts[i].id] = existingDebts[i];
      }
      for (var i = 0; i < debts.length; i++) {
        var d = debts[i];
        if (debtMap[d.id]) {
          // Keep the one with higher paid amount (more recent payment)
          if ((d.paid || 0) > (debtMap[d.id].paid || 0)) {
            debtMap[d.id] = d;
          }
        } else {
          debtMap[d.id] = d;
        }
      }
      debtsSheet.clear();
      debtsSheet.getRange(1, 1, 1, 8).setBackground('#ef4444').setFontWeight('bold').setFontColor('#ffffff');
      debtsSheet.appendRow(['ID', 'العميل', 'الهاتف', 'المبلغ الكلي', 'المدفوع', 'المتبقي', 'التاريخ', 'العناصر']);
      for (var id in debtMap) {
        var d = debtMap[id];
        debtsSheet.appendRow([
          d.id || '',
          d.customer || '',
          d.phone || '',
          d.amount || 0,
          d.paid || 0,
          d.remaining || 0,
          d.date || '',
          d.items || ''
        ]);
      }
      Logger.log('✓ تم دمج ' + Object.keys(debtMap).length + ' دين');
      
      // 4. SUPPLIERS: Merge by ID
      var suppliers = data.suppliers || [];
      var suppliersSheet = sheet.getSheetByName('الموردين') || sheet.insertSheet('الموردين');
      var existingSuppliers = getSheetAsObjects(suppliersSheet);
      var supplierMap = {};
      for (var i = 0; i < existingSuppliers.length; i++) {
        supplierMap[existingSuppliers[i].id] = existingSuppliers[i];
      }
      for (var i = 0; i < suppliers.length; i++) {
        var s = suppliers[i];
        if (supplierMap[s.id]) {
          // Use higher totalPurchases value
          if ((s.totalPurchases || 0) > (supplierMap[s.id].totalPurchases || 0)) {
            supplierMap[s.id] = s;
          }
        } else {
          supplierMap[s.id] = s;
        }
      }
      suppliersSheet.clear();
      suppliersSheet.getRange(1, 1, 1, 5).setBackground('#3b82f6').setFontWeight('bold').setFontColor('#ffffff');
      suppliersSheet.appendRow(['ID', 'الاسم', 'الهاتف', 'العنوان', 'إجمالي المشتريات']);
      for (var id in supplierMap) {
        var s = supplierMap[id];
        suppliersSheet.appendRow([
          s.id || '',
          s.name || '',
          s.phone || '',
          s.address || '',
          s.totalPurchases || 0
        ]);
      }
      Logger.log('✓ تم دمج ' + Object.keys(supplierMap).length + ' مورد');
      
      // 5. WORKERS: Merge by ID
      var workers = data.workers || [];
      var workersSheet = sheet.getSheetByName('الموظفين') || sheet.insertSheet('الموظفين');
      var existingWorkers = getSheetAsObjects(workersSheet);
      var workerMap = {};
      for (var i = 0; i < existingWorkers.length; i++) {
        workerMap[existingWorkers[i].id] = existingWorkers[i];
      }
      for (var i = 0; i < workers.length; i++) {
        var w = workers[i];
        if (workerMap[w.id]) {
          workerMap[w.id] = Object.assign(workerMap[w.id], w);
        } else {
          workerMap[w.id] = w;
        }
      }
      workersSheet.clear();
      workersSheet.getRange(1, 1, 1, 6).setBackground('#8b5cf6').setFontWeight('bold').setFontColor('#ffffff');
      workersSheet.appendRow(['ID', 'الاسم', 'الهاتف', 'الكود', 'الدور', 'آخر نشاط']);
      for (var id in workerMap) {
        var w = workerMap[id];
        var workerPin = w.pin ? "'" + w.pin.toString() : '';
        workersSheet.appendRow([
          w.id || '',
          w.name || '',
          w.phone || '',
          workerPin,
          w.role || 'عامل',
          w.lastActivity || ''
        ]);
      }
      Logger.log('✓ تم دمج ' + Object.keys(workerMap).length + ' موظف');
      
      // 6. PURCHASES: Append only
      var purchases = data.purchases || [];
      var purchasesSheet = sheet.getSheetByName('المشتريات') || sheet.insertSheet('المشتريات');
      if (purchasesSheet.getLastRow() === 0) {
        purchasesSheet.getRange(1, 1, 1, 6).setBackground('#06b6d4').setFontWeight('bold').setFontColor('#ffffff');
        purchasesSheet.appendRow(['ID', 'اسم المورد', 'المبلغ', 'ملاحظات', 'التاريخ', 'معرف المورد']);
      }
      var existingPurchases = getSheetAsObjects(purchasesSheet);
      var existingPurchaseIds = {};
      for (var i = 0; i < existingPurchases.length; i++) {
        existingPurchaseIds[existingPurchases[i].id] = true;
      }
      var newPurchasesCount = 0;
      for (var i = 0; i < purchases.length; i++) {
        var p = purchases[i];
        if (!existingPurchaseIds[p.id]) {
          purchasesSheet.appendRow([
            p.id || '',
            p.supplierName || '',
            p.amount || 0,
            p.notes || '',
            p.date || '',
            p.supplierId || ''
          ]);
          newPurchasesCount++;
        }
      }
      Logger.log('✓ تمت إضافة ' + newPurchasesCount + ' عملية شراء جديدة');
      
      // 7. TRANSACTIONS: Append only
      var transactions = data.transactions || [];
      var transactionsSheet = sheet.getSheetByName('المعاملات المالية') || sheet.insertSheet('المعاملات المالية');
      if (transactionsSheet.getLastRow() === 0) {
        transactionsSheet.getRange(1, 1, 1, 7).setBackground('#ec4899').setFontWeight('bold').setFontColor('#ffffff');
        transactionsSheet.appendRow(['ID', 'التاريخ', 'النوع', 'الوصف', 'المبلغ', 'الرصيد بعد العملية', 'المستخدم']);
      }
      var existingTransactions = getSheetAsObjects(transactionsSheet);
      var existingTxIds = {};
      for (var i = 0; i < existingTransactions.length; i++) {
        existingTxIds[existingTransactions[i].id] = true;
      }
      var newTxCount = 0;
      for (var i = 0; i < transactions.length; i++) {
        var t = transactions[i];
        if (!existingTxIds[t.id]) {
          transactionsSheet.appendRow([
            t.id || '',
            t.date || '',
            t.type || '',
            t.description || '',
            t.amount || 0,
            t.balanceAfter || 0,
            t.user || ''
          ]);
          newTxCount++;
        }
      }
      Logger.log('✓ تمت إضافة ' + newTxCount + ' معاملة مالية جديدة');
      
      // 8. CASH REGISTER: Use the highest value (most recent/largest)
      var infoSheet = sheet.getSheetByName('معلومات المتجر') || sheet.insertSheet('معلومات المتجر');
      var existingInfo = getSheetAsObjects(infoSheet);
      var infoMap = {};
      for (var i = 0; i < existingInfo.length; i++) {
        if (existingInfo[i][0]) infoMap[existingInfo[i][0]] = existingInfo[i][1];
      }
      
      // Update with incoming info
      if (data.shopName) infoMap['اسم المتجر'] = data.shopName;
      if (data.manager && data.manager.name) infoMap['اسم المدير'] = data.manager.name;
      if (data.manager && data.manager.pin) infoMap['كود المدير'] = "'" + data.manager.pin.toString();
      if (data.manager && data.manager.phone) infoMap['هاتف المدير'] = data.manager.phone;
      
      // For cash register, use the maximum value (prevents losing money data)
      var currentCash = infoMap['رصيد الخزنة'] || 0;
      var incomingCash = data.cashRegister || 0;
      infoMap['رصيد الخزنة'] = Math.max(currentCash, incomingCash);
      
      infoMap['آخر مزامنة'] = new Date().toISOString();
      if (data.nextInvoice) {
        var currentNextInv = infoMap['رقم الفاتورة القادم'] || 1001;
        infoMap['رقم الفاتورة القادم'] = Math.max(currentNextInv, data.nextInvoice);
      }
      
      infoSheet.clear();
      infoSheet.getRange(1, 1, 1, 2).setBackground('#f59e0b').setFontWeight('bold').setFontColor('#ffffff');
      infoSheet.appendRow(['المفتاح', 'القيمة']);
      for (var key in infoMap) {
        infoSheet.appendRow([key, infoMap[key]]);
      }
      Logger.log('✓ تم تحديث معلومات المتجر والخزنة');
      
      // ===== 9. سجل المزامنة =====
      var syncLogSheet = sheet.getSheetByName('سجل المزامنة') || sheet.insertSheet('سجل المزامنة');
      if (syncLogSheet.getLastRow() === 0) {
        syncLogSheet.getRange(1, 1, 1, 9).setBackground('#6b7280').setFontWeight('bold').setFontColor('#ffffff');
        syncLogSheet.appendRow(['التوقيت', 'الجهاز', 'المتجر', 'المنتجات', 'المبيعات الجديدة', 'الموظفين', 'الموردين', 'المعاملات الجديدة', 'الحالة']);
      }
      syncLogSheet.appendRow([
        syncTime,
        deviceId,
        data.shopName || 'Unknown',
        products.length,
        newSalesCount,
        workers.length,
        suppliers.length,
        newTxCount,
        'SUCCESS'
      ]);
      
      Logger.log('=== انتهت المزامنة بنجاح ===');
      
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'success',
          message: 'تمت المزامنة بنجاح ✅',
          timestamp: new Date().toISOString(),
          merged: {
            products: Object.keys(productMap).length,
            salesAdded: newSalesCount,
            debts: Object.keys(debtMap).length,
            workers: Object.keys(workerMap).length,
            suppliers: Object.keys(supplierMap).length,
            purchasesAdded: newPurchasesCount,
            transactionsAdded: newTxCount,
            cashRegister: infoMap['رصيد الخزنة']
          }
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    if (data.action === 'get') {
      return handleGetRequest();
    }
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'Action غير معروف'
      }))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    Logger.log('❌ خطأ: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'خطأ في السيرفر: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Helper function to read sheet as array of objects
function getSheetAsObjects(sheet, headers) {
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var data = sheet.getDataRange().getValues();
  if (!headers) {
    // Auto-detect headers from first row
    headers = data[0];
  }
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var obj = {};
    for (var j = 0; j < headers.length; j++) {
      obj[headers[j]] = data[i][j];
    }
    result.push(obj);
  }
  return result;
}

function doGet(e) {
  return handleGetRequest();
}

function handleGetRequest() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet();
    
    Logger.log('=== بداية قراءة البيانات ===');
    
    // ===== معلومات المتجر =====
    var infoSheet = sheet.getSheetByName('معلومات المتجر');
    var shopData = {
      shopName: '',
      managerName: '',
      managerPin: '',
      managerPhone: '',
      cashRegister: 0,
      nextInvoice: 1001
    };
    
    if (infoSheet && infoSheet.getLastRow() > 1) {
      var infoData = infoSheet.getDataRange().getValues();
      for (var i = 1; i < infoData.length; i++) {
        if (infoData[i][0] === 'اسم المتجر') shopData.shopName = infoData[i][1];
        if (infoData[i][0] === 'اسم المدير') shopData.managerName = infoData[i][1];
        
        if (infoData[i][0] === 'كود المدير') {
          var rawPin = infoData[i][1];
          if (rawPin) {
            shopData.managerPin = rawPin.toString().replace(/^'/, '');
          }
        }
        
        if (infoData[i][0] === 'هاتف المدير') shopData.managerPhone = infoData[i][1];
        if (infoData[i][0] === 'رصيد الخزنة') shopData.cashRegister = infoData[i][1] || 0;
        if (infoData[i][0] === 'رقم الفاتورة القادم') shopData.nextInvoice = infoData[i][1];
      }
    }
    
    // ===== المنتجات =====
    var productsSheet = sheet.getSheetByName('المنتجات');
    var products = [];
    if (productsSheet && productsSheet.getLastRow() > 1) {
      var prodData = productsSheet.getDataRange().getValues();
      for (var i = 1; i < prodData.length; i++) {
        products.push({
          id: prodData[i][0],
          name: prodData[i][1],
          barcode: prodData[i][2],
          cost: prodData[i][3],
          retail: prodData[i][4],
          wholesale: prodData[i][5],
          stock: prodData[i][6],
          category: prodData[i][7],
          unit: prodData[i][8]
        });
      }
    }
    
    // ===== المبيعات =====
    var salesSheet = sheet.getSheetByName('المبيعات');
    var sales = [];
    if (salesSheet && salesSheet.getLastRow() > 1) {
      var salesData = salesSheet.getDataRange().getValues();
      for (var i = 1; i < salesData.length; i++) {
        sales.push({
          id: salesData[i][0],
          date: salesData[i][1],
          total: salesData[i][2],
          cashier: salesData[i][3],
          sellerId: salesData[i][4],
          items: []
        });
      }
    }
    
    // ===== الديون =====
    var debtsSheet = sheet.getSheetByName('الديون');
    var debts = [];
    if (debtsSheet && debtsSheet.getLastRow() > 1) {
      var debtsData = debtsSheet.getDataRange().getValues();
      for (var i = 1; i < debtsData.length; i++) {
        debts.push({
          id: debtsData[i][0],
          customer: debtsData[i][1],
          phone: debtsData[i][2],
          amount: debtsData[i][3],
          paid: debtsData[i][4],
          remaining: debtsData[i][5],
          date: debtsData[i][6],
          items: debtsData[i][7]
        });
      }
    }
    
    // ===== الموردين =====
    var suppliersSheet = sheet.getSheetByName('الموردين');
    var suppliers = [];
    if (suppliersSheet && suppliersSheet.getLastRow() > 1) {
      var suppData = suppliersSheet.getDataRange().getValues();
      for (var i = 1; i < suppData.length; i++) {
        suppliers.push({
          id: suppData[i][0],
          name: suppData[i][1],
          phone: suppData[i][2],
          address: suppData[i][3],
          totalPurchases: suppData[i][4]
        });
      }
    }
    
    // ===== الموظفين =====
    var workersSheet = sheet.getSheetByName('الموظفين');
    var workers = [];
    if (workersSheet && workersSheet.getLastRow() > 1) {
      var workData = workersSheet.getDataRange().getValues();
      for (var i = 1; i < workData.length; i++) {
        if (workData[i][0] && workData[i][1]) {
          var workerPin = workData[i][3];
          if (workerPin) {
            workerPin = workerPin.toString().replace(/^'/, '');
          }
          
          workers.push({
            id: workData[i][0],
            name: workData[i][1],
            phone: workData[i][2],
            pin: workerPin,
            role: workData[i][4],
            lastActivity: workData[i][5] || ''
          });
        }
      }
    }
    
    // ===== المشتريات =====
    var purchasesSheet = sheet.getSheetByName('المشتريات');
    var purchases = [];
    if (purchasesSheet && purchasesSheet.getLastRow() > 1) {
      var purchData = purchasesSheet.getDataRange().getValues();
      for (var i = 1; i < purchData.length; i++) {
        purchases.push({
          id: purchData[i][0],
          supplierId: purchData[i][5],
          supplierName: purchData[i][1],
          amount: purchData[i][2],
          notes: purchData[i][3],
          date: purchData[i][4]
        });
      }
    }
    
    // ===== المعاملات المالية (جديد) =====
    var transactionsSheet = sheet.getSheetByName('المعاملات المالية');
    var transactions = [];
    if (transactionsSheet && transactionsSheet.getLastRow() > 1) {
      var transData = transactionsSheet.getDataRange().getValues();
      for (var i = 1; i < transData.length; i++) {
        transactions.push({
          id: transData[i][0],
          date: transData[i][1],
          type: transData[i][2],
          description: transData[i][3],
          amount: transData[i][4],
          balanceAfter: transData[i][5],
          user: transData[i][6]
        });
      }
    }
    
    Logger.log('=== انتهت القراءة ===');
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        hasData: shopData.shopName ? true : false,
        shopName: shopData.shopName || '',
        manager: {
          name: shopData.managerName || '',
          pin: shopData.managerPin || '',
          phone: shopData.managerPhone || ''
        },
        cashRegister: shopData.cashRegister || 0,
        nextInvoice: shopData.nextInvoice || 1001,
        products: products,
        sales: sales,
        debts: debts,
        suppliers: suppliers,
        workers: workers,
        purchases: purchases,
        transactions: transactions,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('❌ خطأ في القراءة: ' + error.toString());
    Logger.log('Stack: ' + error.stack);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: 'خطأ في قراءة البيانات: ' + error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
