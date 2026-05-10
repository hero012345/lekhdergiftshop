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
    
    if (data.action === 'sync') {
      var products = data.products || [];
      var sales = data.sales || [];
      var debts = data.debts || [];
      var suppliers = data.suppliers || [];
      var workers = data.workers || [];
      var purchases = data.purchases || [];
      var transactions = data.transactions || [];
      
      Logger.log('عدد المنتجات: ' + products.length);
      Logger.log('عدد المبيعات: ' + sales.length);
      Logger.log('عدد الموظفين: ' + workers.length);
      Logger.log('عدد الموردين: ' + suppliers.length);
      Logger.log('عدد المعاملات: ' + transactions.length);
      
      // ===== 1. المنتجات =====
      var productsSheet = sheet.getSheetByName('المنتجات') || sheet.insertSheet('المنتجات');
      productsSheet.clear();
      productsSheet.getRange(1, 1, 1, 9).setBackground('#f59e0b').setFontWeight('bold').setFontColor('#ffffff');
      productsSheet.appendRow(['ID', 'الاسم', 'الباركود', 'التكلفة', 'التجزئة', 'الجملة', 'المخزون', 'الفئة', 'الوحدة']);
      
      if (products.length > 0) {
        for (var i = 0; i < products.length; i++) {
          var p = products[i];
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
      }
      Logger.log('✓ تم حفظ ' + products.length + ' منتج');
      
      // ===== 2. المبيعات =====
      var salesSheet = sheet.getSheetByName('المبيعات') || sheet.insertSheet('المبيعات');
      salesSheet.clear();
      salesSheet.getRange(1, 1, 1, 6).setBackground('#10b981').setFontWeight('bold').setFontColor('#ffffff');
      salesSheet.appendRow(['رقم الفاتورة', 'التاريخ', 'الإجمالي', 'الكاشير', 'البائع', 'عدد المنتجات']);
      
      if (sales.length > 0) {
        for (var i = 0; i < sales.length; i++) {
          var s = sales[i];
          salesSheet.appendRow([
            s.id || '',
            s.date || '',
            s.total || 0,
            s.cashier || '',
            s.sellerId || '',
            s.items ? s.items.length : 0
          ]);
        }
      }
      Logger.log('✓ تم حفظ ' + sales.length + ' عملية بيع');
      
      // ===== 3. الديون =====
      var debtsSheet = sheet.getSheetByName('الديون') || sheet.insertSheet('الديون');
      debtsSheet.clear();
      debtsSheet.getRange(1, 1, 1, 8).setBackground('#ef4444').setFontWeight('bold').setFontColor('#ffffff');
      debtsSheet.appendRow(['ID', 'العميل', 'الهاتف', 'المبلغ الكلي', 'المدفوع', 'المتبقي', 'التاريخ', 'العناصر']);
      
      if (debts.length > 0) {
        for (var i = 0; i < debts.length; i++) {
          var d = debts[i];
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
      }
      Logger.log('✓ تم حفظ ' + debts.length + ' دين');
      
      // ===== 4. الموردين =====
      var suppliersSheet = sheet.getSheetByName('الموردين') || sheet.insertSheet('الموردين');
      suppliersSheet.clear();
      suppliersSheet.getRange(1, 1, 1, 5).setBackground('#3b82f6').setFontWeight('bold').setFontColor('#ffffff');
      suppliersSheet.appendRow(['ID', 'الاسم', 'الهاتف', 'العنوان', 'إجمالي المشتريات']);
      
      if (suppliers.length > 0) {
        for (var i = 0; i < suppliers.length; i++) {
          var s = suppliers[i];
          suppliersSheet.appendRow([
            s.id || '',
            s.name || '',
            s.phone || '',
            s.address || '',
            s.totalPurchases || 0
          ]);
        }
      }
      Logger.log('✓ تم حفظ ' + suppliers.length + ' مورد');
      
      // ===== 5. الموظفين =====
      var workersSheet = sheet.getSheetByName('الموظفين') || sheet.insertSheet('الموظفين');
      workersSheet.clear();
      workersSheet.getRange(1, 1, 1, 6).setBackground('#8b5cf6').setFontWeight('bold').setFontColor('#ffffff');
      workersSheet.appendRow(['ID', 'الاسم', 'الهاتف', 'الكود', 'الدور', 'آخر نشاط']);
      
      if (workers.length > 0) {
        for (var i = 0; i < workers.length; i++) {
          var w = workers[i];
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
        Logger.log('✓ تم حفظ ' + workers.length + ' موظف');
      } else {
        Logger.log('⚠️ لا يوجد موظفين');
      }
      
      // ===== 6. المشتريات =====
      var purchasesSheet = sheet.getSheetByName('المشتريات') || sheet.insertSheet('المشتريات');
      purchasesSheet.clear();
      purchasesSheet.getRange(1, 1, 1, 6).setBackground('#06b6d4').setFontWeight('bold').setFontColor('#ffffff');
      purchasesSheet.appendRow(['ID', 'اسم المورد', 'المبلغ', 'ملاحظات', 'التاريخ', 'معرف المورد']);
      
      if (purchases.length > 0) {
        for (var i = 0; i < purchases.length; i++) {
          var p = purchases[i];
          purchasesSheet.appendRow([
            p.id || '',
            p.supplierName || '',
            p.amount || 0,
            p.notes || '',
            p.date || '',
            p.supplierId || ''
          ]);
        }
      }
      Logger.log('✓ تم حفظ ' + purchases.length + ' عملية شراء');
      
      // ===== 7. المعاملات المالية (جديد) =====
      var transactionsSheet = sheet.getSheetByName('المعاملات المالية') || sheet.insertSheet('المعاملات المالية');
      transactionsSheet.clear();
      transactionsSheet.getRange(1, 1, 1, 7).setBackground('#ec4899').setFontWeight('bold').setFontColor('#ffffff');
      transactionsSheet.appendRow(['ID', 'التاريخ', 'النوع', 'الوصف', 'المبلغ', 'الرصيد بعد العملية', 'المستخدم']);
      
      if (transactions.length > 0) {
        for (var i = 0; i < transactions.length; i++) {
          var t = transactions[i];
          transactionsSheet.appendRow([
            t.id || '',
            t.date || '',
            t.type || '',
            t.description || '',
            t.amount || 0,
            t.balanceAfter || 0,
            t.user || ''
          ]);
        }
      }
      Logger.log('✓ تم حفظ ' + transactions.length + ' معاملة مالية');
      
      // ===== 8. معلومات المتجر والخزنة =====
      var infoSheet = sheet.getSheetByName('معلومات المتجر') || sheet.insertSheet('معلومات المتجر');
      infoSheet.clear();
      infoSheet.getRange(1, 1, 1, 2).setBackground('#f59e0b').setFontWeight('bold').setFontColor('#ffffff');
      infoSheet.appendRow(['المفتاح', 'القيمة']);
      infoSheet.appendRow(['اسم المتجر', data.shopName || '']);
      infoSheet.appendRow(['اسم المدير', data.manager && data.manager.name ? data.manager.name : '']);
      
      var managerPin = data.manager && data.manager.pin ? "'" + data.manager.pin.toString() : '';
      infoSheet.appendRow(['كود المدير', managerPin]);
      
      infoSheet.appendRow(['هاتف المدير', data.manager && data.manager.phone ? data.manager.phone : '']);
      infoSheet.appendRow(['رصيد الخزنة', data.cashRegister || 0]);
      infoSheet.appendRow(['آخر مزامنة', data.timestamp || new Date().toISOString()]);
      infoSheet.appendRow(['رقم الفاتورة القادم', data.nextInvoice || 1001]);
      
      Logger.log('✓ تم حفظ رصيد الخزنة: ' + (data.cashRegister || 0));
      
      // ===== 9. سجل المزامنة =====
      var syncLogSheet = sheet.getSheetByName('سجل المزامنة') || sheet.insertSheet('سجل المزامنة');
      if (syncLogSheet.getLastRow() === 0) {
        syncLogSheet.getRange(1, 1, 1, 8).setBackground('#6b7280').setFontWeight('bold').setFontColor('#ffffff');
        syncLogSheet.appendRow(['التاريخ', 'المتجر', 'المنتجات', 'المبيعات', 'الموظفين', 'الموردين', 'المعاملات', 'الحالة']);
      }
      syncLogSheet.appendRow([
        new Date().toISOString(),
        data.shopName || 'Unknown',
        products.length,
        sales.length,
        workers.length,
        suppliers.length,
        transactions.length,
        'SUCCESS'
      ]);
      
      Logger.log('=== انتهت المزامنة بنجاح ===');
      
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'success',
          message: 'تمت المزامنة بنجاح ✅',
          timestamp: new Date().toISOString(),
          counts: {
            products: products.length,
            sales: sales.length,
            workers: workers.length,
            suppliers: suppliers.length,
            purchases: purchases.length,
            transactions: transactions.length
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
