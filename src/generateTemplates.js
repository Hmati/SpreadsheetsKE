const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Create Business Bundle (Comprehensive business management template)
async function createBusinessBundle() {
  const workbook = new ExcelJS.Workbook();

  // ===== DASHBOARD SHEET =====
  const dashboardSheet = workbook.addWorksheet('Dashboard');
  dashboardSheet.getCell('A1').value = 'BUSINESS MANAGEMENT DASHBOARD';
  dashboardSheet.getCell('A1').font = { size: 16, bold: true, color: { argb: 'FFC8601A' } };
  dashboardSheet.getCell('A1').alignment = { horizontal: 'center' };
  dashboardSheet.mergeCells('A1:E1');

  // Key metrics section
  dashboardSheet.getCell('A3').value = 'KEY METRICS';
  dashboardSheet.getCell('A3').font = { bold: true };
  dashboardSheet.getCell('A4').value = 'Total Revenue:';
  dashboardSheet.getCell('B4').value = '=SUM(Sales!D:D)';
  dashboardSheet.getCell('A5').value = 'Total Expenses:';
  dashboardSheet.getCell('B5').value = '=SUM(Expenses!D:D)';
  dashboardSheet.getCell('A6').value = 'Net Profit:';
  dashboardSheet.getCell('B6').value = '=B4-B5';
  dashboardSheet.getCell('A7').value = 'Profit Margin:';
  dashboardSheet.getCell('B7').value = '=IF(B4>0,B6/B4,"")';
  dashboardSheet.getCell('B7').numFmt = '0.00%';

  // Monthly overview
  dashboardSheet.getCell('A9').value = 'MONTHLY OVERVIEW';
  dashboardSheet.getCell('A9').font = { bold: true };
  dashboardSheet.getCell('A10').value = 'Month';
  dashboardSheet.getCell('B10').value = 'Revenue';
  dashboardSheet.getCell('C10').value = 'Expenses';
  dashboardSheet.getCell('D10').value = 'Profit';

  // Sample monthly data
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
  months.forEach((month, index) => {
    dashboardSheet.getCell(`A${11+index}`).value = month;
    dashboardSheet.getCell(`B${11+index}`).value = Math.floor(Math.random() * 50000) + 10000;
    dashboardSheet.getCell(`C${11+index}`).value = Math.floor(Math.random() * 30000) + 5000;
    dashboardSheet.getCell(`D${11+index}`).value = `=B${11+index}-C${11+index}`;
  });

  // ===== SALES SHEET =====
  const salesSheet = workbook.addWorksheet('Sales');
  salesSheet.columns = [
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Customer', key: 'customer', width: 20 },
    { header: 'Product/Service', key: 'product', width: 25 },
    { header: 'Quantity', key: 'quantity', width: 10 },
    { header: 'Unit Price', key: 'unitPrice', width: 12 },
    { header: 'Total Amount', key: 'total', width: 15 },
    { header: 'Payment Method', key: 'paymentMethod', width: 15 },
    { header: 'Status', key: 'status', width: 10 }
  ];

  // Sample sales data
  const salesData = [
    { date: '2024-01-15', customer: 'ABC Ltd', product: 'Consulting Services', quantity: 1, unitPrice: 25000, paymentMethod: 'M-Pesa', status: 'Paid' },
    { date: '2024-01-20', customer: 'XYZ Corp', product: 'Software License', quantity: 5, unitPrice: 15000, paymentMethod: 'Bank Transfer', status: 'Paid' },
    { date: '2024-02-01', customer: 'Tech Solutions', product: 'Training Workshop', quantity: 1, unitPrice: 45000, paymentMethod: 'M-Pesa', status: 'Pending' },
    { date: '2024-02-10', customer: 'Global Traders', product: 'Import Services', quantity: 3, unitPrice: 30000, paymentMethod: 'Cheque', status: 'Paid' },
    { date: '2024-02-15', customer: 'Local Enterprises', product: 'Business Plan', quantity: 1, unitPrice: 35000, paymentMethod: 'Cash', status: 'Paid' }
  ];

  salesData.forEach((sale, index) => {
    sale.total = sale.quantity * sale.unitPrice;
    salesSheet.addRow(sale);
  });

  // Add totals row
  salesSheet.addRow({});
  salesSheet.addRow({
    date: 'TOTALS',
    total: '=SUM(F2:F6)',
    status: ''
  });
  salesSheet.getCell('A8').font = { bold: true };
  salesSheet.getCell('F8').font = { bold: true };

  // ===== EXPENSES SHEET =====
  const expensesSheet = workbook.addWorksheet('Expenses');
  expensesSheet.columns = [
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Category', key: 'category', width: 20 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Amount', key: 'amount', width: 15 },
    { header: 'Payment Method', key: 'paymentMethod', width: 15 },
    { header: 'Tax Deductible', key: 'taxDeductible', width: 15 }
  ];

  const expenseData = [
    { date: '2024-01-05', category: 'Office Supplies', description: 'Stationery and printing materials', amount: 15000, paymentMethod: 'Cash', taxDeductible: 'Yes' },
    { date: '2024-01-10', category: 'Rent', description: 'Monthly office rent', amount: 80000, paymentMethod: 'Bank Transfer', taxDeductible: 'Yes' },
    { date: '2024-01-15', category: 'Utilities', description: 'Electricity and water bills', amount: 25000, paymentMethod: 'M-Pesa', taxDeductible: 'Yes' },
    { date: '2024-01-20', category: 'Marketing', description: 'Social media advertising', amount: 30000, paymentMethod: 'Card', taxDeductible: 'Yes' },
    { date: '2024-01-25', category: 'Salaries', description: 'Employee salaries', amount: 150000, paymentMethod: 'Bank Transfer', taxDeductible: 'No' },
    { date: '2024-02-01', category: 'Travel', description: 'Business trip to Nairobi', amount: 45000, paymentMethod: 'Cash', taxDeductible: 'Yes' }
  ];

  expenseData.forEach(expense => expensesSheet.addRow(expense));
  expensesSheet.addRow({});
  expensesSheet.addRow({
    date: 'TOTALS',
    amount: '=SUM(D2:D7)',
    description: ''
  });
  expensesSheet.getCell('A9').font = { bold: true };
  expensesSheet.getCell('D9').font = { bold: true };

  // ===== INVENTORY SHEET =====
  const inventorySheet = workbook.addWorksheet('Inventory');
  inventorySheet.columns = [
    { header: 'Item Code', key: 'itemCode', width: 12 },
    { header: 'Item Name', key: 'itemName', width: 25 },
    { header: 'Category', key: 'category', width: 15 },
    { header: 'Current Stock', key: 'currentStock', width: 12 },
    { header: 'Minimum Stock', key: 'minStock', width: 12 },
    { header: 'Unit Cost', key: 'unitCost', width: 12 },
    { header: 'Total Value', key: 'totalValue', width: 15 },
    { header: 'Supplier', key: 'supplier', width: 20 },
    { header: 'Last Restock', key: 'lastRestock', width: 12 }
  ];

  const inventoryData = [
    { itemCode: 'ITM001', itemName: 'Office Chairs', category: 'Furniture', currentStock: 25, minStock: 5, unitCost: 8500, supplier: 'Furniture Plus', lastRestock: '2024-01-10' },
    { itemCode: 'ITM002', itemName: 'A4 Paper', category: 'Supplies', currentStock: 500, minStock: 100, unitCost: 50, supplier: 'Paper World', lastRestock: '2024-01-15' },
    { itemCode: 'ITM003', itemName: 'Printer Ink', category: 'Supplies', currentStock: 20, minStock: 5, unitCost: 3500, supplier: 'Tech Supplies', lastRestock: '2024-01-20' },
    { itemCode: 'ITM004', itemName: 'Laptops', category: 'Equipment', currentStock: 8, minStock: 2, unitCost: 75000, supplier: 'Computer World', lastRestock: '2024-02-01' },
    { itemCode: 'ITM005', itemName: 'Coffee Machine', category: 'Equipment', currentStock: 1, minStock: 1, unitCost: 25000, supplier: 'Kitchen Equip', lastRestock: '2024-01-05' }
  ];

  inventoryData.forEach(item => {
    item.totalValue = item.currentStock * item.unitCost;
    inventorySheet.addRow(item);
  });

  // Add formula for total value
  for (let i = 2; i <= 6; i++) {
    inventorySheet.getCell(`G${i}`).value = `=D${i}*F${i}`;
  }

  inventorySheet.addRow({});
  inventorySheet.addRow({
    itemName: 'TOTAL INVENTORY VALUE',
    totalValue: '=SUM(G2:G6)',
    itemCode: ''
  });
  inventorySheet.getCell('B8').font = { bold: true };
  inventorySheet.getCell('G8').font = { bold: true };

  // ===== CUSTOMERS SHEET =====
  const customersSheet = workbook.addWorksheet('Customers');
  customersSheet.columns = [
    { header: 'Customer ID', key: 'customerId', width: 12 },
    { header: 'Company Name', key: 'companyName', width: 25 },
    { header: 'Contact Person', key: 'contactPerson', width: 20 },
    { header: 'Phone', key: 'phone', width: 15 },
    { header: 'Email', key: 'email', width: 25 },
    { header: 'Total Orders', key: 'totalOrders', width: 12 },
    { header: 'Total Spent', key: 'totalSpent', width: 15 },
    { header: 'Last Order Date', key: 'lastOrderDate', width: 15 }
  ];

  const customerData = [
    { customerId: 'CUST001', companyName: 'ABC Ltd', contactPerson: 'John Doe', phone: '+254712345678', email: 'john@abc.com', totalOrders: 3, totalSpent: 75000, lastOrderDate: '2024-02-01' },
    { customerId: 'CUST002', companyName: 'XYZ Corp', contactPerson: 'Jane Smith', phone: '+254723456789', email: 'jane@xyz.com', totalOrders: 2, totalSpent: 90000, lastOrderDate: '2024-01-20' },
    { customerId: 'CUST003', companyName: 'Tech Solutions', contactPerson: 'Mike Johnson', phone: '+254734567890', email: 'mike@techsol.com', totalOrders: 1, totalSpent: 45000, lastOrderDate: '2024-02-01' },
    { customerId: 'CUST004', companyName: 'Global Traders', contactPerson: 'Sarah Wilson', phone: '+254745678901', email: 'sarah@global.com', totalOrders: 4, totalSpent: 120000, lastOrderDate: '2024-02-10' },
    { customerId: 'CUST005', companyName: 'Local Enterprises', contactPerson: 'David Brown', phone: '+254756789012', email: 'david@local.com', totalOrders: 2, totalSpent: 70000, lastOrderDate: '2024-02-15' }
  ];

  customerData.forEach(customer => customersSheet.addRow(customer));

  // ===== PAYROLL SHEET =====
  const payrollSheet = workbook.addWorksheet('Payroll');
  payrollSheet.columns = [
    { header: 'Employee ID', key: 'employeeId', width: 12 },
    { header: 'Employee Name', key: 'employeeName', width: 20 },
    { header: 'Basic Salary', key: 'basicSalary', width: 15 },
    { header: 'House Allowance', key: 'houseAllowance', width: 15 },
    { header: 'Transport Allowance', key: 'transportAllowance', width: 18 },
    { header: 'Gross Salary', key: 'grossSalary', width: 15 },
    { header: 'PAYE Tax', key: 'payeTax', width: 12 },
    { header: 'NHIF', key: 'nhif', width: 10 },
    { header: 'NSSF', key: 'nssf', width: 10 },
    { header: 'Net Salary', key: 'netSalary', width: 15 }
  ];

  const payrollData = [
    { employeeId: 'EMP001', employeeName: 'Alice Johnson', basicSalary: 45000, houseAllowance: 15000, transportAllowance: 5000 },
    { employeeId: 'EMP002', employeeName: 'Bob Wilson', basicSalary: 35000, houseAllowance: 12000, transportAllowance: 4000 },
    { employeeId: 'EMP003', employeeName: 'Carol Davis', basicSalary: 55000, houseAllowance: 18000, transportAllowance: 6000 }
  ];

  payrollData.forEach(employee => {
    employee.grossSalary = employee.basicSalary + employee.houseAllowance + employee.transportAllowance;
    // Simplified tax calculations (in reality, use proper tax brackets)
    employee.payeTax = Math.round(employee.grossSalary * 0.15); // 15% PAYE
    employee.nhif = employee.grossSalary > 100000 ? 1700 : employee.grossSalary > 80000 ? 1400 : 750; // Simplified NHIF
    employee.nssf = Math.min(employee.grossSalary * 0.12, 2160); // 12% NSSF, max 2160
    employee.netSalary = employee.grossSalary - employee.payeTax - employee.nhif - employee.nssf;
    payrollSheet.addRow(employee);
  });

  // Add formulas for calculations
  for (let i = 2; i <= 4; i++) {
    payrollSheet.getCell(`F${i}`).value = `=C${i}+D${i}+E${i}`; // Gross Salary
    payrollSheet.getCell(`G${i}`).value = `=ROUND(F${i}*0.15,0)`; // PAYE (simplified)
    payrollSheet.getCell(`H${i}`).value = `=IF(F${i}>100000,1700,IF(F${i}>80000,1400,750))`; // NHIF
    payrollSheet.getCell(`I${i}`).value = `=MIN(F${i}*0.12,2160)`; // NSSF
    payrollSheet.getCell(`J${i}`).value = `=F${i}-G${i}-H${i}-I${i}`; // Net Salary
  }

  await workbook.xlsx.writeFile(path.join(__dirname, '../templates/business_bundle.xlsx'));
  console.log('Business bundle created');
}

// Create Starter Bundle (essential business sheets)
async function createStarterBundle() {
  const workbook = new ExcelJS.Workbook();

  // ===== BUSINESS OVERVIEW DASHBOARD =====
  const dashboardSheet = workbook.addWorksheet('Business Overview');
  dashboardSheet.getCell('A1').value = 'BUSINESS OVERVIEW DASHBOARD';
  dashboardSheet.getCell('A1').font = { size: 14, bold: true, color: { argb: 'FF2E75B6' } };
  dashboardSheet.getCell('A1').alignment = { horizontal: 'center' };
  dashboardSheet.mergeCells('A1:D1');

  // Key metrics
  dashboardSheet.getCell('A3').value = 'KEY METRICS';
  dashboardSheet.getCell('A3').font = { bold: true };
  dashboardSheet.getCell('A4').value = 'Monthly Revenue:';
  dashboardSheet.getCell('B4').value = '=SUM(Sales!E:E)';
  dashboardSheet.getCell('A5').value = 'Monthly Expenses:';
  dashboardSheet.getCell('B5').value = '=SUM(Expenses!D:D)';
  dashboardSheet.getCell('A6').value = 'Net Income:';
  dashboardSheet.getCell('B6').value = '=B4-B5';
  dashboardSheet.getCell('A7').value = 'Profit Margin:';
  dashboardSheet.getCell('B7').value = '=IF(B4>0,B6/B4,"")';
  dashboardSheet.getCell('B7').numFmt = '0.00%';

  // ===== SALES TRACKING =====
  const salesSheet = workbook.addWorksheet('Sales');
  salesSheet.columns = [
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Customer', key: 'customer', width: 20 },
    { header: 'Service/Product', key: 'service', width: 25 },
    { header: 'Quantity', key: 'quantity', width: 10 },
    { header: 'Amount', key: 'amount', width: 15 },
    { header: 'Payment Status', key: 'status', width: 15 }
  ];

  const salesData = [
    { date: '2024-01-15', customer: 'ABC Company', service: 'Consulting Services', quantity: 1, amount: 25000, status: 'Paid' },
    { date: '2024-01-20', customer: 'XYZ Ltd', service: 'Software Setup', quantity: 1, amount: 15000, status: 'Paid' },
    { date: '2024-02-01', customer: 'Tech Solutions', service: 'Training Session', quantity: 1, amount: 30000, status: 'Pending' },
    { date: '2024-02-05', customer: 'Local Business', service: 'Website Design', quantity: 1, amount: 45000, status: 'Paid' },
    { date: '2024-02-10', customer: 'Startup Inc', service: 'Business Plan', quantity: 1, amount: 35000, status: 'Paid' }
  ];

  salesData.forEach(sale => salesSheet.addRow(sale));

  // Add totals
  salesSheet.addRow({});
  salesSheet.addRow({
    date: 'TOTALS',
    amount: '=SUM(E2:E6)',
    service: ''
  });
  salesSheet.getCell('A8').font = { bold: true };
  salesSheet.getCell('E8').font = { bold: true };

  // ===== EXPENSE TRACKING =====
  const expensesSheet = workbook.addWorksheet('Expenses');
  expensesSheet.columns = [
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Category', key: 'category', width: 20 },
    { header: 'Description', key: 'description', width: 30 },
    { header: 'Amount', key: 'amount', width: 15 },
    { header: 'Payment Method', key: 'paymentMethod', width: 15 }
  ];

  const expenseData = [
    { date: '2024-01-05', category: 'Office Supplies', description: 'Stationery and printing', amount: 8000, paymentMethod: 'Cash' },
    { date: '2024-01-10', category: 'Rent', description: 'Monthly office rent', amount: 25000, paymentMethod: 'Bank Transfer' },
    { date: '2024-01-15', category: 'Utilities', description: 'Electricity and internet', amount: 12000, paymentMethod: 'M-Pesa' },
    { date: '2024-01-20', category: 'Marketing', description: 'Business cards and flyers', amount: 15000, paymentMethod: 'Cash' },
    { date: '2024-02-01', category: 'Travel', description: 'Client meeting transport', amount: 5000, paymentMethod: 'Cash' },
    { date: '2024-02-05', category: 'Equipment', description: 'Laptop purchase', amount: 45000, paymentMethod: 'Bank Transfer' }
  ];

  expenseData.forEach(expense => expensesSheet.addRow(expense));

  // Add totals
  expensesSheet.addRow({});
  expensesSheet.addRow({
    date: 'TOTALS',
    amount: '=SUM(D2:D7)',
    description: ''
  });
  expensesSheet.getCell('A9').font = { bold: true };
  expensesSheet.getCell('D9').font = { bold: true };

  // ===== CUSTOMER DATABASE =====
  const customersSheet = workbook.addWorksheet('Customers');
  customersSheet.columns = [
    { header: 'Customer Name', key: 'name', width: 25 },
    { header: 'Contact Person', key: 'contact', width: 20 },
    { header: 'Phone', key: 'phone', width: 15 },
    { header: 'Email', key: 'email', width: 25 },
    { header: 'Total Orders', key: 'orders', width: 12 },
    { header: 'Total Spent', key: 'spent', width: 15 },
    { header: 'Last Contact', key: 'lastContact', width: 12 }
  ];

  const customerData = [
    { name: 'ABC Company', contact: 'John Smith', phone: '+254712345678', email: 'john@abc.com', orders: 2, spent: 40000, lastContact: '2024-02-01' },
    { name: 'XYZ Ltd', contact: 'Jane Doe', phone: '+254723456789', email: 'jane@xyz.com', orders: 1, spent: 15000, lastContact: '2024-01-20' },
    { name: 'Tech Solutions', contact: 'Mike Johnson', phone: '+254734567890', email: 'mike@tech.com', orders: 1, spent: 30000, lastContact: '2024-02-01' },
    { name: 'Local Business', contact: 'Sarah Wilson', phone: '+254745678901', email: 'sarah@local.com', orders: 1, spent: 45000, lastContact: '2024-02-05' },
    { name: 'Startup Inc', contact: 'David Brown', phone: '+254756789012', email: 'david@startup.com', orders: 1, spent: 35000, lastContact: '2024-02-10' }
  ];

  customerData.forEach(customer => customersSheet.addRow(customer));

  // ===== SIMPLE BUDGET =====
  const budgetSheet = workbook.addWorksheet('Budget');
  budgetSheet.columns = [
    { header: 'Category', key: 'category', width: 20 },
    { header: 'Monthly Budget', key: 'budget', width: 15 },
    { header: 'Actual Spent', key: 'spent', width: 15 },
    { header: 'Remaining', key: 'remaining', width: 15 },
    { header: 'Status', key: 'status', width: 12 }
  ];

  const budgetData = [
    { category: 'Rent', budget: 25000, spent: 25000, status: 'On Budget' },
    { category: 'Utilities', budget: 15000, spent: 12000, status: 'Under Budget' },
    { category: 'Marketing', budget: 20000, spent: 15000, status: 'Under Budget' },
    { category: 'Office Supplies', budget: 10000, spent: 8000, status: 'Under Budget' },
    { category: 'Travel', budget: 15000, spent: 5000, status: 'Under Budget' },
    { category: 'Equipment', budget: 50000, spent: 45000, status: 'Under Budget' }
  ];

  budgetData.forEach(item => {
    item.remaining = item.budget - item.spent;
    budgetSheet.addRow(item);
  });

  // Add formulas
  for (let i = 2; i <= 7; i++) {
    budgetSheet.getCell(`D${i}`).value = `=B${i}-C${i}`;
  }

  await workbook.xlsx.writeFile(path.join(__dirname, '../templates/starter_bundle.xlsx'));
  console.log('Starter bundle created');
}

// Create Enterprise Bundle (all sheets)
async function createEnterpriseBundle() {
  const workbook = new ExcelJS.Workbook();

  // ===== EXECUTIVE DASHBOARD =====
  const dashboardSheet = workbook.addWorksheet('Executive Dashboard');
  dashboardSheet.getCell('A1').value = 'EXECUTIVE DASHBOARD - ENTERPRISE OVERVIEW';
  dashboardSheet.getCell('A1').font = { size: 18, bold: true, color: { argb: 'FF1F4E79' } };
  dashboardSheet.getCell('A1').alignment = { horizontal: 'center' };
  dashboardSheet.mergeCells('A1:G1');

  // Key Performance Indicators
  dashboardSheet.getCell('A3').value = 'KEY PERFORMANCE INDICATORS';
  dashboardSheet.getCell('A3').font = { bold: true, size: 14 };
  dashboardSheet.getCell('A4').value = 'Total Revenue (YTD):';
  dashboardSheet.getCell('B4').value = '=SUM(Sales!F:F)';
  dashboardSheet.getCell('A5').value = 'Total Operating Expenses:';
  dashboardSheet.getCell('B5').value = '=SUM(Expenses!D:D)';
  dashboardSheet.getCell('A6').value = 'Net Profit Margin:';
  dashboardSheet.getCell('B6').value = '=IF(B4>0,(B4-B5)/B4,"")';
  dashboardSheet.getCell('B6').numFmt = '0.00%';
  dashboardSheet.getCell('A7').value = 'Customer Acquisition Cost:';
  dashboardSheet.getCell('B7').value = '=SUM(Marketing!D:D)/COUNT(Customers!A:A)';
  dashboardSheet.getCell('A8').value = 'Employee Productivity:';
  dashboardSheet.getCell('B8').value = '=B4/COUNT(Payroll!A:A)';
  dashboardSheet.getCell('A9').value = 'Inventory Turnover:';
  dashboardSheet.getCell('B9').value = '=SUM(Sales!D:D)/SUM(Inventory!D:D)';

  // Quarterly Performance
  dashboardSheet.getCell('A11').value = 'QUARTERLY PERFORMANCE';
  dashboardSheet.getCell('A11').font = { bold: true, size: 14 };
  dashboardSheet.getCell('A12').value = 'Quarter';
  dashboardSheet.getCell('B12').value = 'Revenue';
  dashboardSheet.getCell('C12').value = 'Growth %';
  dashboardSheet.getCell('D12').value = 'Expenses';
  dashboardSheet.getCell('E12').value = 'Profit';

  const quarters = ['Q1 2024', 'Q2 2024', 'Q3 2024', 'Q4 2024'];
  quarters.forEach((quarter, index) => {
    dashboardSheet.getCell(`A${13+index}`).value = quarter;
    dashboardSheet.getCell(`B${13+index}`).value = Math.floor(Math.random() * 200000) + 50000;
    dashboardSheet.getCell(`D${13+index}`).value = Math.floor(Math.random() * 150000) + 30000;
    dashboardSheet.getCell(`E${13+index}`).value = `=B${13+index}-D${13+index}`;
    if (index > 0) {
      dashboardSheet.getCell(`C${13+index}`).value = `=(B${13+index}-B${12+index})/B${12+index}`;
      dashboardSheet.getCell(`C${13+index}`).numFmt = '0.00%';
    }
  });

  // ===== SALES & REVENUE ANALYSIS =====
  const salesSheet = workbook.addWorksheet('Sales & Revenue');
  salesSheet.columns = [
    { header: 'Transaction ID', key: 'transactionId', width: 15 },
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Customer ID', key: 'customerId', width: 12 },
    { header: 'Product Category', key: 'productCategory', width: 20 },
    { header: 'Product Name', key: 'productName', width: 25 },
    { header: 'Quantity', key: 'quantity', width: 10 },
    { header: 'Unit Price', key: 'unitPrice', width: 12 },
    { header: 'Discount %', key: 'discount', width: 10 },
    { header: 'Net Amount', key: 'netAmount', width: 15 },
    { header: 'Tax Amount', key: 'taxAmount', width: 12 },
    { header: 'Total Amount', key: 'totalAmount', width: 15 },
    { header: 'Payment Method', key: 'paymentMethod', width: 15 },
    { header: 'Sales Rep', key: 'salesRep', width: 15 },
    { header: 'Region', key: 'region', width: 12 }
  ];

  const salesData = [
    { transactionId: 'TXN001', date: '2024-01-15', customerId: 'CUST001', productCategory: 'Software', productName: 'Enterprise License', quantity: 10, unitPrice: 50000, discount: 5, paymentMethod: 'Bank Transfer', salesRep: 'Alice Johnson', region: 'Nairobi' },
    { transactionId: 'TXN002', date: '2024-01-20', customerId: 'CUST002', productCategory: 'Services', productName: 'Consulting Services', quantity: 1, unitPrice: 150000, discount: 0, paymentMethod: 'M-Pesa', salesRep: 'Bob Wilson', region: 'Mombasa' },
    { transactionId: 'TXN003', date: '2024-02-01', customerId: 'CUST003', productCategory: 'Hardware', productName: 'Server Equipment', quantity: 5, unitPrice: 200000, discount: 10, paymentMethod: 'Cheque', salesRep: 'Carol Davis', region: 'Kisumu' },
    { transactionId: 'TXN004', date: '2024-02-10', customerId: 'CUST004', productCategory: 'Training', productName: 'Advanced Training', quantity: 25, unitPrice: 8000, discount: 8, paymentMethod: 'Card', salesRep: 'David Brown', region: 'Eldoret' },
    { transactionId: 'TXN005', date: '2024-02-15', customerId: 'CUST005', productCategory: 'Software', productName: 'Cloud Services', quantity: 100, unitPrice: 2500, discount: 12, paymentMethod: 'Bank Transfer', salesRep: 'Eva Green', region: 'Nakuru' },
    { transactionId: 'TXN006', date: '2024-03-01', customerId: 'CUST001', productCategory: 'Support', productName: 'Premium Support', quantity: 1, unitPrice: 75000, discount: 0, paymentMethod: 'M-Pesa', salesRep: 'Alice Johnson', region: 'Nairobi' }
  ];

  salesData.forEach((sale, index) => {
    const grossAmount = sale.quantity * sale.unitPrice;
    const discountAmount = grossAmount * (sale.discount / 100);
    sale.netAmount = grossAmount - discountAmount;
    sale.taxAmount = sale.netAmount * 0.16; // 16% VAT
    sale.totalAmount = sale.netAmount + sale.taxAmount;
    salesSheet.addRow(sale);
  });

  // Add formulas for calculations
  for (let i = 2; i <= 7; i++) {
    salesSheet.getCell(`I${i}`).value = `=F${i}*G${i}*(1-H${i}/100)`; // Net Amount
    salesSheet.getCell(`J${i}`).value = `=I${i}*0.16`; // Tax Amount (16% VAT)
    salesSheet.getCell(`K${i}`).value = `=I${i}+J${i}`; // Total Amount
  }

  // ===== COMPREHENSIVE EXPENSES =====
  const expensesSheet = workbook.addWorksheet('Expenses Analysis');
  expensesSheet.columns = [
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Expense ID', key: 'expenseId', width: 12 },
    { header: 'Category', key: 'category', width: 20 },
    { header: 'Sub-Category', key: 'subCategory', width: 20 },
    { header: 'Description', key: 'description', width: 35 },
    { header: 'Vendor/Supplier', key: 'vendor', width: 20 },
    { header: 'Amount', key: 'amount', width: 15 },
    { header: 'Tax Deductible', key: 'taxDeductible', width: 15 },
    { header: 'Payment Method', key: 'paymentMethod', width: 15 },
    { header: 'Department', key: 'department', width: 15 },
    { header: 'Approved By', key: 'approvedBy', width: 15 }
  ];

  const expenseData = [
    { expenseId: 'EXP001', date: '2024-01-05', category: 'Operations', subCategory: 'Office Supplies', description: 'Stationery, printing materials, and office consumables', vendor: 'Office Depot', amount: 45000, taxDeductible: 'Yes', paymentMethod: 'Cash', department: 'Admin', approvedBy: 'John Manager' },
    { expenseId: 'EXP002', date: '2024-01-10', category: 'Facilities', subCategory: 'Rent', description: 'Monthly office rent for main branch', vendor: 'Property Management Ltd', amount: 250000, taxDeductible: 'Yes', paymentMethod: 'Bank Transfer', department: 'Facilities', approvedBy: 'Sarah Director' },
    { expenseId: 'EXP003', date: '2024-01-15', category: 'Operations', subCategory: 'Utilities', description: 'Electricity, water, and internet services', vendor: 'Kenya Power', amount: 85000, taxDeductible: 'Yes', paymentMethod: 'M-Pesa', department: 'Admin', approvedBy: 'John Manager' },
    { expenseId: 'EXP004', date: '2024-01-20', category: 'Marketing', subCategory: 'Digital Advertising', description: 'Google Ads and social media campaigns', vendor: 'Digital Marketing Agency', amount: 120000, taxDeductible: 'Yes', paymentMethod: 'Card', department: 'Marketing', approvedBy: 'Mike Marketing' },
    { expenseId: 'EXP005', date: '2024-01-25', category: 'Human Resources', subCategory: 'Salaries', description: 'Monthly employee salaries and wages', vendor: 'Internal', amount: 850000, taxDeductible: 'No', paymentMethod: 'Bank Transfer', department: 'HR', approvedBy: 'HR Director' },
    { expenseId: 'EXP006', date: '2024-02-01', category: 'Operations', subCategory: 'Travel', description: 'Business travel and accommodation', vendor: 'Travel Agency', amount: 150000, taxDeductible: 'Yes', paymentMethod: 'Cash', department: 'Sales', approvedBy: 'Sales Manager' },
    { expenseId: 'EXP007', date: '2024-02-05', category: 'IT', subCategory: 'Software Licenses', description: 'Annual software licenses and subscriptions', vendor: 'Microsoft', amount: 300000, taxDeductible: 'Yes', paymentMethod: 'Bank Transfer', department: 'IT', approvedBy: 'IT Manager' },
    { expenseId: 'EXP008', date: '2024-02-10', category: 'Operations', subCategory: 'Equipment', description: 'Office equipment and furniture', vendor: 'Office Furniture Ltd', amount: 200000, taxDeductible: 'Yes', paymentMethod: 'Cheque', department: 'Admin', approvedBy: 'John Manager' }
  ];

  expenseData.forEach(expense => expensesSheet.addRow(expense));

  // ===== ADVANCED INVENTORY MANAGEMENT =====
  const inventorySheet = workbook.addWorksheet('Inventory Management');
  inventorySheet.columns = [
    { header: 'Item Code', key: 'itemCode', width: 12 },
    { header: 'Item Name', key: 'itemName', width: 25 },
    { header: 'Category', key: 'category', width: 15 },
    { header: 'Sub-Category', key: 'subCategory', width: 15 },
    { header: 'Current Stock', key: 'currentStock', width: 12 },
    { header: 'Minimum Stock', key: 'minStock', width: 12 },
    { header: 'Maximum Stock', key: 'maxStock', width: 12 },
    { header: 'Unit Cost', key: 'unitCost', width: 12 },
    { header: 'Selling Price', key: 'sellingPrice', width: 12 },
    { header: 'Total Value', key: 'totalValue', width: 15 },
    { header: 'Supplier', key: 'supplier', width: 20 },
    { header: 'Location', key: 'location', width: 15 },
    { header: 'Last Restock', key: 'lastRestock', width: 12 },
    { header: 'Reorder Point', key: 'reorderPoint', width: 12 }
  ];

  const inventoryData = [
    { itemCode: 'ITM001', itemName: 'Executive Office Chairs', category: 'Furniture', subCategory: 'Seating', currentStock: 25, minStock: 5, maxStock: 50, unitCost: 15000, sellingPrice: 25000, supplier: 'Premium Furniture Ltd', location: 'Warehouse A', lastRestock: '2024-01-10', reorderPoint: 10 },
    { itemCode: 'ITM002', itemName: 'A4 Premium Paper', category: 'Supplies', subCategory: 'Paper', currentStock: 1000, minStock: 200, maxStock: 2000, unitCost: 80, sellingPrice: 120, supplier: 'Paper World', location: 'Storage B', lastRestock: '2024-01-15', reorderPoint: 300 },
    { itemCode: 'ITM003', itemName: 'High-Yield Printer Cartridges', category: 'Supplies', subCategory: 'Ink/Toner', currentStock: 50, minStock: 10, maxStock: 100, unitCost: 5000, sellingPrice: 7500, supplier: 'Tech Supplies', location: 'Storage C', lastRestock: '2024-01-20', reorderPoint: 15 },
    { itemCode: 'ITM004', itemName: 'Gaming Laptops', category: 'Equipment', subCategory: 'Computers', currentStock: 15, minStock: 3, maxStock: 30, unitCost: 120000, sellingPrice: 180000, supplier: 'Computer World', location: 'Warehouse A', lastRestock: '2024-02-01', reorderPoint: 5 },
    { itemCode: 'ITM005', itemName: 'Commercial Coffee Machine', category: 'Equipment', subCategory: 'Appliances', currentStock: 3, minStock: 1, maxStock: 5, unitCost: 45000, sellingPrice: 65000, supplier: 'Kitchen Equip', location: 'Break Room', lastRestock: '2024-01-05', reorderPoint: 2 },
    { itemCode: 'ITM006', itemName: 'Wireless Projectors', category: 'Equipment', subCategory: 'AV Equipment', currentStock: 8, minStock: 2, maxStock: 15, unitCost: 80000, sellingPrice: 120000, supplier: 'AV Solutions', location: 'Conference Room', lastRestock: '2024-02-05', reorderPoint: 3 }
  ];

  inventoryData.forEach(item => {
    item.totalValue = item.currentStock * item.unitCost;
    item.reorderPoint = Math.ceil(item.minStock * 1.5);
    inventorySheet.addRow(item);
  });

  // Add formulas
  for (let i = 2; i <= 7; i++) {
    inventorySheet.getCell(`J${i}`).value = `=D${i}*H${i}`; // Total Value
    inventorySheet.getCell(`N${i}`).value = `=CEILING(F${i}*1.5,1)`; // Reorder Point
  }

  // ===== CUSTOMER RELATIONSHIP MANAGEMENT =====
  const crmSheet = workbook.addWorksheet('CRM Database');
  crmSheet.columns = [
    { header: 'Customer ID', key: 'customerId', width: 12 },
    { header: 'Company Name', key: 'companyName', width: 25 },
    { header: 'Industry', key: 'industry', width: 15 },
    { header: 'Contact Person', key: 'contactPerson', width: 20 },
    { header: 'Job Title', key: 'jobTitle', width: 20 },
    { header: 'Phone', key: 'phone', width: 15 },
    { header: 'Email', key: 'email', width: 25 },
    { header: 'Website', key: 'website', width: 20 },
    { header: 'Total Orders', key: 'totalOrders', width: 12 },
    { header: 'Total Spent', key: 'totalSpent', width: 15 },
    { header: 'Average Order Value', key: 'avgOrderValue', width: 15 },
    { header: 'Last Order Date', key: 'lastOrderDate', width: 15 },
    { header: 'Customer Since', key: 'customerSince', width: 12 },
    { header: 'Loyalty Tier', key: 'loyaltyTier', width: 12 },
    { header: 'Lead Source', key: 'leadSource', width: 15 }
  ];

  const crmData = [
    { customerId: 'CUST001', companyName: 'TechCorp Solutions Ltd', industry: 'Technology', contactPerson: 'John Smith', jobTitle: 'CEO', phone: '+254712345678', email: 'john@techcorp.com', website: 'www.techcorp.com', totalOrders: 15, totalSpent: 2500000, lastOrderDate: '2024-02-15', customerSince: '2022-03-01', loyaltyTier: 'Platinum', leadSource: 'Referral' },
    { customerId: 'CUST002', companyName: 'Global Industries Ltd', industry: 'Manufacturing', contactPerson: 'Jane Doe', jobTitle: 'Procurement Manager', phone: '+254723456789', email: 'jane@globalind.com', website: 'www.globalind.com', totalOrders: 8, totalSpent: 1800000, lastOrderDate: '2024-02-10', customerSince: '2023-01-15', loyaltyTier: 'Gold', leadSource: 'Trade Show' },
    { customerId: 'CUST003', companyName: 'Innovate Solutions', industry: 'Consulting', contactPerson: 'Mike Johnson', jobTitle: 'Partner', phone: '+254734567890', email: 'mike@innovate.com', website: 'www.innovate.com', totalOrders: 5, totalSpent: 950000, lastOrderDate: '2024-02-01', customerSince: '2023-06-20', loyaltyTier: 'Gold', leadSource: 'Website' },
    { customerId: 'CUST004', companyName: 'Premier Logistics', industry: 'Transportation', contactPerson: 'Sarah Wilson', jobTitle: 'Operations Director', phone: '+254745678901', email: 'sarah@premierlog.com', website: 'www.premierlog.com', totalOrders: 22, totalSpent: 3200000, lastOrderDate: '2024-02-12', customerSince: '2021-11-10', loyaltyTier: 'Platinum', leadSource: 'Direct Sales' },
    { customerId: 'CUST005', companyName: 'Elite Enterprises', industry: 'Retail', contactPerson: 'David Brown', jobTitle: 'Owner', phone: '+254756789012', email: 'david@elite.com', website: 'www.elite.com', totalOrders: 12, totalSpent: 1650000, lastOrderDate: '2024-02-08', customerSince: '2022-08-05', loyaltyTier: 'Gold', leadSource: 'Social Media' }
  ];

  crmData.forEach(customer => {
    customer.avgOrderValue = customer.totalSpent / customer.totalOrders;
    crmSheet.addRow(customer);
  });

  // Add formula for average order value
  for (let i = 2; i <= 6; i++) {
    crmSheet.getCell(`K${i}`).value = `=J${i}/I${i}`;
  }

  // ===== HUMAN RESOURCES & PAYROLL =====
  const hrSheet = workbook.addWorksheet('HR & Payroll');
  hrSheet.columns = [
    { header: 'Employee ID', key: 'employeeId', width: 12 },
    { header: 'Full Name', key: 'fullName', width: 20 },
    { header: 'Department', key: 'department', width: 15 },
    { header: 'Job Title', key: 'jobTitle', width: 20 },
    { header: 'Employment Type', key: 'employmentType', width: 15 },
    { header: 'Basic Salary', key: 'basicSalary', width: 15 },
    { header: 'House Allowance', key: 'houseAllowance', width: 15 },
    { header: 'Transport Allowance', key: 'transportAllowance', width: 18 },
    { header: 'Medical Allowance', key: 'medicalAllowance', width: 15 },
    { header: 'Other Benefits', key: 'otherBenefits', width: 15 },
    { header: 'Gross Salary', key: 'grossSalary', width: 15 },
    { header: 'PAYE Tax', key: 'payeTax', width: 12 },
    { header: 'NHIF', key: 'nhif', width: 10 },
    { header: 'NSSF', key: 'nssf', width: 10 },
    { header: 'Net Salary', key: 'netSalary', width: 15 },
    { header: 'Join Date', key: 'joinDate', width: 12 }
  ];

  const hrData = [
    { employeeId: 'EMP001', fullName: 'Alice Johnson', department: 'Sales', jobTitle: 'Senior Sales Manager', employmentType: 'Permanent', basicSalary: 85000, houseAllowance: 25000, transportAllowance: 15000, medicalAllowance: 10000, otherBenefits: 5000, joinDate: '2022-01-15' },
    { employeeId: 'EMP002', fullName: 'Bob Wilson', department: 'Marketing', jobTitle: 'Marketing Director', employmentType: 'Permanent', basicSalary: 95000, houseAllowance: 30000, transportAllowance: 18000, medicalAllowance: 12000, otherBenefits: 8000, joinDate: '2021-08-20' },
    { employeeId: 'EMP003', fullName: 'Carol Davis', department: 'IT', jobTitle: 'IT Manager', employmentType: 'Permanent', basicSalary: 120000, houseAllowance: 35000, transportAllowance: 20000, medicalAllowance: 15000, otherBenefits: 10000, joinDate: '2020-03-10' },
    { employeeId: 'EMP004', fullName: 'David Brown', department: 'Finance', jobTitle: 'Finance Manager', employmentType: 'Permanent', basicSalary: 100000, houseAllowance: 28000, transportAllowance: 16000, medicalAllowance: 13000, otherBenefits: 7000, joinDate: '2021-11-05' },
    { employeeId: 'EMP005', fullName: 'Eva Green', department: 'HR', jobTitle: 'HR Manager', employmentType: 'Permanent', basicSalary: 80000, houseAllowance: 22000, transportAllowance: 14000, medicalAllowance: 11000, otherBenefits: 6000, joinDate: '2022-05-12' },
    { employeeId: 'EMP006', fullName: 'Frank Miller', department: 'Operations', jobTitle: 'Operations Manager', employmentType: 'Permanent', basicSalary: 90000, houseAllowance: 26000, transportAllowance: 17000, medicalAllowance: 12000, otherBenefits: 7500, joinDate: '2021-09-18' }
  ];

  hrData.forEach(employee => {
    employee.grossSalary = employee.basicSalary + employee.houseAllowance + employee.transportAllowance + employee.medicalAllowance + employee.otherBenefits;
    // Tax calculations based on Kenyan tax brackets (simplified)
    let taxableIncome = employee.grossSalary;
    let paye = 0;
    if (taxableIncome <= 24000) paye = 0;
    else if (taxableIncome <= 40667) paye = (taxableIncome - 24000) * 0.1;
    else if (taxableIncome <= 57333) paye = 1667 + (taxableIncome - 40667) * 0.15;
    else if (taxableIncome <= 74000) paye = 4167 + (taxableIncome - 57333) * 0.20;
    else paye = 8167 + (taxableIncome - 74000) * 0.25;
    employee.payeTax = Math.round(paye);

    employee.nhif = taxableIncome > 100000 ? 1700 : taxableIncome > 80000 ? 1400 : taxableIncome > 60000 ? 1100 : taxableIncome > 40000 ? 750 : 500;
    employee.nssf = Math.min(taxableIncome * 0.12, 2160);
    employee.netSalary = employee.grossSalary - employee.payeTax - employee.nhif - employee.nssf;
    hrSheet.addRow(employee);
  });

  // Add formulas for calculations
  for (let i = 2; i <= 7; i++) {
    hrSheet.getCell(`K${i}`).value = `=F${i}+G${i}+H${i}+I${i}+J${i}`; // Gross Salary
    hrSheet.getCell(`O${i}`).value = `=K${i}-L${i}-M${i}-N${i}`; // Net Salary
  }

  // ===== MARKETING & CAMPAIGNS =====
  const marketingSheet = workbook.addWorksheet('Marketing Analytics');
  marketingSheet.columns = [
    { header: 'Campaign ID', key: 'campaignId', width: 12 },
    { header: 'Campaign Name', key: 'campaignName', width: 25 },
    { header: 'Start Date', key: 'startDate', width: 12 },
    { header: 'End Date', key: 'endDate', width: 12 },
    { header: 'Budget', key: 'budget', width: 15 },
    { header: 'Spent', key: 'spent', width: 15 },
    { header: 'Impressions', key: 'impressions', width: 15 },
    { header: 'Clicks', key: 'clicks', width: 12 },
    { header: 'Conversions', key: 'conversions', width: 12 },
    { header: 'Revenue Generated', key: 'revenue', width: 15 },
    { header: 'ROI %', key: 'roi', width: 10 },
    { header: 'Channel', key: 'channel', width: 15 },
    { header: 'Target Audience', key: 'targetAudience', width: 20 }
  ];

  const marketingData = [
    { campaignId: 'CMP001', campaignName: 'Q1 Digital Campaign', startDate: '2024-01-01', endDate: '2024-03-31', budget: 500000, spent: 450000, impressions: 2500000, clicks: 25000, conversions: 1250, revenue: 2500000, channel: 'Google Ads', targetAudience: 'SMEs' },
    { campaignId: 'CMP002', campaignName: 'Social Media Boost', startDate: '2024-02-01', endDate: '2024-02-28', budget: 150000, spent: 135000, impressions: 1800000, clicks: 18000, conversions: 900, revenue: 1800000, channel: 'Facebook/Instagram', targetAudience: 'Startups' },
    { campaignId: 'CMP003', campaignName: 'Email Marketing Series', startDate: '2024-01-15', endDate: '2024-03-15', budget: 80000, spent: 72000, impressions: 50000, clicks: 2500, conversions: 125, revenue: 500000, channel: 'Email', targetAudience: 'Existing Customers' },
    { campaignId: 'CMP004', campaignName: 'Trade Show Presence', startDate: '2024-02-10', endDate: '2024-02-12', budget: 300000, spent: 285000, impressions: 15000, clicks: 0, conversions: 75, revenue: 1500000, channel: 'Events', targetAudience: 'Enterprise Clients' }
  ];

  marketingData.forEach(campaign => {
    campaign.roi = ((campaign.revenue - campaign.spent) / campaign.spent) * 100;
    marketingSheet.addRow(campaign);
  });

  // Add formula for ROI
  for (let i = 2; i <= 5; i++) {
    marketingSheet.getCell(`K${i}`).value = `=((J${i}-F${i})/F${i})*100`;
  }

  await workbook.xlsx.writeFile(path.join(__dirname, '../templates/enterprise_bundle.xlsx'));
  console.log('Enterprise bundle created');
}

// Generate all templates
async function generateAllTemplates() {
  try {
    await createStarterBundle();
    await createBusinessBundle();
    await createEnterpriseBundle();
    console.log('All bundles generated successfully');
  } catch (error) {
    console.error('Error generating bundles:', error);
  }
}

generateAllTemplates();