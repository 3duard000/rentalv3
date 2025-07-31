// TenantManager.gs - Tenant Management System
// Handles all tenant-related operations

const TenantManager = {
  
  /**
   * Column indexes for Tenants sheet (1-based)
   */
  COL: {
    ROOM_NUMBER: 1,
    RENTAL_PRICE: 2,
    NEGOTIATED_PRICE: 3,
    TENANT_NAME: 4,
    TENANT_EMAIL: 5,
    TENANT_PHONE: 6,
    MOVE_IN_DATE: 7,
    SECURITY_DEPOSIT: 8,
    ROOM_STATUS: 9,
    LAST_PAYMENT: 10,
    PAYMENT_STATUS: 11,
    MOVE_OUT_PLANNED: 12,
    EMERGENCY_CONTACT: 13,
    LEASE_END_DATE: 14,
    NOTES: 15
  },

  /**
   * Column indexes for Tenant Applications sheet (1-based)
   */
  APP_COL: {
    APPLICATION_ID: 1,
    TIMESTAMP: 2,
    FULL_NAME: 3,
    EMAIL: 4,
    PHONE: 5,
    CURRENT_ADDRESS: 6,
    MOVE_IN_DATE: 7,
    PREFERRED_ROOM: 8
  },
  
  /**
   * Check and update payment status for all tenants
   */
  checkAllPaymentStatus: function() {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        SpreadsheetApp.getUi().alert('No tenant data found.');
        return;
      }
      
      const today = new Date();
      const firstThisMonth = new Date(today.getFullYear(), today.getMonth(), 1);
      const firstLastMonth = new Date(today.getFullYear(), today.getMonth() - 1, 1);
      
      let updatedCount = 0;
      
      data.forEach((row, index) => {
        const rowNumber = index + 2; // Account for header row
        const roomStatus = row[this.COL.ROOM_STATUS - 1];
        
        if (roomStatus !== CONFIG.STATUS.ROOM.OCCUPIED) {
          // Clear payment status for non-occupied rooms
          sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue('');
          return;
        }
        
        const lastPayment = row[this.COL.LAST_PAYMENT - 1];
        let status = CONFIG.STATUS.PAYMENT.OVERDUE;
        
        if (lastPayment instanceof Date) {
          if (lastPayment >= firstThisMonth) {
            status = CONFIG.STATUS.PAYMENT.PAID;
          } else if (lastPayment >= firstLastMonth) {
            status = CONFIG.STATUS.PAYMENT.DUE;
          }
        }
        
        // Update payment status
        sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue(status);
        updatedCount++;
      });
      
      SpreadsheetApp.getUi().alert(
        'Payment Status Updated',
        `Updated payment status for ${updatedCount} tenants.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'checkAllPaymentStatus');
    }
  },
  
  /**
   * Send rent reminders to tenants with due or overdue payments
   */
  sendRentReminders: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        SpreadsheetApp.getUi().alert('No tenant data found.');
        return;
      }
      
      const monthYear = Utils.formatDate(new Date(), 'MMMM yyyy');
      let sentCount = 0;
      let failedCount = 0;
      
      data.forEach(row => {
        const paymentStatus = row[this.COL.PAYMENT_STATUS - 1];
        const email = row[this.COL.TENANT_EMAIL - 1];
        
        if ((paymentStatus === CONFIG.STATUS.PAYMENT.DUE || paymentStatus === CONFIG.STATUS.PAYMENT.OVERDUE) && email) {
          const tenantName = row[this.COL.TENANT_NAME - 1];
          const roomNumber = row[this.COL.ROOM_NUMBER - 1];
          const rent = row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1];
          
          const emailData = {
            tenantName: tenantName,
            roomNumber: roomNumber,
            rent: rent,
            status: paymentStatus,
            monthYear: monthYear,
            dueDate: this.calculateRentDueDate(),
            lateFee: paymentStatus === CONFIG.STATUS.PAYMENT.OVERDUE ? CONFIG.SYSTEM.LATE_FEE_AMOUNT : 0
          };
          
          try {
            EmailManager.sendRentReminder(email, emailData);
            sentCount++;
          } catch (emailError) {
            Logger.log(`Failed to send reminder to ${email}: ${emailError.message}`);
            failedCount++;
          }
        }
      });
      
      const message = `Rent reminders sent: ${sentCount}` + 
                     (failedCount > 0 ? `\nFailed to send: ${failedCount}` : '');
      
      SpreadsheetApp.getUi().alert('Rent Reminders', message, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      handleSystemError(error, 'sendRentReminders');
    }
  },
  
  /**
   * Send late payment alerts to manager
   */
  sendLatePaymentAlerts: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        return;
      }
      
      const overdueList = [];
      
      data.forEach(row => {
        if (row[this.COL.PAYMENT_STATUS - 1] === CONFIG.STATUS.PAYMENT.OVERDUE) {
          const tenant = row[this.COL.TENANT_NAME - 1];
          const email = row[this.COL.TENANT_EMAIL - 1];
          const room = row[this.COL.ROOM_NUMBER - 1];
          const lastPayment = row[this.COL.LAST_PAYMENT - 1];
          const lastPaymentStr = lastPayment ? Utils.formatDate(lastPayment) : 'Never';
          
          overdueList.push({
            tenant: tenant,
            email: email,
            room: room,
            lastPayment: lastPaymentStr
          });
        }
      });
      
      if (overdueList.length === 0) {
        SpreadsheetApp.getUi().alert('No overdue tenants found.');
        return;
      }
      
      EmailManager.sendLatePaymentAlert(CONFIG.SYSTEM.MANAGER_EMAIL, {
        overdueList: overdueList,
        count: overdueList.length
      });
      
      SpreadsheetApp.getUi().alert(
        'Late Payment Alert Sent',
        `Alert sent to manager about ${overdueList.length} overdue tenant(s).`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      handleSystemError(error, 'sendLatePaymentAlerts');
    }
  },
  
  /**
   * Send monthly invoices to all occupied rooms
   */
  sendMonthlyInvoices: function() {
    try {
      const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
      
      if (data.length === 0) {
        SpreadsheetApp.getUi().alert('No tenant data found.');
        return;
      }
      
      const monthYear = Utils.formatDate(new Date(), 'MMMM yyyy');
      let sentCount = 0;
      let failedCount = 0;
      
      data.forEach(row => {
        const roomStatus = row[this.COL.ROOM_STATUS - 1];
        const email = row[this.COL.TENANT_EMAIL - 1];
        
        if (roomStatus === CONFIG.STATUS.ROOM.OCCUPIED && email) {
          const tenant = row[this.COL.TENANT_NAME - 1];
          const room = row[this.COL.ROOM_NUMBER - 1];
          const rent = row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1];
          const dueDate = this.calculateRentDueDate();
          
          try {
            const pdf = DocumentManager.createRentInvoice({
              tenantName: tenant,
              roomNumber: room,
              rent: rent,
              monthYear: monthYear,
              dueDate: dueDate,
              propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
              email: row[this.COL.TENANT_EMAIL - 1],
              phone: row[this.COL.TENANT_PHONE - 1]
            });

            EmailManager.sendMonthlyInvoice(email, {
              tenantName: tenant,
              monthYear: monthYear
            }, pdf);
            
            sentCount++;
          } catch (invoiceError) {
            Logger.log(`Failed to send invoice to ${email}: ${invoiceError.message}`);
            failedCount++;
          }
        }
      });
      
      const message = `Monthly invoices sent: ${sentCount}` + 
                     (failedCount > 0 ? `\nFailed to send: ${failedCount}` : '');
      
      SpreadsheetApp.getUi().alert('Monthly Invoices', message, SpreadsheetApp.getUi().ButtonSet.OK);
      
    } catch (error) {
      handleSystemError(error, 'sendMonthlyInvoices');
    }
  },
  
  /**
 * Mark payment as received for selected tenant (improved version)
 */
markPaymentReceived: function() {
  try {
    const ui = SpreadsheetApp.getUi();
    const tenants = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);

    const occupiedTenants = tenants
      .map((row, i) => ({
        index: i + 2,
        name: row[this.COL.TENANT_NAME - 1],
        room: row[this.COL.ROOM_NUMBER - 1],
        email: row[this.COL.TENANT_EMAIL - 1],
        rent: row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1],
        status: row[this.COL.PAYMENT_STATUS - 1],
        lastPayment: row[this.COL.LAST_PAYMENT - 1]
      }))
      .filter(t => t.name && row[this.COL.ROOM_STATUS - 1] === CONFIG.STATUS.ROOM.OCCUPIED);

    if (occupiedTenants.length === 0) {
      ui.alert('No occupied tenants found.');
      return;
    }

    const tenantOptions = occupiedTenants.map(t => 
      `<option value="${t.index}" data-rent="${t.rent}" data-status="${t.status}" data-email="${t.email}">
        ${t.name} (Room ${t.room}) - ${t.status || 'Unknown Status'}
      </option>`
    ).join('');

    const html = HtmlService.createHtmlOutput(`
      <style>
        .payment-form {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          max-width: 500px;
          margin: 0 auto;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          padding: 20px;
          box-sizing: border-box;
        }
        .form-container {
          background: white;
          padding: 30px;
          border-radius: 15px;
          box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        .form-title {
          color: #2c3e50;
          text-align: center;
          margin-bottom: 25px;
          font-size: 24px;
          font-weight: 600;
        }
        .form-group {
          margin-bottom: 20px;
        }
        .form-label {
          display: block;
          margin-bottom: 6px;
          color: #34495e;
          font-weight: 500;
          font-size: 14px;
        }
        .form-control {
          width: 100%;
          padding: 12px 15px;
          border: 2px solid #e9ecef;
          border-radius: 8px;
          font-size: 14px;
          transition: all 0.3s ease;
          box-sizing: border-box;
        }
        .form-control:focus {
          outline: none;
          border-color: #667eea;
          box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        .tenant-info {
          background: #f8f9fa;
          padding: 15px;
          border-radius: 8px;
          border-left: 4px solid #17a2b8;
          margin-bottom: 20px;
          display: none;
        }
        .info-item {
          margin-bottom: 8px;
          font-size: 13px;
          color: #495057;
        }
        .info-label {
          font-weight: 600;
          color: #2c3e50;
        }
        .btn {
          padding: 12px 25px;
          border: none;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.3s ease;
          margin-right: 10px;
        }
        .btn-primary {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
        }
        .btn-primary:hover {
          transform: translateY(-2px);
          box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        .btn-secondary {
          background: #6c757d;
          color: white;
        }
        .btn-secondary:hover {
          background: #5a6268;
        }
        .form-footer {
          text-align: center;
          margin-top: 25px;
        }
        .status-badge {
          display: inline-block;
          padding: 4px 12px;
          border-radius: 15px;
          font-size: 12px;
          font-weight: 600;
          text-transform: uppercase;
        }
        .status-paid { background: #d4edda; color: #155724; }
        .status-due { background: #fff3cd; color: #856404; }
        .status-overdue { background: #f8d7da; color: #721c24; }
        .validation-error {
          color: #dc3545;
          font-size: 12px;
          margin-top: 5px;
          display: none;
        }
        .payment-method-grid {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 10px;
        }
      </style>

      <div class="payment-form">
        <div class="form-container">
          <h2 class="form-title">💰 Record Rent Payment</h2>
          
          <form id="paymentForm">
            <div class="form-group">
              <label class="form-label" for="tenantSelect">Select Tenant *</label>
              <select id="tenantSelect" class="form-control" required onchange="updateTenantInfo()">
                <option value="">Choose a tenant...</option>
                ${tenantOptions}
              </select>
              <div class="validation-error" id="tenantError">Please select a tenant</div>
            </div>

            <div id="tenantInfo" class="tenant-info">
              <div class="info-item">
                <span class="info-label">Monthly Rent:</span> 
                <span id="rentAmount">$0.00</span>
              </div>
              <div class="info-item">
                <span class="info-label">Current Status:</span> 
                <span id="currentStatus" class="status-badge">Unknown</span>
              </div>
              <div class="info-item">
                <span class="info-label">Email:</span> 
                <span id="tenantEmail">-</span>
              </div>
            </div>

            <div class="form-group">
              <label class="form-label" for="paymentDate">Payment Date *</label>
              <input type="date" id="paymentDate" class="form-control" 
                     value="${Utils.formatDate(new Date(), 'yyyy-MM-dd')}" 
                     max="${Utils.formatDate(new Date(), 'yyyy-MM-dd')}" required>
              <div class="validation-error" id="dateError">Please select a valid date</div>
            </div>

            <div class="form-group">
              <label class="form-label" for="paymentAmount">Payment Amount *</label>
              <input type="number" id="paymentAmount" class="form-control" 
                     step="0.01" min="0" placeholder="0.00" required>
              <div class="validation-error" id="amountError">Please enter a valid amount</div>
            </div>

            <div class="form-group">
              <label class="form-label" for="paymentMethod">Payment Method</label>
              <select id="paymentMethod" class="form-control">
                <option value="Cash">Cash</option>
                <option value="Check">Check</option>
                <option value="Bank Transfer">Bank Transfer</option>
                <option value="Credit Card">Credit Card</option>
                <option value="Money Order">Money Order</option>
                <option value="Online Payment">Online Payment</option>
              </select>
            </div>

            <div class="form-group">
              <label class="form-label" for="referenceNumber">Reference Number</label>
              <input type="text" id="referenceNumber" class="form-control" 
                     placeholder="Check #, Transaction ID, etc.">
            </div>

            <div class="form-group">
              <label class="form-label" for="paymentNotes">Notes</label>
              <textarea id="paymentNotes" class="form-control" rows="3" 
                        placeholder="Additional notes about this payment..."></textarea>
            </div>

            <div class="form-footer">
              <button type="button" class="btn btn-primary" onclick="submitPayment()">
                Record Payment
              </button>
              <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">
                Cancel
              </button>
            </div>
          </form>
        </div>
      </div>

      <script>
        function updateTenantInfo() {
          const select = document.getElementById('tenantSelect');
          const info = document.getElementById('tenantInfo');
          const amountField = document.getElementById('paymentAmount');
          
          if (select.value) {
            const option = select.options[select.selectedIndex];
            const rent = parseFloat(option.dataset.rent) || 0;
            const status = option.dataset.status || 'Unknown';
            const email = option.dataset.email || '-';
            
            document.getElementById('rentAmount').textContent = '$' + rent.toFixed(2);
            document.getElementById('tenantEmail').textContent = email;
            
            const statusSpan = document.getElementById('currentStatus');
            statusSpan.textContent = status;
            statusSpan.className = 'status-badge status-' + status.toLowerCase();
            
            amountField.value = rent.toFixed(2);
            info.style.display = 'block';
          } else {
            info.style.display = 'none';
            amountField.value = '';
          }
          clearValidationErrors();
        }

        function validateForm() {
          let isValid = true;
          clearValidationErrors();

          const tenant = document.getElementById('tenantSelect').value;
          const date = document.getElementById('paymentDate').value;
          const amount = parseFloat(document.getElementById('paymentAmount').value);

          if (!tenant) {
            showError('tenantError', 'Please select a tenant');
            isValid = false;
          }

          if (!date) {
            showError('dateError', 'Please select a payment date');
            isValid = false;
          } else {
            const paymentDate = new Date(date);
            const today = new Date();
            if (paymentDate > today) {
              showError('dateError', 'Payment date cannot be in the future');
              isValid = false;
            }
          }

          if (!amount || amount <= 0) {
            showError('amountError', 'Please enter a valid payment amount');
            isValid = false;
          }

          return isValid;
        }

        function showError(elementId, message) {
          const errorElement = document.getElementById(elementId);
          errorElement.textContent = message;
          errorElement.style.display = 'block';
        }

        function clearValidationErrors() {
          const errors = document.querySelectorAll('.validation-error');
          errors.forEach(error => error.style.display = 'none');
        }

        function submitPayment() {
          if (!validateForm()) {
            return;
          }

          const formData = {
            rowNumber: document.getElementById('tenantSelect').value,
            paymentDate: document.getElementById('paymentDate').value,
            paymentAmount: parseFloat(document.getElementById('paymentAmount').value),
            paymentMethod: document.getElementById('paymentMethod').value,
            referenceNumber: document.getElementById('referenceNumber').value,
            notes: document.getElementById('paymentNotes').value
          };

          const submitBtn = document.querySelector('.btn-primary');
          submitBtn.disabled = true;
          submitBtn.textContent = 'Processing...';

          google.script.run
            .withSuccessHandler(function() {
              alert('Payment recorded successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('Error recording payment: ' + error.message);
              submitBtn.disabled = false;
              submitBtn.textContent = 'Record Payment';
            })
            .recordTenantPaymentAdvanced(formData);
        }
      </script>
    `).setWidth(550).setHeight(700);

    ui.showModalDialog(html, 'Record Payment');

  } catch (error) {
    handleSystemError(error, 'markPaymentReceived');
  }
},

/**
 * Record tenant payment with advanced data (called from improved form)
 */
recordTenantPaymentAdvanced: function(data) {
  try {
    const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
    const row = parseInt(data.rowNumber);
    
    // Update tenant record
    sheet.getRange(row, this.COL.LAST_PAYMENT).setValue(new Date(data.paymentDate));
    sheet.getRange(row, this.COL.PAYMENT_STATUS).setValue(CONFIG.STATUS.PAYMENT.PAID);
    
    // Add notes if provided
    if (data.notes) {
      const currentNotes = sheet.getRange(row, this.COL.NOTES).getValue() || '';
      const newNotes = currentNotes + (currentNotes ? '\n' : '') + 
        `Payment ${data.paymentAmount} on ${data.paymentDate}: ${data.notes}`;
      sheet.getRange(row, this.COL.NOTES).setValue(newNotes);
    }

    const tenantName = sheet.getRange(row, this.COL.TENANT_NAME).getValue();
    const roomNumber = sheet.getRange(row, this.COL.ROOM_NUMBER).getValue();

    // Log payment in financial records
    FinancialManager.logPayment({
      date: new Date(data.paymentDate),
      type: 'Rent Income',
      description: `Rent payment from ${tenantName} - Room ${roomNumber}` + 
        (data.notes ? ` (${data.notes})` : ''),
      amount: data.paymentAmount,
      category: 'Rent',
      paymentMethod: data.paymentMethod,
      reference: data.referenceNumber || `RENT-${roomNumber}-${Utils.formatDate(new Date(data.paymentDate), 'yyyyMM')}`,
      tenant: tenantName
    });

    return { success: true };

  } catch (error) {
    Logger.log(`Error recording advanced payment: ${error.toString()}`);
    throw error;
  }
},

  recordTenantPayment: function(rowNumber, dateStr) {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      const row = parseInt(rowNumber);
      const rent = sheet.getRange(row, this.COL.NEGOTIATED_PRICE).getValue() || sheet.getRange(row, this.COL.RENTAL_PRICE).getValue();

      sheet.getRange(row, this.COL.LAST_PAYMENT).setValue(new Date(dateStr));
      sheet.getRange(row, this.COL.PAYMENT_STATUS).setValue(CONFIG.STATUS.PAYMENT.PAID);

      const tenantName = sheet.getRange(row, this.COL.TENANT_NAME).getValue();
      const roomNumber = sheet.getRange(row, this.COL.ROOM_NUMBER).getValue();

      FinancialManager.logPayment({
        date: new Date(dateStr),
        type: 'Rent Income',
        description: `Rent payment from ${tenantName} - Room ${roomNumber}`,
        amount: rent,
        category: 'Rent',
        tenant: tenantName,
        reference: `RENT-${roomNumber}-${Utils.formatDate(new Date(dateStr), 'yyyyMM')}`
      });

    } catch (error) {
      Logger.log(`Error recording payment: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
 * Process tenant move-in (improved version)
 */
processMoveIn: function() {
  try {
    const ui = SpreadsheetApp.getUi();

    const applications = SheetManager.getAllData(CONFIG.SHEETS.APPLICATIONS);
    const appHeaders = SheetManager.getHeaderMap(CONFIG.SHEETS.APPLICATIONS);
    const vacantRooms = this.getVacantRooms();

    if (vacantRooms.length === 0) {
      ui.alert('No vacant rooms available for move-in.');
      return;
    }

    const nameIdx = appHeaders['Full Name'] || this.APP_COL.FULL_NAME;
    const emailIdx = appHeaders['Email'] || this.APP_COL.EMAIL;
    const phoneIdx = appHeaders['Phone'] || this.APP_COL.PHONE;
    const moveInIdx = appHeaders['Desired Move-in Date'] || this.APP_COL.MOVE_IN_DATE;
    const roomIdx = appHeaders['Preferred Room'] || this.APP_COL.PREFERRED_ROOM;

    const appData = applications.length > 0 ? applications.map((row, i) => ({
      index: i,
      name: row[nameIdx - 1],
      email: row[emailIdx - 1],
      phone: row[phoneIdx - 1],
      moveIn: row[moveInIdx - 1],
      room: (row[roomIdx - 1] || '').replace(/Room\s*/, '')
    })).filter(a => a.name) : [];

    const appOptions = appData.length > 0 ? 
      appData.map((a, i) => `<option value="${i}">${a.name} - ${a.email}</option>`).join('') :
      '<option value="">No applications found</option>';

    const roomOptions = vacantRooms.map(room =>
      `<option value="${room.number}" data-rent="${room.rent}">Room ${room.number} - ${Utils.formatCurrency(room.rent)}/month</option>`
    ).join('');

    const html = HtmlService.createHtmlOutput(`
      <style>
        .move-in-form {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          min-height: 100vh;
          padding: 20px;
          box-sizing: border-box;
        }
        .form-container {
          max-width: 700px;
          margin: 0 auto;
          background: white;
          padding: 30px;
          border-radius: 15px;
          box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        .form-title {
          color: #2c3e50;
          text-align: center;
          margin-bottom: 30px;
          font-size: 26px;
          font-weight: 600;
        }
        .form-section {
          margin-bottom: 25px;
          padding: 20px;
          background: #f8f9fa;
          border-radius: 10px;
          border-left: 4px solid #667eea;
        }
        .section-title {
          color: #2c3e50;
          font-size: 16px;
          font-weight: 600;
          margin-bottom: 15px;
          display: flex;
          align-items: center;
        }
        .form-row {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 15px;
          margin-bottom: 15px;
        }
        .form-group {
          margin-bottom: 15px;
        }
        .form-group.full-width {
          grid-column: 1 / -1;
        }
        .form-label {
          display: block;
          margin-bottom: 6px;
          color: #34495e;
          font-weight: 500;
          font-size: 14px;
        }
        .required {
          color: #dc3545;
        }
        .form-control {
          width: 100%;
          padding: 12px 15px;
          border: 2px solid #e9ecef;
          border-radius: 8px;
          font-size: 14px;
          transition: all 0.3s ease;
          box-sizing: border-box;
        }
        .form-control:focus {
          outline: none;
          border-color: #667eea;
          box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
        }
        .btn {
          padding: 12px 25px;
          border: none;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.3s ease;
          margin-right: 10px;
        }
        .btn-primary {
          background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
          color: white;
        }
        .btn-primary:hover {
          transform: translateY(-2px);
          box-shadow: 0 5px 15px rgba(102, 126, 234, 0.4);
        }
        .btn-secondary {
          background: #6c757d;
          color: white;
        }
        .btn-secondary:hover {
          background: #5a6268;
        }
        .form-footer {
          text-align: center;
          margin-top: 30px;
          padding-top: 20px;
          border-top: 1px solid #e9ecef;
        }
        .validation-error {
          color: #dc3545;
          font-size: 12px;
          margin-top: 5px;
          display: none;
        }
        .app-info {
          background: #e3f2fd;
          padding: 15px;
          border-radius: 8px;
          margin-top: 10px;
          display: none;
        }
        .info-item {
          margin-bottom: 5px;
          font-size: 13px;
          color: #1976d2;
        }
        .room-info {
          background: #e8f5e8;
          padding: 10px;
          border-radius: 6px;
          margin-top: 8px;
          font-size: 13px;
          color: #2e7d32;
          display: none;
        }
      </style>

      <div class="move-in-form">
        <div class="form-container">
          <h2 class="form-title">🏠 Process Tenant Move-In</h2>
          
          <form id="moveInForm">
            <div class="form-section">
              <div class="section-title">📋 Application Selection</div>
              <div class="form-group">
                <label class="form-label" for="appSelect">
                  Select from Applications ${appData.length === 0 ? '(or enter manually below)' : ''}
                </label>
                <select id="appSelect" class="form-control" onchange="fillFromApplication()" ${appData.length === 0 ? 'disabled' : ''}>
                  <option value="">Select an application or enter manually</option>
                  ${appOptions}
                </select>
                <div id="appInfo" class="app-info"></div>
              </div>
            </div>

            <div class="form-section">
              <div class="section-title">🏠 Room Assignment</div>
              <div class="form-group">
                <label class="form-label" for="roomNumber">Available Room <span class="required">*</span></label>
                <select id="roomNumber" name="roomNumber" class="form-control" required onchange="updateRoomInfo()">
                  <option value="">Select a room...</option>
                  ${roomOptions}
                </select>
                <div id="roomInfo" class="room-info"></div>
                <div class="validation-error" id="roomError">Please select a room</div>
              </div>
            </div>

            <div class="form-section">
              <div class="section-title">👤 Tenant Information</div>
              <div class="form-row">
                <div class="form-group">
                  <label class="form-label" for="tenantName">Full Name <span class="required">*</span></label>
                  <input type="text" id="tenantName" name="tenantName" class="form-control" required>
                  <div class="validation-error" id="nameError">Please enter tenant's full name</div>
                </div>
                <div class="form-group">
                  <label class="form-label" for="email">Email Address <span class="required">*</span></label>
                  <input type="email" id="email" name="email" class="form-control" required>
                  <div class="validation-error" id="emailError">Please enter a valid email address</div>
                </div>
              </div>
              <div class="form-row">
                <div class="form-group">
                  <label class="form-label" for="phone">Phone Number <span class="required">*</span></label>
                  <input type="tel" id="phone" name="phone" class="form-control" required 
                         pattern="[0-9\-\(\)\+\s]+" placeholder="(555) 123-4567">
                  <div class="validation-error" id="phoneError">Please enter a valid phone number</div>
                </div>
                <div class="form-group">
                  <label class="form-label" for="moveInDate">Move-in Date <span class="required">*</span></label>
                  <input type="date" id="moveInDate" name="moveInDate" class="form-control" 
                         value="${Utils.formatDate(new Date(), 'yyyy-MM-dd')}" required>
                  <div class="validation-error" id="dateError">Please select a move-in date</div>
                </div>
              </div>
            </div>

            <div class="form-section">
              <div class="section-title">💰 Financial Information</div>
              <div class="form-row">
                <div class="form-group">
                  <label class="form-label" for="securityDeposit">Security Deposit <span class="required">*</span></label>
                  <input type="number" id="securityDeposit" name="securityDeposit" class="form-control" 
                         step="0.01" min="0" required placeholder="0.00">
                  <div class="validation-error" id="depositError">Please enter security deposit amount</div>
                </div>
                <div class="form-group">
                  <label class="form-label" for="negotiatedRent">Negotiated Rent (optional)</label>
                  <input type="number" id="negotiatedRent" name="negotiatedRent" class="form-control" 
                         step="0.01" min="0" placeholder="Leave blank to use standard rate">
                </div>
              </div>
            </div>

            <div class="form-section">
              <div class="section-title">📞 Emergency & Lease Information</div>
              <div class="form-group">
                <label class="form-label" for="emergencyContact">Emergency Contact</label>
                <input type="text" id="emergencyContact" name="emergencyContact" class="form-control" 
                       placeholder="Name, relationship, and phone number">
              </div>
              <div class="form-row">
                <div class="form-group">
                  <label class="form-label" for="leaseEndDate">Lease End Date</label>
                  <input type="date" id="leaseEndDate" name="leaseEndDate" class="form-control">
                </div>
                <div class="form-group">
                  <label class="form-label" for="depositMethod">Deposit Payment Method</label>
                  <select id="depositMethod" name="depositMethod" class="form-control">
                    <option value="Cash">Cash</option>
                    <option value="Check">Check</option>
                    <option value="Bank Transfer">Bank Transfer</option>
                    <option value="Money Order">Money Order</option>
                  </select>
                </div>
              </div>
              <div class="form-group">
                <label class="form-label" for="notes">Additional Notes</label>
                <textarea id="notes" name="notes" class="form-control" rows="3" 
                          placeholder="Any additional information about the tenant or move-in..."></textarea>
              </div>
            </div>

            <div class="form-footer">
              <button type="button" class="btn btn-primary" onclick="processMoveIn()">
                Complete Move-In
              </button>
              <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">
                Cancel
              </button>
            </div>
          </form>
        </div>
      </div>

      <script>
        const apps = ${JSON.stringify(appData)};
        
        function fillFromApplication() {
          const select = document.getElementById('appSelect');
          const info = document.getElementById('appInfo');
          
          if (select.value !== '') {
            const app = apps[parseInt(select.value)];
            document.getElementById('tenantName').value = app.name || '';
            document.getElementById('email').value = app.email || '';
            document.getElementById('phone').value = app.phone || '';
            
            if (app.moveIn) {
              try {
                const date = new Date(app.moveIn);
                if (!isNaN(date.getTime())) {
                  document.getElementById('moveInDate').value = date.toISOString().slice(0, 10);
                }
              } catch (e) {}
            }
            
            // Try to select preferred room
            const roomSelect = document.getElementById('roomNumber');
            for (let i = 0; i < roomSelect.options.length; i++) {
              if (roomSelect.options[i].value == app.room) {
                roomSelect.selectedIndex = i;
                updateRoomInfo();
                break;
              }
            }
            
            info.innerHTML = \`
              <div class="info-item"><strong>Application Info:</strong> \${app.name} (\${app.email})</div>
              <div class="info-item"><strong>Preferred Room:</strong> \${app.room || 'Any available'}</div>
            \`;
            info.style.display = 'block';
          } else {
            info.style.display = 'none';
          }
          clearValidationErrors();
        }
        
        function updateRoomInfo() {
          const select = document.getElementById('roomNumber');
          const info = document.getElementById('roomInfo');
          const depositField = document.getElementById('securityDeposit');
          
          if (select.value) {
            const option = select.options[select.selectedIndex];
            const rent = parseFloat(option.dataset.rent) || 0;
            
            info.innerHTML = \`Monthly rent: \${rent.toFixed(2)} | Suggested deposit: \${rent.toFixed(2)}\`;
            info.style.display = 'block';
            
            if (!depositField.value) {
              depositField.value = rent.toFixed(2);
            }
          } else {
            info.style.display = 'none';
          }
        }

        function validateForm() {
          let isValid = true;
          clearValidationErrors();

          const room = document.getElementById('roomNumber').value;
          const name = document.getElementById('tenantName').value.trim();
          const email = document.getElementById('email').value.trim();
          const phone = document.getElementById('phone').value.trim();
          const moveInDate = document.getElementById('moveInDate').value;
          const deposit = parseFloat(document.getElementById('securityDeposit').value);

          if (!room) {
            showError('roomError', 'Please select a room');
            isValid = false;
          }

          if (!name) {
            showError('nameError', 'Please enter tenant\\'s full name');
            isValid = false;
          }

          if (!email || !email.includes('@')) {
            showError('emailError', 'Please enter a valid email address');
            isValid = false;
          }

          if (!phone) {
            showError('phoneError', 'Please enter a phone number');
            isValid = false;
          }

          if (!moveInDate) {
            showError('dateError', 'Please select a move-in date');
            isValid = false;
          }

          if (!deposit || deposit <= 0) {
            showError('depositError', 'Please enter a valid security deposit amount');
            isValid = false;
          }

          return isValid;
        }

        function showError(elementId, message) {
          const errorElement = document.getElementById(elementId);
          if (errorElement) {
            errorElement.textContent = message;
            errorElement.style.display = 'block';
          }
        }

        function clearValidationErrors() {
          const errors = document.querySelectorAll('.validation-error');
          errors.forEach(error => error.style.display = 'none');
        }

        function processMoveIn() {
          if (!validateForm()) {
            return;
          }

          const formData = new FormData(document.getElementById('moveInForm'));
          const data = Object.fromEntries(formData);
          
          const submitBtn = document.querySelector('.btn-primary');
          submitBtn.disabled = true;
          submitBtn.textContent = 'Processing Move-In...';

          google.script.run
            .withSuccessHandler(function() {
              alert('Move-in processed successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('Error processing move-in: ' + error.message);
              submitBtn.disabled = false;
              submitBtn.textContent = 'Complete Move-In';
            })
            .completeMoveIn(data);
        }
      </script>
    `).setWidth(750).setHeight(800);

    ui.showModalDialog(html, 'Process Move-In');
      
  } catch (error) {
    handleSystemError(error, 'processMoveIn');
  }
},
  
  /**
   * Complete move-in process (called from HTML form)
   */
  completeMoveIn: function(data) {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      
      // Find the room row
      const roomRows = SheetManager.findRows(CONFIG.SHEETS.TENANTS, this.COL.ROOM_NUMBER, data.roomNumber);
      
      if (roomRows.length === 0) {
        throw new Error('Room not found');
      }
      
      const roomRow = roomRows[0];
      const rowNumber = roomRow.rowNumber;
      
      // Update tenant information
      const moveInDate = new Date(data.moveInDate);
      const leaseEndDate = data.leaseEndDate ? new Date(data.leaseEndDate) : '';
      
      sheet.getRange(rowNumber, this.COL.NEGOTIATED_PRICE).setValue(data.negotiatedRent || '');
      sheet.getRange(rowNumber, this.COL.TENANT_NAME).setValue(data.tenantName);
      sheet.getRange(rowNumber, this.COL.TENANT_EMAIL).setValue(data.email);
      sheet.getRange(rowNumber, this.COL.TENANT_PHONE).setValue(data.phone);
      sheet.getRange(rowNumber, this.COL.MOVE_IN_DATE).setValue(moveInDate);
      sheet.getRange(rowNumber, this.COL.SECURITY_DEPOSIT).setValue(parseFloat(data.securityDeposit));
      sheet.getRange(rowNumber, this.COL.ROOM_STATUS).setValue(CONFIG.STATUS.ROOM.OCCUPIED);
      sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue(CONFIG.STATUS.PAYMENT.DUE);
      sheet.getRange(rowNumber, this.COL.EMERGENCY_CONTACT).setValue(data.emergencyContact || '');
      sheet.getRange(rowNumber, this.COL.LEASE_END_DATE).setValue(leaseEndDate);
      sheet.getRange(rowNumber, this.COL.NOTES).setValue(data.notes || '');
      
      // Log security deposit payment
      if (data.securityDeposit > 0) {
        FinancialManager.logPayment({
          date: moveInDate,
          type: 'Security Deposit',
          description: `Security deposit from ${data.tenantName} - Room ${data.roomNumber}`,
          amount: parseFloat(data.securityDeposit),
          category: 'Deposit',
          tenant: data.tenantName,
          reference: `DEPOSIT-${data.roomNumber}-${Utils.formatDate(moveInDate, 'yyyyMMdd')}`
        });
      }
      
      // Send welcome email
      EmailManager.sendWelcomeEmail(data.email, {
        tenantName: data.tenantName,
        roomNumber: data.roomNumber,
        moveInDate: Utils.formatDate(moveInDate, 'MMMM dd, yyyy'),
        rent: data.negotiatedRent || roomRow.data[this.COL.RENTAL_PRICE - 1],
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME
      });
      
      return { success: true, message: 'Move-in processed successfully' };
      
    } catch (error) {
      Logger.log(`Error in completeMoveIn: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
 * Process tenant move-out (improved version - FIXED)
 */
processMoveOut: function() {
  try {
    const ui = SpreadsheetApp.getUi();
    const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);

    const tenants = [];
    data.forEach((row, i) => {
      const name = row[this.COL.TENANT_NAME - 1];
      const status = row[this.COL.ROOM_STATUS - 1];
      if (name && status === CONFIG.STATUS.ROOM.OCCUPIED) {
        tenants.push({
          index: i,
          rowNumber: i + 2,
          name: name,
          room: row[this.COL.ROOM_NUMBER - 1],
          email: row[this.COL.TENANT_EMAIL - 1],
          phone: row[this.COL.TENANT_PHONE - 1],
          deposit: row[this.COL.SECURITY_DEPOSIT - 1] || 0,
          rent: row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1],
          moveIn: row[this.COL.MOVE_IN_DATE - 1],
          emergencyContact: row[this.COL.EMERGENCY_CONTACT - 1]
        });
      }
    });

    if (tenants.length === 0) {
      ui.alert('No current tenants found for move-out.');
      return;
    }

    const tenantOptions = tenants.map((t, i) => 
      `<option value="${i}" data-deposit="${t.deposit}" data-name="${t.name}" data-room="${t.room}" data-email="${t.email}">
        ${t.name} (Room ${t.room}) - Deposit: ${Utils.formatCurrency(t.deposit)}
      </option>`
    ).join('');

    const html = HtmlService.createHtmlOutput(`
      <style>
        .move-out-form {
          font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
          background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
          min-height: 100vh;
          padding: 20px;
          box-sizing: border-box;
        }
        .form-container {
          max-width: 700px;
          margin: 0 auto;
          background: white;
          padding: 30px;
          border-radius: 15px;
          box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        .form-title {
          color: #2c3e50;
          text-align: center;
          margin-bottom: 30px;
          font-size: 26px;
          font-weight: 600;
        }
        .form-section {
          margin-bottom: 25px;
          padding: 20px;
          background: #f8f9fa;
          border-radius: 10px;
          border-left: 4px solid #ff6b6b;
        }
        .section-title {
          color: #2c3e50;
          font-size: 16px;
          font-weight: 600;
          margin-bottom: 15px;
          display: flex;
          align-items: center;
        }
        .form-row {
          display: grid;
          grid-template-columns: 1fr 1fr;
          gap: 15px;
          margin-bottom: 15px;
        }
        .form-group {
          margin-bottom: 15px;
        }
        .form-group.full-width {
          grid-column: 1 / -1;
        }
        .form-label {
          display: block;
          margin-bottom: 6px;
          color: #34495e;
          font-weight: 500;
          font-size: 14px;
        }
        .required {
          color: #dc3545;
        }
        .form-control {
          width: 100%;
          padding: 12px 15px;
          border: 2px solid #e9ecef;
          border-radius: 8px;
          font-size: 14px;
          transition: all 0.3s ease;
          box-sizing: border-box;
        }
        .form-control:focus {
          outline: none;
          border-color: #ff6b6b;
          box-shadow: 0 0 0 3px rgba(255, 107, 107, 0.1);
        }
        .btn {
          padding: 12px 25px;
          border: none;
          border-radius: 8px;
          font-size: 14px;
          font-weight: 600;
          cursor: pointer;
          transition: all 0.3s ease;
          margin-right: 10px;
        }
        .btn-danger {
          background: linear-gradient(135deg, #ff6b6b 0%, #ee5a24 100%);
          color: white;
        }
        .btn-danger:hover {
          transform: translateY(-2px);
          box-shadow: 0 5px 15px rgba(255, 107, 107, 0.4);
        }
        .btn-secondary {
          background: #6c757d;
          color: white;
        }
        .btn-secondary:hover {
          background: #5a6268;
        }
        .form-footer {
          text-align: center;
          margin-top: 30px;
          padding-top: 20px;
          border-top: 1px solid #e9ecef;
        }
        .validation-error {
          color: #dc3545;
          font-size: 12px;
          margin-top: 5px;
          display: none;
        }
        .tenant-info {
          background: #e3f2fd;
          padding: 15px;
          border-radius: 8px;
          margin-top: 10px;
          display: none;
        }
        .info-item {
          margin-bottom: 8px;
          font-size: 13px;
          color: #1976d2;
        }
        .deposit-summary {
          background: #e8f5e8;
          padding: 15px;
          border-radius: 8px;
          border-left: 4px solid #28a745;
        }
        .deposit-calculation {
          display: grid;
          grid-template-columns: 1fr auto;
          gap: 10px;
          margin-bottom: 8px;
          padding: 5px 0;
        }
        .total-line {
          border-top: 2px solid #28a745;
          padding-top: 8px;
          font-weight: bold;
          font-size: 16px;
          color: #155724;
        }
        .condition-grid {
          display: grid;
          grid-template-columns: repeat(3, 1fr);
          gap: 10px;
        }
        .condition-item {
          text-align: center;
          padding: 10px;
          border-radius: 6px;
          cursor: pointer;
          transition: all 0.3s ease;
          border: 2px solid #e9ecef;
        }
        .condition-item.selected {
          border-color: #ff6b6b;
          background: #fff5f5;
        }
        .condition-excellent {
          background: #d4edda;
          border-color: #28a745;
        }
        .condition-good {
          background: #fff3cd;
          border-color: #ffc107;
        }
        .condition-poor {
          background: #f8d7da;
          border-color: #dc3545;
        }
      </style>

      <div class="move-out-form">
        <div class="form-container">
          <h2 class="form-title">📤 Process Tenant Move-Out</h2>
          
          <form id="moveOutForm">
            <div class="form-section">
              <div class="section-title">👤 Tenant Selection</div>
              <div class="form-group">
                <label class="form-label" for="tenantSelect">Select Tenant <span class="required">*</span></label>
                <select id="tenantSelect" class="form-control" required onchange="updateTenantInfo()">
                  <option value="">Choose a tenant...</option>
                  ${tenantOptions}
                </select>
                <div class="validation-error" id="tenantError">Please select a tenant</div>
                <div id="tenantInfo" class="tenant-info"></div>
              </div>
            </div>

            <div class="form-section">
              <div class="section-title">📅 Move-Out Details</div>
              <div class="form-row">
                <div class="form-group">
                  <label class="form-label" for="moveOutDate">Move-Out Date <span class="required">*</span></label>
                  <input type="date" id="moveOutDate" name="moveOutDate" class="form-control" 
                         value="${Utils.formatDate(new Date(), 'yyyy-MM-dd')}" required>
                  <div class="validation-error" id="dateError">Please select a move-out date</div>
                </div>
                <div class="form-group">
                  <label class="form-label" for="keyReturnMethod">Key Return Method</label>
                  <select id="keyReturnMethod" name="keyReturnMethod" class="form-control">
                    <option value="Hand delivery">Hand delivery to manager</option>
                    <option value="Drop box">Secure drop box</option>
                    <option value="Mail slot">Mail slot</option>
                    <option value="Pickup arranged">Pickup arranged</option>
                  </select>
                </div>
              </div>
              <div class="form-group">
                <label class="form-label" for="forwardingAddress">Forwarding Address <span class="required">*</span></label>
                <textarea id="forwardingAddress" name="forwardingAddress" class="form-control" rows="3" 
                          placeholder="Complete mailing address for security deposit refund" required></textarea>
                <div class="validation-error" id="addressError">Please provide a forwarding address</div>
              </div>
            </div>

            <div class="form-section">
              <div class="section-title">🏠 Room Condition Assessment</div>
              <div class="form-group">
                <label class="form-label">Overall Room Condition</label>
                <div class="condition-grid">
                  <div class="condition-item condition-excellent" onclick="selectCondition('Excellent')">
                    <div style="font-size: 20px;">✨</div>
                    <div>Excellent</div>
                    <small>No damage</small>
                  </div>
                  <div class="condition-item condition-good" onclick="selectCondition('Good')">
                    <div style="font-size: 20px;">👍</div>
                    <div>Good</div>
                    <small>Minor wear</small>
                  </div>
                  <div class="condition-item condition-poor" onclick="selectCondition('Poor')">
                    <div style="font-size: 20px;">⚠️</div>
                    <div>Poor</div>
                    <small>Needs repairs</small>
                  </div>
                </div>
                <input type="hidden" id="roomCondition" name="roomCondition" value="Good">
              </div>
              <div class="form-group">
                <label class="form-label" for="conditionNotes">Condition Notes</label>
                <textarea id="conditionNotes" name="conditionNotes" class="form-control" rows="3" 
                          placeholder="Details about room condition, any damages, etc."></textarea>
              </div>
            </div>

            <div class="form-section">
              <div class="section-title">💰 Security Deposit Calculation</div>
              <div class="form-row">
                <div class="form-group">
                  <label class="form-label" for="cleaningFee">Cleaning Fee</label>
                  <input type="number" id="cleaningFee" class="form-control" step="0.01" min="0" 
                         value="0" onchange="calculateRefund()">
                </div>
                <div class="form-group">
                  <label class="form-label" for="repairCosts">Repair Costs</label>
                  <input type="number" id="repairCosts" class="form-control" step="0.01" min="0" 
                         value="0" onchange="calculateRefund()">
                </div>
              </div>
              <div class="form-row">
                <div class="form-group">
                  <label class="form-label" for="otherDeductions">Other Deductions</label>
                  <input type="number" id="otherDeductions" class="form-control" step="0.01" min="0" 
                         value="0" onchange="calculateRefund()">
                </div>
                <div class="form-group">
                  <label class="form-label" for="deductionReason">Deduction Reason</label>
                  <input type="text" id="deductionReason" name="deductionReason" class="form-control" 
                         placeholder="Reason for deductions">
                </div>
              </div>
              
              <div id="depositSummary" class="deposit-summary">
                <div class="deposit-calculation">
                  <span>Original Security Deposit:</span>
                  <span id="originalDeposit">$0.00</span>
                </div>
                <div class="deposit-calculation">
                  <span>Less: Cleaning Fee:</span>
                  <span id="cleaningDisplay">$0.00</span>
                </div>
                <div class="deposit-calculation">
                  <span>Less: Repair Costs:</span>
                  <span id="repairDisplay">$0.00</span>
                </div>
                <div class="deposit-calculation">
                  <span>Less: Other Deductions:</span>
                  <span id="otherDisplay">$0.00</span>
                </div>
                <div class="deposit-calculation total-line">
                  <span>Refund Amount:</span>
                  <span id="refundAmount">$0.00</span>
                </div>
              </div>
              <input type="hidden" id="totalDeductions" name="totalDeductions" value="0">
              <input type="hidden" id="depositRefund" name="depositRefund" value="0">
            </div>

            <div class="form-section">
              <div class="section-title">📝 Additional Information</div>
              <div class="form-group">
                <label class="form-label" for="moveOutReason">Reason for Moving</label>
                <select id="moveOutReason" name="moveOutReason" class="form-control">
                  <option value="">Select reason...</option>
                  <option value="Job relocation">Job relocation</option>
                  <option value="Buying a home">Buying a home</option>
                  <option value="Moving closer to family">Moving closer to family</option>
                  <option value="Found different housing">Found different housing</option>
                  <option value="Financial reasons">Financial reasons</option>
                  <option value="Dissatisfied with property">Dissatisfied with property</option>
                  <option value="Lease completion">Lease completion</option>
                  <option value="Other">Other</option>
                </select>
              </div>
              <div class="form-group">
                <label class="form-label" for="finalNotes">Final Notes</label>
                <textarea id="finalNotes" name="finalNotes" class="form-control" rows="3" 
                          placeholder="Any additional notes about the move-out process..."></textarea>
              </div>
            </div>

            <div class="form-footer">
              <button type="button" class="btn btn-danger" onclick="processMoveOut()">
                Complete Move-Out
              </button>
              <button type="button" class="btn btn-secondary" onclick="google.script.host.close()">
                Cancel
              </button>
            </div>
          </form>
        </div>
      </div>

      <script>
        const tenants = ${JSON.stringify(tenants)};
        let selectedCondition = 'Good';
        
        function updateTenantInfo() {
          const select = document.getElementById('tenantSelect');
          const info = document.getElementById('tenantInfo');
          
          if (select.value !== '') {
            const tenant = tenants[parseInt(select.value)];
            document.getElementById('originalDeposit').textContent = '$' + tenant.deposit.toFixed(2);
            
            info.innerHTML = 
              '<div class="info-item"><strong>Tenant:</strong> ' + tenant.name + '</div>' +
              '<div class="info-item"><strong>Room:</strong> ' + tenant.room + '</div>' +
              '<div class="info-item"><strong>Email:</strong> ' + tenant.email + '</div>' +
              '<div class="info-item"><strong>Security Deposit:</strong> $' + tenant.deposit.toFixed(2) + '</div>' +
              '<div class="info-item"><strong>Move-in Date:</strong> ' + (tenant.moveIn ? new Date(tenant.moveIn).toLocaleDateString() : 'Not specified') + '</div>';
            
            info.style.display = 'block';
            calculateRefund();
          } else {
            info.style.display = 'none';
            document.getElementById('originalDeposit').textContent = '$0.00';
            calculateRefund();
          }
          clearValidationErrors();
        }
        
        function selectCondition(condition) {
          selectedCondition = condition;
          document.getElementById('roomCondition').value = condition;
          
          // Update visual selection
          document.querySelectorAll('.condition-item').forEach(function(item) {
            item.classList.remove('selected');
          });
          event.target.closest('.condition-item').classList.add('selected');
          
          // Auto-suggest fees based on condition
          const cleaningFee = document.getElementById('cleaningFee');
          const repairCosts = document.getElementById('repairCosts');
          
          switch(condition) {
            case 'Excellent':
              if (!cleaningFee.value || cleaningFee.value == '0') cleaningFee.value = '50';
              if (!repairCosts.value || repairCosts.value == '0') repairCosts.value = '0';
              break;
            case 'Good':
              if (!cleaningFee.value || cleaningFee.value == '0') cleaningFee.value = '75';
              if (!repairCosts.value || repairCosts.value == '0') repairCosts.value = '0';
              break;
            case 'Poor':
              if (!cleaningFee.value || cleaningFee.value == '0') cleaningFee.value = '100';
              if (!repairCosts.value || repairCosts.value == '0') repairCosts.value = '150';
              break;
          }
          calculateRefund();
        }
        
        function calculateRefund() {
          const originalDeposit = parseFloat(document.getElementById('originalDeposit').textContent.replace('$', '')) || 0;
          const cleaningFee = parseFloat(document.getElementById('cleaningFee').value) || 0;
          const repairCosts = parseFloat(document.getElementById('repairCosts').value) || 0;
          const otherDeductions = parseFloat(document.getElementById('otherDeductions').value) || 0;
          
          const totalDeductions = cleaningFee + repairCosts + otherDeductions;
          const refundAmount = Math.max(0, originalDeposit - totalDeductions);
          
          document.getElementById('cleaningDisplay').textContent = '$' + cleaningFee.toFixed(2);
          document.getElementById('repairDisplay').textContent = '$' + repairCosts.toFixed(2);
          document.getElementById('otherDisplay').textContent = '$' + otherDeductions.toFixed(2);
          document.getElementById('refundAmount').textContent = '$' + refundAmount.toFixed(2);
          
          document.getElementById('totalDeductions').value = totalDeductions.toFixed(2);
          document.getElementById('depositRefund').value = refundAmount.toFixed(2);
          
          // Update refund amount color
          const refundElement = document.getElementById('refundAmount');
          if (refundAmount === originalDeposit) {
            refundElement.style.color = '#155724'; // Green - full refund
          } else if (refundAmount === 0) {
            refundElement.style.color = '#721c24'; // Red - no refund
          } else {
            refundElement.style.color = '#856404'; // Yellow - partial refund
          }
        }

        function validateForm() {
          let isValid = true;
          clearValidationErrors();

          const tenant = document.getElementById('tenantSelect').value;
          const moveOutDate = document.getElementById('moveOutDate').value;
          const forwardingAddress = document.getElementById('forwardingAddress').value.trim();

          if (!tenant) {
            showError('tenantError', 'Please select a tenant');
            isValid = false;
          }

          if (!moveOutDate) {
            showError('dateError', 'Please select a move-out date');
            isValid = false;
          }

          if (!forwardingAddress) {
            showError('addressError', 'Please provide a forwarding address');
            isValid = false;
          }

          return isValid;
        }

        function showError(elementId, message) {
          const errorElement = document.getElementById(elementId);
          if (errorElement) {
            errorElement.textContent = message;
            errorElement.style.display = 'block';
          }
        }

        function clearValidationErrors() {
          const errors = document.querySelectorAll('.validation-error');
          errors.forEach(function(error) {
            error.style.display = 'none';
          });
        }

        function processMoveOut() {
          if (!validateForm()) {
            return;
          }

          const tenantIndex = parseInt(document.getElementById('tenantSelect').value);
          const tenant = tenants[tenantIndex];
          
          const formData = {
            rowNumber: tenant.rowNumber,
            tenantName: tenant.name,
            roomNumber: tenant.room,
            email: tenant.email,
            moveOutDate: document.getElementById('moveOutDate').value,
            forwardingAddress: document.getElementById('forwardingAddress').value,
            roomCondition: document.getElementById('roomCondition').value,
            conditionNotes: document.getElementById('conditionNotes').value,
            keyReturnMethod: document.getElementById('keyReturnMethod').value,
            cleaningFee: parseFloat(document.getElementById('cleaningFee').value) || 0,
            repairCosts: parseFloat(document.getElementById('repairCosts').value) || 0,
            otherDeductions: parseFloat(document.getElementById('otherDeductions').value) || 0,
            deductionReason: document.getElementById('deductionReason').value,
            totalDeductions: parseFloat(document.getElementById('totalDeductions').value) || 0,
            depositRefund: parseFloat(document.getElementById('depositRefund').value) || 0,
            securityDeposit: tenant.deposit,
            moveOutReason: document.getElementById('moveOutReason').value,
            finalNotes: document.getElementById('finalNotes').value
          };
          
          const submitBtn = document.querySelector('.btn-danger');
          submitBtn.disabled = true;
          submitBtn.textContent = 'Processing Move-Out...';

          google.script.run
            .withSuccessHandler(function() {
              alert('Move-out processed successfully!');
              google.script.host.close();
            })
            .withFailureHandler(function(error) {
              alert('Error processing move-out: ' + error.message);
              submitBtn.disabled = false;
              submitBtn.textContent = 'Complete Move-Out';
            })
            .completeMoveOutAdvanced(formData);
        }
        
        // Initialize condition selection
        setTimeout(function() {
          selectCondition('Good');
        }, 100);
      </script>
    `).setWidth(750).setHeight(900);

    ui.showModalDialog(html, 'Process Move-Out');

  } catch (error) {
    handleSystemError(error, 'processMoveOut');
  }
},
  
  /**
   * Complete move-out process (called from HTML form)
   */
  completeMoveOut: function(data) {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      const rowNumber = parseInt(data.rowNumber);
      const moveOutDate = new Date(data.moveOutDate);
      const deductions = parseFloat(data.deductions) || 0;
      const depositRefund = parseFloat(data.securityDeposit) - deductions;
      
      // Update tenant record
      sheet.getRange(rowNumber, this.COL.ROOM_STATUS).setValue(CONFIG.STATUS.ROOM.VACANT);
      sheet.getRange(rowNumber, this.COL.MOVE_OUT_PLANNED).setValue(moveOutDate);
      sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue('');
      sheet.getRange(rowNumber, this.COL.NOTES).setValue(
        `MOVED OUT: ${Utils.formatDate(moveOutDate)} - ${data.finalNotes || ''}`
      );
      
      // Clear tenant information but keep historical data in notes
      const historicalNote = `Former tenant: ${data.tenantName} (${Utils.formatDate(moveOutDate)})`;
      sheet.getRange(rowNumber, this.COL.TENANT_NAME).setValue('');
      sheet.getRange(rowNumber, this.COL.TENANT_EMAIL).setValue('');
      sheet.getRange(rowNumber, this.COL.TENANT_PHONE).setValue('');
      sheet.getRange(rowNumber, this.COL.LAST_PAYMENT).setValue('');
      sheet.getRange(rowNumber, this.COL.EMERGENCY_CONTACT).setValue('');
      sheet.getRange(rowNumber, this.COL.LEASE_END_DATE).setValue('');
      
      // Log security deposit refund
      if (depositRefund > 0) {
        FinancialManager.logPayment({
          date: moveOutDate,
          type: 'Security Deposit Refund',
          description: `Security deposit refund to ${data.tenantName} - Room ${data.roomNumber}`,
          amount: -depositRefund, // Negative because it's money going out
          category: 'Deposit Refund',
          tenant: data.tenantName,
          reference: `REFUND-${data.roomNumber}-${Utils.formatDate(moveOutDate, 'yyyyMMdd')}`
        });
      }
      
      // Log any deductions
      if (deductions > 0) {
        FinancialManager.logPayment({
          date: moveOutDate,
          type: 'Deposit Deduction',
          description: `Deposit deduction - ${data.tenantName} - ${data.deductionReason}`,
          amount: deductions,
          category: 'Maintenance',
          tenant: data.tenantName,
          reference: `DEDUCTION-${data.roomNumber}-${Utils.formatDate(moveOutDate, 'yyyyMMdd')}`
        });
      }
      
      // Send move-out confirmation email
      EmailManager.sendMoveOutConfirmation(sheet.getRange(rowNumber, this.COL.TENANT_EMAIL).getValue(), {
        tenantName: data.tenantName,
        roomNumber: data.roomNumber,
        moveOutDate: Utils.formatDate(moveOutDate, 'MMMM dd, yyyy'),
        depositRefund: depositRefund,
        deductions: deductions,
        forwardingAddress: data.forwardingAddress
      });
      
      // Create move-out document
      DocumentManager.createMoveOutReport({
        tenantName: data.tenantName,
        roomNumber: data.roomNumber,
        moveOutDate: moveOutDate,
        roomCondition: data.roomCondition,
        deductions: deductions,
        depositRefund: depositRefund,
        deductionReason: data.deductionReason,
        finalNotes: data.finalNotes,
        forwardingAddress: data.forwardingAddress
      });
      
      return { success: true, message: 'Move-out processed successfully' };
      
    } catch (error) {
      Logger.log(`Error in completeMoveOut: ${error.toString()}`);
      throw error;
    }
  },

  /**
   * Simplified move-out processing used by dropdown panel
   */
  simpleMoveOut: function(data) {
    try {
      const sheet = SheetManager.getSheet(CONFIG.SHEETS.TENANTS);
      const rowNumber = parseInt(data.rowNumber);
      const moveOutDate = new Date(data.moveOutDate);

      sheet.getRange(rowNumber, this.COL.ROOM_STATUS).setValue(CONFIG.STATUS.ROOM.VACANT);
      sheet.getRange(rowNumber, this.COL.MOVE_OUT_PLANNED).setValue(moveOutDate);
      sheet.getRange(rowNumber, this.COL.PAYMENT_STATUS).setValue('');
      sheet.getRange(rowNumber, this.COL.NOTES).setValue(`MOVED OUT: ${Utils.formatDate(moveOutDate)}`);
      sheet.getRange(rowNumber, this.COL.TENANT_NAME).setValue('');
      sheet.getRange(rowNumber, this.COL.TENANT_EMAIL).setValue('');
      sheet.getRange(rowNumber, this.COL.TENANT_PHONE).setValue('');
      sheet.getRange(rowNumber, this.COL.LAST_PAYMENT).setValue('');
      sheet.getRange(rowNumber, this.COL.EMERGENCY_CONTACT).setValue('');
      sheet.getRange(rowNumber, this.COL.LEASE_END_DATE).setValue('');

      return { success: true };

    } catch (error) {
      Logger.log(`Error in simpleMoveOut: ${error.toString()}`);
      throw error;
    }
  },
  
  /**
 * Show tenant dashboard (improved version)
 */
showTenantDashboard: function() {
  try {
    const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
    const stats = this.calculateTenantStats(data);
    
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 20px; background: #f8f9fa;">
        <h2 style="color: #2c3e50; margin-bottom: 25px; text-align: center;">🏠 Tenant Dashboard</h2>
        
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin-bottom: 30px;">
          <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); padding: 20px; border-radius: 12px; text-align: center; color: white; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
            <h3 style="margin: 0; font-size: 14px; opacity: 0.9;">Total Rooms</h3>
            <p style="font-size: 32px; margin: 10px 0; font-weight: bold;">${stats.totalRooms}</p>
            <small style="opacity: 0.8;">Property Units</small>
          </div>
          <div style="background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%); padding: 20px; border-radius: 12px; text-align: center; color: white; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
            <h3 style="margin: 0; font-size: 14px; opacity: 0.9;">Occupied</h3>
            <p style="font-size: 32px; margin: 10px 0; font-weight: bold;">${stats.occupiedRooms}</p>
            <small style="opacity: 0.8;">${stats.occupancyRate}% Occupancy</small>
          </div>
          <div style="background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%); padding: 20px; border-radius: 12px; text-align: center; color: white; box-shadow: 0 4px 8px rgba(0,0,0,0.1);">
            <h3 style="margin: 0; font-size: 14px; opacity: 0.9;">Vacant</h3>
            <p style="font-size: 32px; margin: 10px 0; font-weight: bold;">${stats.vacantRooms}</p>
            <small style="opacity: 0.8;">Available Units</small>
          </div>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 25px;">
          <h3 style="color: #2c3e50; margin-bottom: 20px; display: flex; align-items: center;">
            💰 Payment Status Overview
          </h3>
          <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px;">
            <div style="text-align: center; padding: 15px; background: #d4edda; border-radius: 8px; border-left: 4px solid #28a745;">
              <h4 style="margin: 0; color: #155724;">Payments Paid</h4>
              <p style="font-size: 28px; margin: 10px 0; font-weight: bold; color: #155724;">${stats.paidCount}</p>
              <small style="color: #155724;">Up to date</small>
            </div>
            <div style="text-align: center; padding: 15px; background: #fff3cd; border-radius: 8px; border-left: 4px solid #ffc107;">
              <h4 style="margin: 0; color: #856404;">Payments Due</h4>
              <p style="font-size: 28px; margin: 10px 0; font-weight: bold; color: #856404;">${stats.dueCount}</p>
              <small style="color: #856404;">Due this month</small>
            </div>
            <div style="text-align: center; padding: 15px; background: #f8d7da; border-radius: 8px; border-left: 4px solid #dc3545;">
              <h4 style="margin: 0; color: #721c24;">Overdue</h4>
              <p style="font-size: 28px; margin: 10px 0; font-weight: bold; color: #721c24;">${stats.overdueCount}</p>
              <small style="color: #721c24;">Needs attention</small>
            </div>
          </div>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 25px;">
          <h3 style="color: #2c3e50; margin-bottom: 15px;">📊 Occupancy Rate</h3>
          <div style="background: #e9ecef; height: 25px; border-radius: 12px; overflow: hidden; position: relative;">
            <div style="background: linear-gradient(90deg, #28a745, #20c997); height: 100%; width: ${stats.occupancyRate}%; transition: width 0.8s ease-in-out; border-radius: 12px;"></div>
            <span style="position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%); font-weight: bold; color: white; text-shadow: 1px 1px 2px rgba(0,0,0,0.5);">${stats.occupancyRate}%</span>
          </div>
        </div>
        
        <div style="background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); margin-bottom: 25px;">
          <h3 style="color: #2c3e50; margin-bottom: 20px;">💰 Revenue Analysis</h3>
          <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
            <div style="padding: 15px; background: #e8f6f3; border-radius: 8px; border-left: 4px solid #17a2b8;">
              <h4 style="margin: 0 0 10px 0; color: #0c5460;">Current Monthly Income</h4>
              <p style="font-size: 24px; margin: 0; font-weight: bold; color: #0c5460;">${Utils.formatCurrency(stats.currentRevenue)}</p>
            </div>
            <div style="padding: 15px; background: #fef9e7; border-radius: 8px; border-left: 4px solid #ffc107;">
              <h4 style="margin: 0 0 10px 0; color: #856404;">Potential Monthly Income</h4>
              <p style="font-size: 24px; margin: 0; font-weight: bold; color: #856404;">${Utils.formatCurrency(stats.potentialRevenue)}</p>
            </div>
          </div>
          ${stats.potentialRevenue > stats.currentRevenue ? `
            <div style="margin-top: 15px; padding: 15px; background: #fff3cd; border-radius: 8px; border-left: 4px solid #ffc107;">
              <p style="margin: 0; color: #856404;"><strong>Revenue Loss from Vacancies:</strong> ${Utils.formatCurrency(stats.potentialRevenue - stats.currentRevenue)}</p>
            </div>
          ` : ''}
        </div>
        
        ${stats.overdueList.length > 0 ? `
          <div style="background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1);">
            <h3 style="color: #dc3545; margin-bottom: 20px; display: flex; align-items: center;">
              ⚠️ Overdue Tenants Requiring Attention
            </h3>
            <div style="background: #f8d7da; padding: 20px; border-radius: 8px; border-left: 4px solid #dc3545;">
              ${stats.overdueList.map(tenant => 
                `<div style="margin-bottom: 10px; padding: 10px; background: white; border-radius: 6px;">
                  <strong style="color: #721c24;">${tenant.name}</strong> 
                  <span style="color: #6c757d;">(Room ${tenant.room})</span><br>
                  <small style="color: #856404;">Last payment: ${tenant.lastPayment}</small>
                </div>`
              ).join('')}
            </div>
          </div>
        ` : `
          <div style="background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); text-align: center;">
            <h3 style="color: #28a745; margin-bottom: 15px;">✅ All Payments Current</h3>
            <p style="color: #6c757d; margin: 0;">No overdue payments at this time.</p>
          </div>
        `}
      </div>
    `)
      .setWidth(900)
      .setHeight(700);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Tenant Dashboard');
    
  } catch (error) {
    handleSystemError(error, 'showTenantDashboard');
  }
},
  /**
   * Calculate tenant statistics
   */
  calculateTenantStats: function(data) {
    const stats = {
      totalRooms: 0,
      occupiedRooms: 0,
      vacantRooms: 0,
      maintenanceRooms: 0,
      paidCount: 0,
      dueCount: 0,
      overdueCount: 0,
      currentRevenue: 0,
      potentialRevenue: 0,
      overdueList: []
    };
    
    data.forEach(row => {
      if (!row[this.COL.ROOM_NUMBER - 1]) return; // Skip empty rows
      
      stats.totalRooms++;
      
      const roomStatus = row[this.COL.ROOM_STATUS - 1];
      const paymentStatus = row[this.COL.PAYMENT_STATUS - 1];
      const rent = row[this.COL.NEGOTIATED_PRICE - 1] || row[this.COL.RENTAL_PRICE - 1] || 0;
      
      // Room status counts
      switch (roomStatus) {
        case CONFIG.STATUS.ROOM.OCCUPIED:
          stats.occupiedRooms++;
          stats.currentRevenue += rent;
          break;
        case CONFIG.STATUS.ROOM.VACANT:
          stats.vacantRooms++;
          break;
        case CONFIG.STATUS.ROOM.MAINTENANCE:
          stats.maintenanceRooms++;
          break;
      }
      
      // Potential revenue (all rooms)
      stats.potentialRevenue += rent;
      
      // Payment status counts
      switch (paymentStatus) {
        case CONFIG.STATUS.PAYMENT.PAID:
          stats.paidCount++;
          break;
        case CONFIG.STATUS.PAYMENT.DUE:
          stats.dueCount++;
          break;
        case CONFIG.STATUS.PAYMENT.OVERDUE:
          stats.overdueCount++;
          const lastPayment = row[this.COL.LAST_PAYMENT - 1];
          stats.overdueList.push({
            name: row[this.COL.TENANT_NAME - 1],
            room: row[this.COL.ROOM_NUMBER - 1],
            lastPayment: lastPayment ? Utils.formatDate(lastPayment) : 'Never'
          });
          break;
      }
    });
    
    stats.occupancyRate = stats.totalRooms > 0 ? 
      Math.round((stats.occupiedRooms / stats.totalRooms) * 100) : 0;
    
    return stats;
  },
  
  /**
   * Get vacant rooms
   */
  getVacantRooms: function() {
    const data = SheetManager.getAllData(CONFIG.SHEETS.TENANTS);
    const vacantRooms = [];
    
    data.forEach(row => {
      if (row[this.COL.ROOM_STATUS - 1] === CONFIG.STATUS.ROOM.VACANT) {
        vacantRooms.push({
          number: row[this.COL.ROOM_NUMBER - 1],
          rent: row[this.COL.RENTAL_PRICE - 1]
        });
      }
    });
    
    return vacantRooms;
  },
  
  /**
   * Calculate rent due date
   */
  calculateRentDueDate: function() {
    const today = new Date();
    const dueDate = new Date(today.getFullYear(), today.getMonth(), 5); // 5th of current month
    
    if (today.getDate() > 5) {
      // If past the 5th, return next month's due date
      dueDate.setMonth(dueDate.getMonth() + 1);
    }
    
    return Utils.formatDate(dueDate, 'MMMM dd, yyyy');
  },
  
  /**
   * Get tenant by room number
   */
  getTenantByRoom: function(roomNumber) {
    const tenantRows = SheetManager.findRows(CONFIG.SHEETS.TENANTS, this.COL.ROOM_NUMBER, roomNumber);
    return tenantRows.length > 0 ? tenantRows[0] : null;
  },
  
  /**
   * Get tenant by email
   */
  getTenantByEmail: function(email) {
    const tenantRows = SheetManager.findRows(CONFIG.SHEETS.TENANTS, this.COL.TENANT_EMAIL, email);
    return tenantRows.length > 0 ? tenantRows[0] : null;
  }
};

Logger.log('TenantManager module loaded successfully');
