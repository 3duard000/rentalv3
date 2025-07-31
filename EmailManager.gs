// EmailManager.gs - Enhanced Email Communication System
// Handles all automated email communications with improved professional layouts

const EmailManager = {
  
  /**
   * Enhanced email template configurations with professional HTML layouts
   */
  TEMPLATES: {
    RENT_REMINDER: {
      subject: 'Rent Payment Reminder - {monthYear} | {propertyName}',
      template: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Rent Payment Reminder</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background-color: #f8f9fa; }
    .container { max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
    .header h1 { margin: 0; font-size: 28px; font-weight: 300; }
    .header p { margin: 10px 0 0 0; opacity: 0.9; }
    .content { padding: 30px; }
    .alert-box { background: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0; border-radius: 4px; }
    .alert-box.overdue { background: #f8d7da; border-left-color: #dc3545; }
    .payment-details { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .payment-row { display: flex; justify-content: space-between; margin-bottom: 10px; }
    .payment-row:last-child { margin-bottom: 0; font-weight: bold; border-top: 1px solid #dee2e6; padding-top: 10px; }
    .payment-methods { background: #e3f2fd; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .method-item { margin-bottom: 10px; }
    .footer { background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #dee2e6; font-size: 14px; color: #6c757d; }
    .button { display: inline-block; background: #667eea; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: 600; margin: 10px 0; }
    .status-badge { display: inline-block; padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: bold; text-transform: uppercase; }
    .status-due { background: #fff3cd; color: #856404; }
    .status-overdue { background: #f8d7da; color: #721c24; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>💰 Rent Payment Reminder</h1>
      <p>{propertyName}</p>
    </div>
    
    <div class="content">
      <p>Dear {tenantName},</p>
      
      <p>This is a friendly reminder regarding your rent payment for Room {roomNumber}.</p>
      
      <div class="{statusClass}">
        <strong>Payment Status:</strong> 
        <span class="status-badge {statusBadgeClass}">{status}</span>
        {urgencyMessage}
      </div>
      
      <div class="payment-details">
        <h3 style="margin-top: 0; color: #495057;">💳 Payment Details - {monthYear}</h3>
        <div class="payment-row">
          <span>Room Number:</span>
          <span><strong>{roomNumber}</strong></span>
        </div>
        <div class="payment-row">
          <span>Monthly Rent:</span>
          <span><strong>{rent}</strong></span>
        </div>
        <div class="payment-row">
          <span>Due Date:</span>
          <span><strong>{dueDate}</strong></span>
        </div>
        {lateFeeSection}
        <div class="payment-row">
          <span>Total Amount Due:</span>
          <span style="color: #dc3545; font-size: 18px;"><strong>{totalDue}</strong></span>
        </div>
      </div>
      
      <div class="payment-methods">
        <h4 style="margin-top: 0; color: #495057;">💡 Payment Methods Available</h4>
        <div class="method-item">🏦 <strong>Bank Transfer:</strong> Contact us for account details</div>
        <div class="method-item">💳 <strong>Online Payment:</strong> Use our secure payment portal</div>
        <div class="method-item">📝 <strong>Check:</strong> Made payable to "{propertyName}"</div>
        <div class="method-item">💵 <strong>Cash:</strong> Deliver to management office with receipt</div>
      </div>
      
      <div style="text-align: center; margin: 30px 0;">
        <a href="mailto:{managerEmail}?subject=Payment Question - Room {roomNumber}" class="button">
          Contact Property Manager
        </a>
      </div>
      
      <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745;">
        <strong>📋 Already Paid?</strong><br>
        If you have already made your payment, please disregard this message and contact us so we can update our records.
      </div>
      
      <p style="margin-top: 30px;">Thank you for your prompt attention to this matter and for being a valued resident of our community.</p>
      
      <p>Best regards,<br>
      <strong>{propertyName} Management Team</strong></p>
    </div>
    
    <div class="footer">
      <p><strong>Contact Information</strong><br>
      📧 Email: {managerEmail}<br>
      🏠 Property: {propertyName}</p>
      
      <p style="margin-top: 15px; font-size: 12px;">
        This is an automated message from your property management system.
      </p>
    </div>
  </div>
</body>
</html>`
    },
    
    LATE_PAYMENT_ALERT: {
      subject: '🚨 Late Payment Alert - {count} Tenant(s) Overdue | {propertyName}',
      template: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Late Payment Alert</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background-color: #f8f9fa; }
    .container { max-width: 700px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #dc3545 0%, #fd7e14 100%); color: white; padding: 30px; text-align: center; }
    .header h1 { margin: 0; font-size: 28px; font-weight: 300; }
    .header .alert-count { font-size: 48px; font-weight: bold; margin: 10px 0; }
    .content { padding: 30px; }
    .summary-card { background: #f8d7da; border-left: 4px solid #dc3545; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .tenant-list { margin: 20px 0; }
    .tenant-item { background: #fff3cd; padding: 15px; border-radius: 8px; margin-bottom: 15px; border-left: 4px solid #ffc107; }
    .tenant-details { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 15px; margin-top: 10px; }
    .detail-item { font-size: 14px; }
    .detail-label { color: #6c757d; font-weight: 600; }
    .actions-section { background: #e3f2fd; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .action-item { margin-bottom: 10px; padding-left: 20px; position: relative; }
    .action-item:before { content: "→"; position: absolute; left: 0; color: #667eea; font-weight: bold; }
    .footer { background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #dee2e6; font-size: 14px; color: #6c757d; }
    .priority-high { border-left-color: #dc3545; background: #f8d7da; }
    .stats-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; margin: 20px 0; }
    .stat-item { text-align: center; padding: 15px; background: #f8f9fa; border-radius: 8px; }
    .stat-number { font-size: 24px; font-weight: bold; color: #dc3545; }
    .stat-label { font-size: 12px; color: #6c757d; text-transform: uppercase; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>🚨 Late Payment Alert</h1>
      <div class="alert-count">{count}</div>
      <p>Tenant(s) with overdue payments require immediate attention</p>
    </div>
    
    <div class="content">
      <div class="summary-card">
        <h3 style="margin-top: 0; color: #721c24;">⚠️ Urgent Action Required</h3>
        <p><strong>{count} tenant(s)</strong> currently have overdue rent payments that require immediate follow-up.</p>
        <p><strong>Generated:</strong> {currentDate}</p>
      </div>
      
      <div class="stats-grid">
        <div class="stat-item">
          <div class="stat-number">{count}</div>
          <div class="stat-label">Overdue Tenants</div>
        </div>
        <div class="stat-item">
          <div class="stat-number">HIGH</div>
          <div class="stat-label">Priority Level</div>
        </div>
        <div class="stat-item">
          <div class="stat-number">24H</div>
          <div class="stat-label">Response Time</div>
        </div>
      </div>
      
      <h3>📋 Overdue Tenant Details</h3>
      <div class="tenant-list">
        {overdueListHtml}
      </div>
      
      <div class="actions-section">
        <h4 style="margin-top: 0; color: #495057;">🎯 Recommended Actions</h4>
        <div class="action-item">Contact each tenant personally via phone call within 24 hours</div>
        <div class="action-item">Send formal late payment notices with updated balances</div>
        <div class="action-item">Review payment plan options for tenants experiencing difficulties</div>
        <div class="action-item">Consider escalation procedures for significantly overdue accounts</div>
        <div class="action-item">Update tenant payment status in the management system</div>
        <div class="action-item">Document all communication attempts and responses</div>
      </div>
      
      <div style="background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107; margin: 20px 0;">
        <strong>💡 Pro Tip:</strong> Early intervention often prevents small payment delays from becoming major collection issues. 
        Reach out with understanding and offer flexible solutions when appropriate.
      </div>
      
      <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; border-left: 4px solid #28a745;">
        <strong>📊 System Information:</strong><br>
        This automated alert was generated by the {propertyName} Management System.<br>
        For the most current information, please check the tenant dashboard.
      </div>
    </div>
    
    <div class="footer">
      <p><strong>Property Management System</strong><br>
      🏠 {propertyName}<br>
      📧 Manager: {managerEmail}</p>
      
      <p style="margin-top: 15px; font-size: 12px;">
        Alert generated automatically on {currentDate}
      </p>
    </div>
  </div>
</body>
</html>`
    },
    
    MONTHLY_INVOICE: {
      subject: '📄 Monthly Rent Invoice - {monthYear} | {propertyName}',
      template: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Monthly Rent Invoice</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background-color: #f8f9fa; }
    .container { max-width: 600px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #28a745 0%, #20c997 100%); color: white; padding: 30px; text-align: center; }
    .header h1 { margin: 0; font-size: 28px; font-weight: 300; }
    .content { padding: 30px; }
    .invoice-details { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; border: 2px solid #e9ecef; }
    .invoice-row { display: flex; justify-content: space-between; margin-bottom: 8px; }
    .invoice-total { border-top: 2px solid #dee2e6; padding-top: 10px; font-weight: bold; font-size: 18px; color: #28a745; }
    .payment-info { background: #e3f2fd; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .footer { background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #dee2e6; font-size: 14px; color: #6c757d; }
    .button { display: inline-block; background: #28a745; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: 600; margin: 10px 0; }
    .highlight-box { background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107; margin: 20px 0; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📄 Monthly Invoice</h1>
      <p>{propertyName}</p>
      <p style="font-size: 18px; margin: 10px 0 0 0;">{monthYear}</p>
    </div>
    
    <div class="content">
      <p>Dear {tenantName},</p>
      
      <p>Please find your monthly rent invoice for {monthYear} below. Your continued residency and prompt payment help us maintain our wonderful community.</p>
      
      <div class="invoice-details">
        <h3 style="margin-top: 0; color: #495057;">🧾 Invoice Details</h3>
        <div class="invoice-row">
          <span>Tenant Name:</span>
          <span><strong>{tenantName}</strong></span>
        </div>
        <div class="invoice-row">
          <span>Room Number:</span>
          <span><strong>{roomNumber}</strong></span>
        </div>
        <div class="invoice-row">
          <span>Billing Period:</span>
          <span><strong>{monthYear}</strong></span>
        </div>
        <div class="invoice-row">
          <span>Due Date:</span>
          <span><strong>{dueDate}</strong></span>
        </div>
        <div class="invoice-row invoice-total">
          <span>Monthly Rent:</span>
          <span>{rent}</span>
        </div>
      </div>
      
      <div class="payment-info">
        <h4 style="margin-top: 0; color: #495057;">💳 Payment Information</h4>
        <p><strong>Payment is due by the 5th of each month.</strong></p>
        <p>Late fees may apply after {lateFeeDay} days past the due date.</p>
        
        <div style="margin-top: 15px;">
          <strong>Payment Methods:</strong>
          <ul style="margin: 10px 0;">
            <li>🏦 Bank Transfer (contact for details)</li>
            <li>💳 Online Payment Portal</li>
            <li>📝 Check payable to "{propertyName}"</li>
            <li>💵 Cash (office hours only)</li>
          </ul>
        </div>
      </div>
      
      <div style="text-align: center; margin: 30px 0;">
        <a href="mailto:{managerEmail}?subject=Invoice Question - {monthYear} - Room {roomNumber}" class="button">
          Questions About This Invoice?
        </a>
      </div>
      
      <div class="highlight-box">
        <strong>📋 Attached Documents:</strong><br>
        Your detailed PDF invoice is attached to this email for your records.
      </div>
      
      <p>Thank you for being a valued member of our community. If you have any questions about this invoice or need to discuss payment arrangements, please don't hesitate to contact us.</p>
      
      <p>Best regards,<br>
      <strong>{propertyName} Management</strong></p>
    </div>
    
    <div class="footer">
      <p><strong>Contact Information</strong><br>
      📧 Email: {managerEmail}<br>
      🏠 Property: {propertyName}</p>
    </div>
  </div>
</body>
</html>`
    },

    ENHANCED_WELCOME_EMAIL: {
      subject: '🎉 Welcome to {propertyName}! Your New Home Awaits',
      template: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Welcome to Your New Home</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background-color: #f8f9fa; }
    .container { max-width: 700px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 40px; text-align: center; }
    .header h1 { margin: 0; font-size: 32px; font-weight: 300; }
    .welcome-banner { font-size: 64px; margin: 20px 0; }
    .content { padding: 30px; }
    .info-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 20px 0; }
    .info-card { background: #f8f9fa; padding: 20px; border-radius: 8px; border-left: 4px solid #667eea; }
    .info-card h4 { margin-top: 0; color: #495057; }
    .highlight-section { background: #e8f5e8; padding: 20px; border-radius: 8px; border-left: 4px solid #28a745; margin: 20px 0; }
    .rules-section { background: #fff3cd; padding: 20px; border-radius: 8px; border-left: 4px solid #ffc107; margin: 20px 0; }
    .contact-section { background: #e3f2fd; padding: 20px; border-radius: 8px; border-left: 4px solid #2196f3; margin: 20px 0; }
    .checklist { list-style: none; padding: 0; }
    .checklist li { padding: 8px 0; }
    .checklist li:before { content: "✅"; margin-right: 10px; }
    .footer { background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #dee2e6; font-size: 14px; color: #6c757d; }
    .button { display: inline-block; background: #667eea; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; font-weight: 600; margin: 10px 5px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <div class="welcome-banner">🏠</div>
      <h1>Welcome to {propertyName}!</h1>
      <p style="font-size: 18px; margin: 15px 0 0 0; opacity: 0.9;">
        We're thrilled to have you as part of our community
      </p>
    </div>
    
    <div class="content">
      <p>Dear {tenantName},</p>
      
      <p>Welcome to your new home! We're delighted to have you join the {propertyName} community. This email contains everything you need to know to get settled in and make the most of your stay with us.</p>
      
      <div class="info-grid">
        <div class="info-card">
          <h4>🏠 Your Accommodation Details</h4>
          <p><strong>Room:</strong> {roomNumber}<br>
          <strong>Move-in Date:</strong> {moveInDate}<br>
          <strong>Monthly Rent:</strong> {rent}<br>
          <strong>Security Deposit:</strong> {securityDeposit}<br>
          <strong>Lease Type:</strong> {leaseEndDate}</p>
        </div>
        
        <div class="info-card">
          <h4>📋 Important Dates & Info</h4>
          <p><strong>Rent Due:</strong> 5th of each month<br>
          <strong>Late Fee:</strong> After {lateFeeDay} days<br>
          <strong>Emergency Contact:</strong> {emergencyContact}<br>
          <strong>Office Hours:</strong> Mon-Fri 9AM-6PM</p>
        </div>
      </div>
      
      <div class="highlight-section">
        <h3 style="margin-top: 0;">🎯 Getting Started Checklist</h3>
        <ul class="checklist">
          <li>Save our contact information in your phone</li>
          <li>Connect to WiFi (details in your welcome packet)</li>
          <li>Review house rules and community guidelines</li>
          <li>Set up your preferred payment method</li>
          <li>Introduce yourself to your neighbors</li>
          <li>Familiarize yourself with common areas</li>
        </ul>
      </div>
      
      <div class="rules-section">
        <h3 style="margin-top: 0;">📜 House Rules & Guidelines</h3>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 15px;">
          <div>
            <h4>⏰ Quiet Hours</h4>
            <p>10:00 PM - 7:00 AM daily</p>
            
            <h4>🧹 Common Areas</h4>
            <p>Keep clean and tidy after use</p>
          </div>
          <div>
            <h4>🚭 No Smoking</h4>
            <p>Inside the building (outdoor areas OK)</p>
            
            <h4>👥 Guests</h4>
            <p>Welcome with prior notice to management</p>
          </div>
        </div>
      </div>
      
      <div class="contact-section">
        <h3 style="margin-top: 0;">📞 Contact Information</h3>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
          <div>
            <h4>Property Management</h4>
            <p>📧 Email: {managerEmail}<br>
            📱 Emergency: 24/7 available<br>
            🏢 Office: On-site management</p>
          </div>
          <div>
            <h4>Important Services</h4>
            <p>🚑 Emergency: 911<br>
            ⚡ Utilities: Contact office<br>
            📦 Packages: Secure delivery area</p>
          </div>
        </div>
      </div>
      
      <div style="background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0;">
        <h3 style="margin-top: 0; color: #495057;">🌟 Community Amenities</h3>
        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 15px;">
          <div style="text-align: center;">
            <div style="font-size: 24px; margin-bottom: 5px;">📶</div>
            <strong>Free WiFi</strong><br>
            <small>High-speed internet throughout</small>
          </div>
          <div style="text-align: center;">
            <div style="font-size: 24px; margin-bottom: 5px;">🧺</div>
            <strong>Laundry</strong><br>
            <small>Shared laundry facilities</small>
          </div>
          <div style="text-align: center;">
            <div style="font-size: 24px; margin-bottom: 5px;">🍳</div>
            <strong>Kitchen</strong><br>
            <small>Fully equipped common kitchen</small>
          </div>
        </div>
      </div>
      
      <div style="text-align: center; margin: 30px 0;">
        <a href="mailto:{managerEmail}?subject=Welcome Question - Room {roomNumber}" class="button">
          Ask a Question
        </a>
        <a href="mailto:{managerEmail}?subject=Maintenance Request - Room {roomNumber}" class="button">
          Report an Issue
        </a>
      </div>
      
      <p>We're here to help make your stay comfortable and enjoyable. Please don't hesitate to reach out if you have any questions, concerns, or suggestions. Welcome home!</p>
      
      <p>Warmest regards,<br>
      <strong>The {propertyName} Team</strong></p>
    </div>
    
    <div class="footer">
      <p><strong>{propertyName}</strong><br>
      📧 {managerEmail} | 🏠 Your Home Away From Home</p>
      
      <p style="margin-top: 15px; font-size: 12px;">
        This welcome email was sent on {moveInDate}
      </p>
    </div>
  </div>
</body>
</html>`
    },

    ENHANCED_MOVE_OUT_CONFIRMATION: {
      subject: '📤 Move-Out Confirmation & Deposit Information | {propertyName}',
      template: `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Move-Out Confirmation</title>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 0; background-color: #f8f9fa; }
    .container { max-width: 700px; margin: 0 auto; background: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { background: linear-gradient(135deg, #dc3545 0%, #fd7e14 100%); color: white; padding: 30px; text-align: center; }
    .header h1 { margin: 0; font-size: 28px; font-weight: 300; }
    .content { padding: 30px; }
    .deposit-summary { background: #e3f2fd; padding: 25px; border-radius: 8px; margin: 20px 0; border: 2px solid #2196f3; }
    .deposit-calculation { display: grid; grid-template-columns: 1fr auto 1fr auto 1fr; gap: 15px; align-items: center; margin: 20px 0; text-align: center; }
    .calculation-amount { background: #f8f9fa; padding: 15px; border-radius: 8px; }
    .calculation-operator { font-size: 24px; font-weight: bold; color: #495057; }
    .original-deposit { border-left: 4px solid #28a745; }
    .deductions { border-left: 4px solid #dc3545; }
    .refund-amount { border-left: 4px solid #17a2b8; font-weight: bold; }
    .details-section { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; }
    .deduction-details { background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107; margin: 15px 0; }
    .timeline { background: #e8f5e8; padding: 20px; border-radius: 8px; border-left: 4px solid #28a745; margin: 20px 0; }
    .footer { background: #f8f9fa; padding: 20px; text-align: center; border-top: 1px solid #dee2e6; font-size: 14px; color: #6c757d; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>📤 Move-Out Confirmation</h1>
      <p>{propertyName}</p>
    </div>
    
    <div class="content">
      <p>Dear {tenantName},</p>
      
      <p>This confirms your successful move-out from Room {roomNumber} on {moveOutDate}. Thank you for being a valued resident of {propertyName}.</p>
      
      <div class="details-section">
        <h3 style="margin-top: 0; color: #495057;">📋 Move-Out Summary</h3>
        <div style="display: grid; grid-template-columns: 1fr 1fr; gap: 20px;">
          <div>
            <p><strong>Tenant:</strong> {tenantName}<br>
            <strong>Room:</strong> {roomNumber}<br>
            <strong>Move-Out Date:</strong> {moveOutDate}</p>
          </div>
          <div>
            <p><strong>Room Condition:</strong> {roomCondition}<br>
            <strong>Final Inspection:</strong> Completed<br>
            <strong>Keys Returned:</strong> ✅ Received</p>
          </div>
        </div>
      </div>
      
      <div class="deposit-summary">
        <h3 style="margin-top: 0; color: #1976d2;">💰 Security Deposit Summary</h3>
        
        <div class="deposit-calculation">
          <div class="calculation-amount original-deposit">
            <div style="color: #6c757d; font-size: 14px;">Original Deposit</div>
            <div style="font-size: 20px; font-weight: bold; color: #28a745;">{originalDeposit}</div>
          </div>
          
          <div class="calculation-operator">−</div>
          
          <div class="calculation-amount deductions">
            <div style="color: #6c757d; font-size: 14px;">Deductions</div>
            <div style="font-size: 20px; font-weight: bold; color: #dc3545;">{deductions}</div>
          </div>
          
          <div class="calculation-operator">=</div>
          
          <div class="calculation-amount refund-amount">
            <div style="color: #6c757d; font-size: 14px;">Refund Amount</div>
            <div style="font-size: 24px; font-weight: bold; color: #17a2b8;">{depositRefund}</div>
          </div>
        </div>
        
        {deductionDetailsSection}
      </div>
      
      <div class="timeline">
        <h3 style="margin-top: 0; color: #495057;">📅 Refund Processing Timeline</h3>
        <p><strong>Your deposit refund will be processed as follows:</strong></p>
        <ul>
          <li>✅ <strong>Move-out completed:</strong> {moveOutDate}</li>
          <li>📋 <strong>Final inspection:</strong> Completed</li>
          <li>💳 <strong>Refund processing:</strong> 5-7 business days</li>
          <li>📬 <strong>Delivery method:</strong> Check mailed to forwarding address</li>
        </ul>
      </div>
      
      <div style="background: #fff3cd; padding: 15px; border-radius: 8px; border-left: 4px solid #ffc107; margin: 20px 0;">
        <h4 style="margin-top: 0;">📬 Forwarding Address</h4>
        <p><strong>Your refund check will be mailed to:</strong><br>
        {forwardingAddress}</p>
        <p><small>Please allow 7-10 business days for delivery. If you need to update this address, contact us immediately.</small></p>
      </div>
      
      {finalNotesSection}
      
      <div style="background: #e8f5e8; padding: 20px; border-radius: 8px; border-left: 4px solid #28a745; margin: 20px 0;">
        <h3 style="margin-top: 0;">💝 Thank You!</h3>
        <p>We truly appreciate the time you spent as part of our community. Your residency helped make {propertyName} a wonderful place to live.</p>
        <p>We wish you all the best in your new home and future endeavors!</p>
      </div>
      
      <p>If you have any questions about your deposit refund or need any documentation for your records, please don't hesitate to contact us.</p>
      
      <p>Best wishes,<br>
      <strong>{propertyName} Management Team</strong></p>
    </div>
    
    <div class="footer">
      <p><strong>Contact Information</strong><br>
      📧 Email: {managerEmail}<br>
      🏠 Property: {propertyName}</p>
      
      <p style="margin-top: 15px; font-size: 12px;">
        Move-out confirmation sent on {moveOutDate}
      </p>
    </div>
  </div>
</body>
</html>`
    }
  },
/**
   * Enhanced rent reminder with professional HTML layout
   */
  sendRentReminder: function(email, data) {
    try {
      const template = this.TEMPLATES.RENT_REMINDER;
      
      // Determine status styling and messaging
      const isOverdue = data.status === CONFIG.STATUS.PAYMENT.OVERDUE;
      const statusClass = isOverdue ? 'alert-box overdue' : 'alert-box';
      const statusBadgeClass = isOverdue ? 'status-overdue' : 'status-due';
      
      const urgencyMessage = isOverdue ? 
        '<br><strong>⚠️ URGENT:</strong> This payment is now overdue. Please contact us immediately to avoid additional fees.' : 
        '<br>Your prompt payment helps us maintain our community standards.';
      
      // Calculate total due including late fees
      const lateFee = data.lateFee || 0;
      const rentAmount = parseFloat(data.rent) || 0;
      const totalDue = rentAmount + lateFee;
      
      // Late fee section
      const lateFeeSection = lateFee > 0 ? `
        <div class="payment-row">
          <span>Late Fee:</span>
          <span style="color: #dc3545;"><strong>${Utils.formatCurrency(lateFee)}</strong></span>
        </div>` : '';
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        statusClass: statusClass,
        statusBadgeClass: statusBadgeClass,
        urgencyMessage: urgencyMessage,
        lateFeeSection: lateFeeSection,
        totalDue: Utils.formatCurrency(totalDue),
        rent: Utils.formatCurrency(rentAmount),
        lateFeeDay: CONFIG.SYSTEM.LATE_FEE_DAYS
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body
      });
      
      Logger.log(`Enhanced rent reminder sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send enhanced rent reminder to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Enhanced late payment alert with improved design
   */
  sendLatePaymentAlert: function(managerEmail, data) {
    try {
      const template = this.TEMPLATES.LATE_PAYMENT_ALERT;
      
      // Generate HTML for overdue tenant list
      const overdueListHtml = data.overdueList.map((tenant, index) => `
        <div class="tenant-item ${index === 0 ? 'priority-high' : ''}">
          <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px;">
            <h4 style="margin: 0; color: #856404;">${tenant.tenant}</h4>
            <span style="background: #dc3545; color: white; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: bold;">
              ${index === 0 ? 'PRIORITY' : 'OVERDUE'}
            </span>
          </div>
          <div class="tenant-details">
            <div class="detail-item">
              <div class="detail-label">Room</div>
              <div style="font-weight: bold;">${tenant.room}</div>
            </div>
            <div class="detail-item">
              <div class="detail-label">Email</div>
              <div>${tenant.email}</div>
            </div>
            <div class="detail-item">
              <div class="detail-label">Last Payment</div>
              <div style="color: #dc3545; font-weight: bold;">${tenant.lastPayment}</div>
            </div>
          </div>
        </div>
      `).join('');
      
      const emailData = {
        count: data.count,
        overdueListHtml: overdueListHtml,
        currentDate: Utils.formatDate(new Date(), 'MMMM dd, yyyy'),
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail({
        to: managerEmail,
        subject: subject,
        htmlBody: body
      });
      
      Logger.log(`Enhanced late payment alert sent to manager: ${managerEmail}`);
      
    } catch (error) {
      Logger.log(`Failed to send enhanced late payment alert: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Enhanced monthly invoice with professional layout
   */
  sendEnhancedMonthlyInvoice: function(email, data, pdfAttachment) {
    try {
      const template = this.TEMPLATES.MONTHLY_INVOICE;
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        rent: Utils.formatCurrency(data.rent),
        dueDate: this.calculateRentDueDate(),
        lateFeeDay: CONFIG.SYSTEM.LATE_FEE_DAYS
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      const mailOptions = {
        to: email,
        subject: subject,
        htmlBody: body
      };
      
      if (pdfAttachment) {
        mailOptions.attachments = [pdfAttachment];
      }
      
      MailApp.sendEmail(mailOptions);
      Logger.log(`Enhanced monthly invoice sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send enhanced monthly invoice to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Enhanced welcome email for new tenants
   */
  sendEnhancedWelcomeEmail: function(email, data) {
    try {
      const template = this.TEMPLATES.ENHANCED_WELCOME_EMAIL;
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        lateFeeDay: CONFIG.SYSTEM.LATE_FEE_DAYS,
        rent: Utils.formatCurrency(data.rent),
        securityDeposit: Utils.formatCurrency(data.securityDeposit),
        leaseEndDate: data.leaseEndDate || 'Month-to-month lease'
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body
      });
      
      Logger.log(`Enhanced welcome email sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send enhanced welcome email to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Enhanced move-out confirmation email
   */
  sendEnhancedMoveOutConfirmation: function(email, data) {
    try {
      const template = this.TEMPLATES.ENHANCED_MOVE_OUT_CONFIRMATION;
      
      // Format monetary amounts
      const originalDeposit = parseFloat(data.depositRefund) + parseFloat(data.deductions || 0);
      
      // Generate deduction details section if there are deductions
      const deductionDetailsSection = data.deductions > 0 ? `
        <div class="deduction-details">
          <h4 style="margin-top: 0;">📋 Deduction Details</h4>
          <p><strong>Amount Deducted:</strong> ${Utils.formatCurrency(data.deductions)}</p>
          <p><strong>Reason:</strong> ${data.deductionReason || 'See detailed statement for itemized deductions'}</p>
        </div>
      ` : `
        <div style="background: #e8f5e8; padding: 15px; border-radius: 8px; margin: 15px 0;">
          <strong>✅ No Deductions Applied</strong><br>
          Your security deposit is being refunded in full.
        </div>
      `;
      
      // Generate final notes section if provided
      const finalNotesSection = data.finalNotes ? `
        <div style="background: #f8f9fa; padding: 15px; border-radius: 8px; margin: 20px 0;">
          <h4 style="margin-top: 0;">📝 Final Notes</h4>
          <p>${data.finalNotes}</p>
        </div>
      ` : '';
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        originalDeposit: Utils.formatCurrency(originalDeposit),
        deductions: Utils.formatCurrency(data.deductions || 0),
        depositRefund: Utils.formatCurrency(data.depositRefund),
        deductionDetailsSection: deductionDetailsSection,
        finalNotesSection: finalNotesSection
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail({
        to: email,
        subject: subject,
        htmlBody: body
      });
      
      Logger.log(`Enhanced move-out confirmation sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send enhanced move-out confirmation to ${email}: ${error.message}`);
      throw error;
    }
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

  // Keep original functions for backwards compatibility
  sendMonthlyInvoice: function(email, data, pdfAttachment) {
    return this.sendEnhancedMonthlyInvoice(email, data, pdfAttachment);
  },

  sendWelcomeEmail: function(email, data) {
    return this.sendEnhancedWelcomeEmail(email, data);
  },

  sendMoveOutConfirmation: function(email, data) {
    return this.sendEnhancedMoveOutConfirmation(email, data);
  },
  
  /**
   * Guest booking confirmation (unchanged - keeping original)
   */
  sendGuestBookingConfirmation: function(email, data) {
    try {
      const template = this.TEMPLATES.GUEST_BOOKING_CONFIRMATION;
      
      const specialRequests = data.specialRequests ? 
        `\nSpecial Requests Noted:\n${data.specialRequests}\n` : '';
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        totalAmount: Utils.formatCurrency(data.totalAmount),
        specialRequests: specialRequests
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail(email, subject, body);
      Logger.log(`Guest booking confirmation sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send guest booking confirmation to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Guest welcome email (unchanged)
   */
  sendGuestWelcome: function(email, data) {
    try {
      const template = this.TEMPLATES.GUEST_WELCOME;
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail(email, subject, body);
      Logger.log(`Guest welcome email sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send guest welcome email to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Guest checkout confirmation (unchanged)
   */
  sendGuestCheckoutConfirmation: function(email, data) {
    try {
      const template = this.TEMPLATES.GUEST_CHECKOUT_CONFIRMATION;
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL,
        totalAmount: Utils.formatCurrency(data.totalAmount),
        amountPaid: Utils.formatCurrency(data.amountPaid)
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail(email, subject, body);
      Logger.log(`Guest checkout confirmation sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send guest checkout confirmation to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Guest check-in reminder (unchanged)
   */
  sendGuestCheckInReminder: function(email, data) {
    try {
      const template = this.TEMPLATES.GUEST_CHECK_IN_REMINDER;
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail(email, subject, body);
      Logger.log(`Guest check-in reminder sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send guest check-in reminder to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Maintenance request confirmation (unchanged)
   */
  sendMaintenanceConfirmation: function(email, data) {
    try {
      const template = this.TEMPLATES.MAINTENANCE_REQUEST_CONFIRMATION;
      
      const emailData = {
        ...data,
        propertyName: CONFIG.SYSTEM.PROPERTY_NAME,
        managerEmail: CONFIG.SYSTEM.MANAGER_EMAIL
      };
      
      const subject = this.processTemplate(template.subject, emailData);
      const body = this.processTemplate(template.template, emailData);
      
      MailApp.sendEmail(email, subject, body);
      Logger.log(`Maintenance confirmation sent to: ${email}`);
      
    } catch (error) {
      Logger.log(`Failed to send maintenance confirmation to ${email}: ${error.message}`);
      throw error;
    }
  },
  
  /**
   * Process email template with data substitution
   */
  processTemplate: function(template, data) {
    let processed = template;
    
    Object.keys(data).forEach(key => {
      const placeholder = `{${key}}`;
      const value = data[key] || '';
      processed = processed.replace(new RegExp(placeholder, 'g'), value);
    });
    
    return processed;
  },
  
  /**
   * Send test email with enhanced layout
   */
  sendTestEmail: function() {
    try {
      const testEmail = CONFIG.SYSTEM.MANAGER_EMAIL;
      const subject = `✅ Test Email - ${CONFIG.SYSTEM.PROPERTY_NAME} Management System`;
      
      const htmlBody = `
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; line-height: 1.6; color: #333; margin: 0; padding: 20px; background-color: #f8f9fa; }
    .container { max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; padding: 30px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); }
    .header { text-align: center; margin-bottom: 30px; }
    .success { background: #e8f5e8; padding: 20px; border-radius: 8px; border-left: 4px solid #28a745; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 style="color: #28a745;">✅ Email System Test</h1>
      <p style="color: #6c757d;">System functionality verification</p>
    </div>
    
    <div class="success">
      <h3 style="margin-top: 0;">🎉 Email System Working Correctly!</h3>
      <p>If you received this email, your enhanced email system is functioning properly.</p>
    </div>
    
    <h3>📊 System Information</h3>
    <ul>
      <li><strong>Property:</strong> ${CONFIG.SYSTEM.PROPERTY_NAME}</li>
      <li><strong>Timestamp:</strong> ${new Date().toLocaleString()}</li>
      <li><strong>System Version:</strong> 2.0 Enhanced</li>
      <li><strong>Email Templates:</strong> Professional HTML Layout</li>
    </ul>
    
    <h3>✨ New Features Included</h3>
    <ul>
      <li>Professional HTML email layouts</li>
      <li>Enhanced rent reminder design</li>
      <li>Improved late payment alerts</li>
      <li>Better invoice presentation</li>
      <li>Enhanced welcome emails</li>
      <li>Professional move-out confirmations</li>
    </ul>
    
    <div style="background: #e3f2fd; padding: 15px; border-radius: 8px; margin: 20px 0;">
      <strong>📧 Email Template Status:</strong> All templates updated with responsive HTML design
    </div>
    
    <p style="text-align: center; margin-top: 30px;">
      <strong>System Administrator</strong><br>
      ${CONFIG.SYSTEM.PROPERTY_NAME} Management System
    </p>
  </div>
</body>
</html>`;
      
      MailApp.sendEmail({
        to: testEmail,
        subject: subject,
        htmlBody: htmlBody
      });
      
      SpreadsheetApp.getUi().alert(
        'Enhanced Test Email Sent',
        `A test email with the new professional layout has been sent to ${testEmail}. Please check your inbox.`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      
    } catch (error) {
      SpreadsheetApp.getUi().alert(
        'Email Error',
        `Failed to send test email: ${error.message}`,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  },
  
  /**
   * Enhanced email template configuration
   */
  configureEmailTemplates: function() {
    const html = HtmlService.createHtmlOutput(`
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background: #f8f9fa; margin: 0; }
        .container { max-width: 800px; margin: 0 auto; background: white; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.1); overflow: hidden; }
        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }
        .content { padding: 30px; line-height: 1.6; }
        .template-section { background: #f8f9fa; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #667eea; }
        .feature-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; margin: 20px 0; }
        .feature-card { background: #e3f2fd; padding: 15px; border-radius: 8px; }
        .btn { background: #28a745; color: white; padding: 12px 24px; border: none; border-radius: 6px; cursor: pointer; font-weight: 600; }
      </style>
      
      <div class="container">
        <div class="header">
          <h2>📧 Enhanced Email Template Configuration</h2>
          <p>Professional HTML email templates with responsive design</p>
        </div>
        
        <div class="content">
          <div style="background: #e8f5e8; padding: 20px; border-radius: 8px; margin-bottom: 30px; border-left: 4px solid #28a745;">
            <h3 style="margin-top: 0;">✨ Template Enhancements</h3>
            <p>All email templates have been upgraded with:</p>
            <ul>
              <li>Professional HTML layouts with CSS styling</li>
              <li>Responsive design for mobile devices</li>
              <li>Brand-consistent color schemes</li>
              <li>Improved readability and visual hierarchy</li>
              <li>Interactive elements and call-to-action buttons</li>
            </ul>
          </div>
          
          <div class="template-section">
            <h3>🏠 Tenant Communication Templates</h3>
            <div class="feature-grid">
              <div class="feature-card">
                <h4>💰 Rent Reminder</h4>
                <p>Professional payment reminders with payment details, late fee calculations, and multiple payment options.</p>
              </div>
              <div class="feature-card">
                <h4>🚨 Late Payment Alert</h4>
                <p>Manager notifications with tenant details, overdue summaries, and recommended actions.</p>
              </div>
              <div class="feature-card">
                <h4>📄 Monthly Invoice</h4>
                <p>Branded invoice emails with payment information and PDF attachments.</p>
              </div>
              <div class="feature-card">
                <h4>🎉 Welcome Email</h4>
                <p>Comprehensive onboarding emails with community information and amenities.</p>
              </div>
            </div>
          </div>
          
          <div class="template-section">
            <h3>📤 Move-Out Communications</h3>
            <div class="feature-card">
              <h4>📋 Move-Out Confirmation</h4>
              <p>Detailed deposit calculations, refund timelines, and professional farewell messaging.</p>
            </div>
          </div>
          
          <h3>📝 Template Customization Guide</h3>
          <ol>
            <li><strong>Access Templates:</strong> Open EmailManager.gs in the script editor</li>
            <li><strong>Find TEMPLATES Object:</strong> Locate the specific template you want to modify</li>
            <li><strong>Edit Content:</strong> Modify the HTML content while preserving variable placeholders</li>
            <li><strong>Variables:</strong> Use {variableName} format for dynamic content</li>
            <li><strong>Styling:</strong> Modify CSS within &lt;style&gt; tags for visual changes</li>
            <li><strong>Test Changes:</strong> Use the test email function to verify modifications</li>
          </ol>
          
          <div style="background: #fff3cd; padding: 20px; border-radius: 8px; margin: 20px 0; border-left: 4px solid #ffc107;">
            <h4>📋 Available Template Variables</h4>
            <div class="feature-grid">
              <div>
                <p><strong>Universal:</strong></p>
                <ul>
                  <li>{propertyName}</li>
                  <li>{managerEmail}</li>
                  <li>{currentDate}</li>
                </ul>
              </div>
              <div>
                <p><strong>Tenant Specific:</strong></p>
                <ul>
                  <li>{tenantName}</li>
                  <li>{roomNumber}</li>
                  <li>{rent}</li>
                  <li>{monthYear}</li>
                  <li>{dueDate}</li>
                </ul>
              </div>
            </div>
          </div>
          
          <div style="text-align: center; margin-top: 30px;">
            <button onclick="google.script.run.sendTestEmail()" class="btn">
              📧 Send Test Email
            </button>
          </div>
        </div>
      </div>
    `)
      .setWidth(900)
      .setHeight(700);
    
    SpreadsheetApp.getUi().showModalDialog(html, 'Enhanced Email Template Configuration');
  },
  
  /**
   * Validate email address format
   */
  validateEmail: function(email) {
    if (!email || typeof email !== 'string') {
      return false;
    }
    return Utils.isValidEmail(email);
  },
  
  /**
   * Send bulk emails with rate limiting (unchanged)
   */
  sendBulkEmails: function(emailList, subject, body, attachments = []) {
    const maxEmailsPerBatch = 100; // Gmail daily limit consideration
    let sentCount = 0;
    let failedCount = 0;
    
    try {
      for (let i = 0; i < emailList.length; i += maxEmailsPerBatch) {
        const batch = emailList.slice(i, i + maxEmailsPerBatch);
        
        batch.forEach(emailData => {
          try {
            if (this.validateEmail(emailData.email)) {
              const personalizedSubject = this.processTemplate(subject, emailData);
              const personalizedBody = this.processTemplate(body, emailData);
              
              MailApp.sendEmail(
                emailData.email,
                personalizedSubject,
                personalizedBody,
                { attachments: attachments }
              );
              sentCount++;
            } else {
              Logger.log(`Invalid email address: ${emailData.email}`);
              failedCount++;
            }
          } catch (error) {
            Logger.log(`Failed to send email to ${emailData.email}: ${error.message}`);
            failedCount++;
          }
        });
        
        // Rate limiting: pause between batches
        if (i + maxEmailsPerBatch < emailList.length) {
          Utilities.sleep(1000); // 1 second pause
        }
      }
      
      return {
        sent: sentCount,
        failed: failedCount,
        total: emailList.length
      };
      
    } catch (error) {
      Logger.log(`Bulk email error: ${error.message}`);
      throw error;
    }
  }
};

Logger.log('Enhanced EmailManager module loaded successfully');
