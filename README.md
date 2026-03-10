# SpreadsheetsKE Store
Professional Excel Templates for Kenyan Businesses with InstaSend Payment Integration

## Features
- 3 Professional Excel Bundles (Starter, Business, Enterprise)
- InstaSend Payment Integration with redirect to thank-you page
- Automatic Email Delivery of Bundles after payment
- Modern Web Interface with Template Preview

## Bundles Available
1. **Starter Bundle** - KES 1,900 (50 templates, 1 user, 6 months updates)
2. **Business Bundle** - KES 3,200 (150 templates, 10 users, 2 years updates) - Most Popular
3. **Enterprise Bundle** - KES 4,800 (150+ templates, unlimited users, lifetime updates)

## Setup Instructions

1. **Install Dependencies**
   ```bash
   npm install
   ```

2. **Configure Environment Variables**
   Edit the `.env` file with your credentials:
   - Get InstaSend API credentials from [InstaSend Dashboard](https://intasend.com/)
   - Set INSTASEND_PUBLISHABLE_KEY and INSTASEND_SECRET_KEY
   - Set up a Gmail account for email sending (use App Passwords)
   - Update the CALLBACK_URL to your server's public URL for webhooks

3. **Generate Excel Bundles**
   ```bash
   node src/generateTemplates.js
   ```

4. **Start the Server**
   ```bash
   npm start
   ```
   For development:
   ```bash
   npm run dev
   ```

5. **Access the Store**
   Open `http://localhost:3000` in your browser

## Debugging Payment Issues

If you get "Error creating checkout", check:

1. **InstaSend Credentials**: Ensure PUBLISHABLE_KEY and SECRET_KEY are correct from InstaSend dashboard
2. **Environment**: Use sandbox for testing: https://api.sandbox.intasend.com
3. **Webhook URL**: Set webhook URL in InstaSend dashboard to your-server/api/webhook
4. **Server Logs**: Check console for detailed error messages

For support, chat on WhatsApp: +254 720 288583

## M-Pesa Integration
- Uses Safaricom's Daraja API for STK Push
- Payments go to Paybill 329329, Account Number 0100451196700
- Automatic callback handling for payment confirmation
- Email delivery upon successful payment

## Production Deployment
- Set up a proper database for transaction storage
- Use HTTPS for the callback URL
- Configure proper email service (SendGrid, etc.)
- Set up monitoring and logging
