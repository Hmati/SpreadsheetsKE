require('dotenv').config();
const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const axios = require('axios');
const nodemailer = require('nodemailer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, '../public')));

// InstaSend credentials
const INSTASEND_PUBLISHABLE_KEY = process.env.INSTASEND_PUBLISHABLE_KEY;
const INSTASEND_SECRET_KEY = process.env.INSTASEND_SECRET_KEY;
const INSTASEND_BASE_URL = process.env.INSTASEND_BASE_URL || 'https://api.sandbox.intasend.com';

// Email config
const EMAIL_USER = process.env.EMAIL_USER;
const EMAIL_PASS = process.env.EMAIL_PASS;

// Products (Excel bundles)
const products = [
  { id: 1, name: 'Starter Bundle', price: 1900, file: 'starter_bundle.xlsx' },
  { id: 2, name: 'Business Bundle', price: 3200, file: 'business_bundle.xlsx' },
  { id: 3, name: 'Enterprise Bundle', price: 4800, file: 'enterprise_bundle.xlsx' }
];

// InstaSend headers
const getInstaSendHeaders = () => ({
  'Authorization': `Bearer ${INSTASEND_SECRET_KEY}`,
  'Content-Type': 'application/json',
  'X-IntaSend-Public-Key': INSTASEND_PUBLISHABLE_KEY
});

// Create InstaSend checkout
app.post('/api/create-checkout', async (req, res) => {
  const { productId, email } = req.body;

  const product = products.find(p => p.id == productId);
  if (!product) {
    return res.status(400).json({ error: 'Product not found' });
  }

  try {
    const checkoutData = {
      public_key: INSTASEND_PUBLISHABLE_KEY,
      amount: product.price,
      currency: 'KES',
      email: email,
      redirect_url: `${req.protocol}://${req.get('host')}/thank-you.html`,
      api_ref: `order_${Date.now()}_${productId}`,
      comment: `Payment for ${product.name}`,
      // Store product info for webhook processing
      metadata: {
        productId: productId,
        email: email
      }
    };

    const response = await axios.post(`${INSTASEND_BASE_URL}/api/v1/checkout/`, checkoutData, {
      headers: getInstaSendHeaders()
    });

    // Store checkout details for webhook verification
    checkouts[response.data.id] = { productId, email };

    res.json({
      success: true,
      checkout_url: response.data.url,
      checkout_id: response.data.id
    });
  } catch (error) {
    console.error('InstaSend checkout error:', error.response ? error.response.data : error.message);
    if (process.env.NODE_ENV !== 'production') {
      res.status(500).json({
        error: 'Failed to create checkout',
        details: error.response ? error.response.data : error.message
      });
    } else {
      res.status(500).json({ error: 'Failed to create checkout' });
    }
  }
});

// InstaSend webhook
app.post('/api/webhook', (req, res) => {
  const webhookData = req.body;
  console.log('InstaSend Webhook:', webhookData);

  // Verify webhook signature if needed (InstaSend provides signature verification)
  // For now, we'll trust the webhook

  if (webhookData.state === 'COMPLETE' && webhookData.invoice) {
    const checkoutId = webhookData.invoice.checkout_id;
    const checkout = checkouts[checkoutId];

    if (checkout) {
      // Payment successful - send email
      sendTemplateEmail(checkout.productId, checkout.email);
      delete checkouts[checkoutId];
    }
  }

  res.json({ success: true });
});

// Send template email
async function sendTemplateEmail(productId, email) {
  const product = products.find(p => p.id == productId);
  if (!product) return;

  const transporter = nodemailer.createTransport({
    service: 'gmail',
    auth: {
      user: EMAIL_USER,
      pass: EMAIL_PASS
    }
  });

  const mailOptions = {
    from: EMAIL_USER,
    to: email,
    subject: `Your ${product.name} Template`,
    text: `Thank you for your purchase! Attached is your ${product.name}.`,
    attachments: [
      {
        filename: product.file,
        path: path.join(__dirname, '../templates', product.file)
      }
    ]
  };

  try {
    await transporter.sendMail(mailOptions);
    console.log('Email sent successfully');
  } catch (error) {
    console.error('Email send error:', error);
  }
}

// Get products
app.get('/api/products', (req, res) => {
  res.json(products);
});

// Preview product
app.get('/api/preview/:productId', async (req, res) => {
  const productId = parseInt(req.params.productId);
  const product = products.find(p => p.id === productId);

  if (!product) {
    return res.status(404).json({ error: 'Product not found' });
  }

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(__dirname, '../templates', product.file));

    const previewData = {};

    workbook.eachSheet((worksheet, sheetId) => {
      const sheetName = worksheet.name;
      const rows = [];

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber <= 10) { // Only first 10 rows for preview
          const rowData = [];
          row.eachCell((cell, colNumber) => {
            rowData.push(cell.value || '');
          });
          rows.push(rowData);
        }
      });

      previewData[sheetName] = rows;
    });

    res.json({ productName: product.name, data: previewData });
  } catch (error) {
    console.error('Preview error:', error);
    res.status(500).json({ error: 'Failed to load preview' });
  }
});

// In-memory checkout store (replace with database in production)
const checkouts = {};

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});