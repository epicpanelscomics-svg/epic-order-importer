require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const axios = require('axios');
const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();

// Middleware
app.use(bodyParser.json({ limit: '50mb' }));
app.use(bodyParser.urlencoded({ limit: '50mb', extended: true }));
app.use(express.static('public'));

// Configuration from environment
const SHOPIFY_STORE = process.env.SHOPIFY_STORE || '0239a6-57';
const SHOPIFY_ACCESS_TOKEN = process.env.SHOPIFY_ACCESS_TOKEN;
const SHOPIFY_API_VERSION = '2024-01';

// Shopify API Helper
const shopifyAPI = {
  call: async (query, variables = {}) => {
    try {
      const response = await axios.post(
        `https://${SHOPIFY_STORE}.myshopify.com/admin/api/${SHOPIFY_API_VERSION}/graphql.json`,
        { query, variables },
        {
          headers: {
            'X-Shopify-Access-Token': SHOPIFY_ACCESS_TOKEN,
            'Content-Type': 'application/json',
          },
        }
      );
      return response.data;
    } catch (error) {
      console.error('Shopify API Error:', error.response?.data || error.message);
      throw error;
    }
  },
};

// Hidden Admin Page Route
app.get('/admin/epic-order-importer', (req, res) => {
  res.sendFile(path.join(__dirname, 'public/admin-page.html'));
});

// API: Create Orders from TikTok Data
app.post('/api/create-orders', async (req, res) => {
  try {
    const { orders } = req.body;

    if (!Array.isArray(orders) || orders.length === 0) {
      return res.status(400).json({ error: 'Invalid order data' });
    }

    const createdOrders = [];

    for (const order of orders) {
      if (!order.customer_email || !order.line_items || order.line_items.length === 0) {
        console.warn('Skipping invalid order:', order);
        continue;
      }

      const lineItems = order.line_items.map(item => ({
        variantId: item.variant_id || null,
        quantity: item.quantity || 1,
        customAttributes: item.custom_attributes || [],
      }));

      const mutation = `
        mutation CreateDraftOrder($input: DraftOrderInput!) {
          draftOrderCreate(input: $input) {
            draftOrder {
              id
              order {
                id
                name
                sourceOrderId
                fulfillments {
                  trackingInfo {
                    number
                    company
                  }
                }
              }
            }
            userErrors {
              field
              message
            }
          }
        }
      `;

      const variables = {
        input: {
          lineItems,
          customAttributes: [
            {
              key: 'source_platform',
              value: order.source_platform || 'TikTok',
            },
          ],
          email: order.customer_email,
          note: order.order_note || 'Imported from TikTok',
        },
      };

      try {
        const result = await shopifyAPI.call(mutation, variables);

        if (result.data?.draftOrderCreate?.draftOrder) {
          createdOrders.push({
            draftOrderId: result.data.draftOrderCreate.draftOrder.id,
            email: order.customer_email,
            source_data: order.source_order_id || 'N/A',
          });
        }
      } catch (error) {
        console.error('Error creating order:', error);
      }
    }

    res.json({
      success: true,
      message: `Created ${createdOrders.length} orders`,
      orders: createdOrders,
    });
  } catch (error) {
    console.error('Order creation error:', error);
    res.status(500).json({ error: error.message });
  }
});

// API: Fetch Order Data
app.post('/api/fetch-orders', async (req, res) => {
  try {
    const { order_ids } = req.body;

    if (!Array.isArray(order_ids) || order_ids.length === 0) {
      return res.status(400).json({ error: 'No order IDs provided' });
    }

    const ordersData = [];

    for (const orderId of order_ids) {
      const query = `
        query {
          order(id: "${orderId}") {
            id
            name
            sourceOrderId
            fulfillments {
              trackingInfo {
                number
                company
              }
            }
          }
        }
      `;

      try {
        const result = await shopifyAPI.call(query);

        if (result.data?.order) {
          const order = result.data.order;
          const tracking = order.fulfillments?.[0]?.trackingInfo || {};

          ordersData.push({
            orderId: order.id,
            orderName: order.name,
            sourceOrderId: order.sourceOrderId || 'N/A',
            trackingNumber: tracking.number || 'N/A',
            carrier: tracking.company || 'USPS',
          });
        }
      } catch (error) {
        console.error(`Error fetching order ${orderId}:`, error);
      }
    }

    res.json({
      success: true,
      orders: ordersData,
    });
  } catch (error) {
    console.error('Fetch orders error:', error);
    res.status(500).json({ error: error.message });
  }
});

// API: Populate and Export XLSX Template
app.post('/api/export-template', async (req, res) => {
  try {
    const { templateFile, orderData } = req.body;

    if (!templateFile || !Array.isArray(orderData)) {
      return res.status(400).json({ error: 'Invalid template or order data' });
    }

    const templateBuffer = Buffer.from(templateFile, 'base64');
    const workbook = XLSX.read(templateBuffer, { type: 'buffer' });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    let row = 3;
    for (const order of orderData) {
      worksheet[`A${row}`] = { v: order.sourceOrderId, t: 's' };
      worksheet[`B${row}`] = { v: order.trackingNumber, t: 's' };
      worksheet[`C${row}`] = { v: order.carrier, t: 's' };
      row++;
    }

    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    range.e.r = Math.max(range.e.r, row - 1);
    worksheet['!ref'] = XLSX.utils.encode_range(range);

    const outputFile = path.join(__dirname, 'public/Shipment_Info_Exported.xlsx');
    XLSX.writeFile(workbook, outputFile);

    res.download(outputFile, 'Shipment_Info_Exported.xlsx', (err) => {
      if (err) console.error('Download error:', err);
      fs.unlink(outputFile, (err) => {
        if (err) console.error('Cleanup error:', err);
      });
    });
  } catch (error) {
    console.error('Export error:', error);
    res.status(500).json({ error: error.message });
  }
});

// API: Parse uploaded file
app.post('/api/parse-file', (req, res) => {
  try {
    const { fileData } = req.body;

    if (!fileData) {
      return res.status(400).json({ error: 'No file data provided' });
    }

    const buffer = Buffer.from(fileData, 'base64');
    const workbook = XLSX.read(buffer, { type: 'buffer' });
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);

    res.json({
      success: true,
      data,
      rowCount: data.length,
    });
  } catch (error) {
    console.error('Parse error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Health check
app.get('/health', (req, res) => {
  res.json({ status: 'ok', store: SHOPIFY_STORE });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`✅ Epic Panels Order Importer running on port ${PORT}`);
  console.log(`📍 Visit: http://localhost:${PORT}/admin/epic-order-importer`);
});
