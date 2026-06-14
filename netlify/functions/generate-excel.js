const XLSX = require('xlsx');

const CAT_ORDER = [
  ['sauces','SAUCES'],['oven','OVEN'],['desserts','DESSERTS'],
  ['small-packs','SMALL PACKS'],['cheeses','CHEESES'],
  ['cured-meat','CURED MEAT'],['vegg','SALAD'],['frozen','FROZEN'],
  ['dry','DRY'],['pastas','PASTAS'],['drinks','DRINKS'],['oils','OILS'],
  ['packaging','PACKAGING'],['chemical','CHEMICAL'],
  ['office','OFFICE'],['backup','BACKUP'],
];

exports.handler = async (event) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Access-Control-Allow-Methods': 'POST, OPTIONS',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  try {
    const { orders, products } = JSON.parse(event.body || '{}');
    if (!orders || !orders.length) {
      return { statusCode: 400, headers, body: JSON.stringify({ error: 'No orders provided' }) };
    }

    const clients = [...new Set(orders.map(o => o.username))].sort();
    const wb = XLSX.utils.book_new();

    // ── SUMMARY SHEET ──
    const header1 = ['', 'PRODUCT', ...clients.map((_, i) => String(i+1).padStart(2,'0')), 'TOTAL'];
    const header2 = ['', '', ...clients, ''];
    const rows = [header1, header2];

    for (const [catId, catLabel] of CAT_ORDER) {
      const prods = products.filter(p => p.cat === catId);
      if (!prods.length) continue;

      rows.push([catLabel, '', ...clients.map(() => ''), '']);

      for (const p of prods) {
        const qtys = clients.map(cl => {
          const order = orders.find(o => o.username === cl);
          return order ? (parseInt(order.items[String(p.id)] || 0)) : 0;
        });
        const total = qtys.reduce((a, b) => a + b, 0);
        rows.push(['', p.name.toUpperCase(), ...qtys, total]);
      }

      rows.push(['', '', ...clients.map(() => ''), '']);
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = [{ wch: 12 }, { wch: 26 }, ...clients.map(() => ({ wch: 12 })), { wch: 10 }];
    XLSX.utils.book_append_sheet(wb, ws, 'SUMMARY');

    // ── PER-CLIENT SHEETS ──
    for (const cl of clients) {
      const order = orders.find(o => o.username === cl);
      if (!order) continue;
      const data = [['PRODUCT', 'UNIT', 'QTY']];
      for (const [catId, catLabel] of CAT_ORDER) {
        const prods = products.filter(p => p.cat === catId);
        const items = prods.filter(p => parseInt(order.items[String(p.id)] || 0) > 0);
        if (!items.length) continue;
        data.push([catLabel, '', '']);
        for (const p of items) {
          data.push([p.name.toUpperCase(), p.description || '', parseInt(order.items[String(p.id)])]);
        }
      }
      const ws2 = XLSX.utils.aoa_to_sheet(data);
      ws2['!cols'] = [{ wch: 26 }, { wch: 14 }, { wch: 8 }];
      XLSX.utils.book_append_sheet(wb, ws2, cl.substring(0, 31));
    }

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const date = new Date().toLocaleDateString('en-GB').replace(/\//g, '-');

    return {
      statusCode: 200,
      headers: {
        ...headers,
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="matteo-orders-${date}.xlsx"`,
      },
      body: buf.toString('base64'),
      isBase64Encoded: true,
    };
  } catch (e) {
    return { statusCode: 500, headers, body: JSON.stringify({ error: e.message }) };
  }
};
