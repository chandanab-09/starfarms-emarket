const express = require('express');
const mysql = require('mysql2');
const bcrypt = require('bcrypt');
const bodyParser = require('body-parser');
const cors = require('cors');
const path = require('path');
const ExcelJS = require('exceljs');

const app = express();
const PORT = 3000;

// Middleware
app.use(cors());
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Database connection
const db = mysql.createConnection({
    host: 'localhost',
    user: 'root',
    password: 'YOUR_DB_PASSWORD',
    database: 'login_system',
    multipleStatements: true
});

db.connect((err) => {
    if (err) {
        console.error('Database connection failed:', err);
        return;
    }
    console.log('Connected to MySQL database: login_system');
});

// --- API Endpoint: Signup ---
app.post('/api/signup', async (req, res) => {
    const { firstName, lastName, email, password, phone, pincode, address, city } = req.body;

    if (!firstName || !email || !password || password.length < 8) {
        return res.status(400).json({ error: 'Invalid input' });
    }

    try {
        const hashedPassword = await bcrypt.hash(password, 10);
        const query = "INSERT INTO users (firstName, lastName, email, password, phone, pincode, address, city) VALUES (?, ?, ?, ?, ?, ?, ?, ?)";
        
        db.execute(query, [firstName, lastName, email, hashedPassword, phone, pincode, address, city], (err, result) => {
            if (err) {
                console.error("Signup DB Error:", err);
                if (err.code === 'ER_DUP_ENTRY') return res.status(400).json({ error: 'Email already exists' });
                return res.status(500).json({ error: 'Database error' });
            }
            res.json({ message: 'Signup successful', userId: result.insertId });
        });
    } catch (error) {
        res.status(500).json({ error: 'Server error' });
    }
});

// --- API Endpoint: Login (Handles both Admin and User) ---
app.post('/api/login', (req, res) => {
    const { username, password } = req.body;

    if (!username || !password) {
        return res.status(400).json({ error: 'Credentials required' });
    }

    const ADMIN_ID = "admin@starfarms.com";
    const ADMIN_PASS = "ADMIN_PASSWORD";

    if (username === ADMIN_ID && password === ADMIN_PASS) {
        return res.json({
            message: 'Admin login successful',
            role: 'admin', 
            user: { firstName: 'System', lastName: 'Admin', email: ADMIN_ID }
        });
    }

    const query = `SELECT * FROM users WHERE email = ? OR phone = ?`;
    db.execute(query, [username, username], async (err, results) => {
        if (err) return res.status(500).json({ error: 'Database error' });
        if (results.length === 0) return res.status(401).json({ error: 'Invalid credentials' });

        const user = results[0];
        const isValid = await bcrypt.compare(password, user.password);

        if (!isValid) return res.status(401).json({ error: 'Invalid credentials' });

        res.json({
            message: 'User login successful',
            role: 'user', 
            user: {
                id: user.id,
                firstName: user.firstName,
                lastName: user.lastName,
                email: user.email,
                phone: user.phone,
                address: user.address,
                city: user.city,
                pincode: user.pincode
            }
        });
    });
}); // <--- THIS WAS THE MISSING CLOSING BRACKET

// --- API Endpoint: Place Order ---
app.post('/api/place-order', (req, res) => {
    const { customer, cart, total } = req.body;

    if (!customer || !cart || cart.length === 0) {
        return res.status(400).json({ success: false, message: 'Invalid order data' });
    }

    const orderQuery = `
        INSERT INTO orders (customer_name, email, phone, address, city, pincode, payment, total, created_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, NOW())
    `;

    db.execute(orderQuery, [
        customer.name, customer.email, customer.phone,
        customer.address, customer.city, customer.pincode,
        customer.payment, total
    ], (err, orderResult) => {
        if (err) {
            console.error("Order Table Error:", err);
            return res.status(500).json({ success: false, message: 'Database error: Main Order' });
        }

        const orderId = orderResult.insertId;
        const itemQuery = `INSERT INTO order_items (order_id, product_id, product_name, quantity, price) VALUES (?, ?, ?, ?, ?)`;
        const updateStockQuery = `UPDATE products SET stock = stock - ?, stock_quantity = stock_quantity - ? WHERE id = ?`;

        let itemPromises = cart.map(item => {
            return new Promise((resolve, reject) => {
                const productId = item.id || item.product_id;
                let priceValue = typeof item.price === 'string' ? parseFloat(item.price.replace(/[^0-9.]/g, '')) || 0 : item.price;

                if (!productId) return reject(new Error(`Missing ID for item: ${item.name}`));

                db.execute(itemQuery, [orderId, String(productId), item.name, item.quantity, priceValue], (itemErr) => {
                    if (itemErr) return reject(itemErr);

                    db.execute(updateStockQuery, [item.quantity, item.quantity, productId], (stockErr) => {
                        if (stockErr) console.warn(`Stock update skipped for ${productId}: ${stockErr.message}`);
                        resolve();
                    });
                });
            });
        });

        Promise.all(itemPromises)
            .then(() => res.json({ success: true, orderId }))
            .catch(error => res.status(500).json({ success: false, message: error.message || 'Failed to save items' }));
    });
});

// --- API Endpoint: Generate Bill ---
app.post('/api/generate-bill', async (req, res) => {
    const { customer, cart } = req.body;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Bill');

    sheet.addRow(["STAR FARMS BILL"]).font = { size: 18, bold: true };
    sheet.addRow([]);
    const orderId = "SF-" + Math.floor(Math.random() * 10000);
    const date = new Date().toLocaleString();

    sheet.addRow(["Order ID:", orderId]);
    sheet.addRow(["Date:", date]);
    sheet.addRow([]);
    sheet.addRow(["Name:", customer.name]);
    sheet.addRow(["Email:", customer.email]);
    sheet.addRow(["Phone:", customer.phone]);
    sheet.addRow([]);
    
    const headerRow = sheet.addRow(["Product", "Quantity", "Price", "Subtotal"]);
    headerRow.font = { bold: true };

    let totalAmount = 0;
    cart.forEach(item => {
        let price = parseFloat(String(item.price).replace(/[^0-9.]/g,'')) || 0;
        let qty = item.quantity || 1;
        let subtotal = price * qty;
        totalAmount += subtotal;
        sheet.addRow([item.name, qty, price, subtotal]);
    });

    sheet.addRow([]);
    sheet.addRow(["", "", "Total", totalAmount]);
    sheet.columns = [{ width: 30 }, { width: 15 }, { width: 15 }, { width: 20 }];

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", "attachment; filename=StarFarms_Bill.xlsx");

    await workbook.xlsx.write(res);
    res.end();
});

// --- ADMIN APIs ---

app.get('/api/admin/overview-stats', (req, res) => {
    // Ensure there are no trailing spaces and semicolon at the end of each
    const query = "SELECT COUNT(*) as userCount FROM users; SELECT COUNT(*) as orderCount FROM orders; SELECT COUNT(*) as contactCount FROM contacts;";

    db.query(query, (err, results) => {
        if (err) {
            console.error("Dashboard Query Error:", err);
            return res.status(500).json({ error: err.message });
        }

        // Add a check: If results[2] doesn't exist, it means multipleStatements failed
        if (!results || results.length < 3) {
            console.error("Multiple statements failed. Results received:", results.length);
            return res.status(500).json({ error: "Database did not return all stats. Check if 'contacts' table exists." });
        }

        res.json({
            totalUsers: results[0][0].userCount,
            totalOrders: results[1][0].orderCount,
            totalContacts: results[2][0].contactCount
        });
    });
});




app.get('/api/admin/get-all-users', (req, res) => {
    const query = "SELECT id, firstName, lastName, email, city FROM users ORDER BY id DESC";
    db.query(query, (err, results) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(results);
    });
});

app.get('/api/admin/get-all-orders', (req, res) => {
    const query = `
        SELECT o.id, o.customer_name, o.total, o.created_at, 
        (SELECT product_name FROM order_items WHERE order_id = o.id LIMIT 1) as main_product
        FROM orders o 
        ORDER BY o.created_at DESC`;
    db.query(query, (err, results) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(results);
    });
});

app.post('/api/admin/add-product', (req, res) => {
    const { name, description, price, category, image } = req.body;
    const query = "INSERT INTO products (name, description, price, category, image, stock) VALUES (?, ?, ?, ?, ?, 10)";
    db.execute(query, [name, description, price, category, image], (err, result) => {
        if (err) return res.status(500).json({ error: "Failed to add product" });
        res.json({ message: "Product added successfully!", id: result.insertId });
    });
});

// GET ALL PRODUCTS
app.get('/api/admin/get-all-products', (req, res) => {
    const sql = "SELECT * FROM products ORDER BY id DESC";

    db.query(sql, (err, result) => {
        if (err) {
            console.error(err);
            return res.status(500).json({ error: "Failed to fetch products" });
        }
        res.json(result);
    });
});


// DELETE PRODUCT
app.delete('/api/admin/delete-product/:id', (req, res) => {
    const { id } = req.params;

    const sql = "DELETE FROM products WHERE id = ?";

    db.query(sql, [id], (err, result) => {
        if (err) {
            console.error(err);
            return res.status(500).json({ error: "Delete failed" });
        }
        res.json({ message: "Product deleted successfully" });
    });
});

// GET ALL ORDERS FOR REPORTS
app.get('/api/orders', (req, res) => {
    // Basic security check matching your frontend key
    if (req.query.key !== 'SECRET_KEY') return res.status(403).send('Forbidden');

    const query = "SELECT id, customer_name, email, total, created_at FROM orders ORDER BY created_at DESC";
    db.query(query, (err, results) => {
        if (err) return res.status(500).json({ error: err.message });
        res.json(results);
    });
});

//----sales-Report---
app.get('/api/sales-report', async (req, res) => {
    if (req.query.key !== 'SECRET_KEY') return res.status(403).send('Forbidden');

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Total Sales');

    // --- FIX: THIS SECTION DEFINES THE COLUMN WIDTHS ---
    sheet.columns = [
        { header: 'Order ID', key: 'id', width: 12 },
        { header: 'Customer', key: 'customer_name', width: 25 },
        { header: 'Email', key: 'email', width: 35 }, // Wide enough for long email addresses
        { header: 'Total Amount', key: 'total', width: 15 },
        { header: 'Date', key: 'created_at', width: 25 }
    ];

    db.query("SELECT id, customer_name, email, total, created_at FROM orders", async (err, results) => {
        if (err) return res.status(500).send("Database error");

        // Use the row objects directly or map them to the keys defined above
        results.forEach(order => {
            sheet.addRow({
                id: order.id,
                customer_name: order.customer_name,
                email: order.email,
                total: order.total,
                created_at: new Date(order.created_at).toLocaleString()
            });
        });

        // Make the header row bold so it looks professional
        sheet.getRow(1).font = { bold: true };

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", "attachment; filename=StarFarms_Full_Sales_Report.xlsx");
        
        await workbook.xlsx.write(res);
        res.end();
    });
});

// download individual bill
app.get('/api/download-bill/:id', async (req, res) => {
    const orderId = req.params.id;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Invoice');

    // Set Column Widths
    sheet.columns = [
        { header: 'Item Details', key: 'col1', width: 25 },
        { header: '', key: 'col2', width: 15 },
        { header: '', key: 'col3', width: 15 },
        { header: '', key: 'col4', width: 15 }
    ];

    // We use a JOIN or two queries to get Order Info + Product Items
    const sql = `
        SELECT * FROM orders WHERE id = ?;
        SELECT * FROM order_items WHERE order_id = ?;
    `;

    db.query(sql, [orderId, orderId], async (err, results) => {
        if (err || results[0].length === 0) return res.status(404).send("Order not found");

        const order = results[0][0];
        const items = results[1]; // This contains all products for this order
        const formattedDate = new Date(order.created_at).toLocaleString();

        // Header Section
        sheet.addRow(["STAR FARMS INVOICE"]).font = { bold: true, size: 16 };
        sheet.addRow([]);
        sheet.addRow(["Order ID:", order.id]);
        sheet.addRow(["Customer:", order.customer_name]);
        sheet.addRow(["Date:", formattedDate]);
        sheet.addRow([]);

        // Product Table Header
        const tableHeader = sheet.addRow(["Product Name", "Quantity", "Price", "Subtotal"]);
        tableHeader.font = { bold: true };
        tableHeader.alignment = { horizontal: 'center' };

        // --- NEW: Loop through items to show Product Names ---
        items.forEach(item => {
            const subtotal = item.quantity * item.price;
            sheet.addRow([
                item.product_name, 
                item.quantity, 
                "₹" + item.price, 
                "₹" + subtotal
            ]);
        });

        sheet.addRow([]);
        const totalRow = sheet.addRow(["", "", "Total Paid:", "₹" + order.total]);
        totalRow.font = { bold: true };

        // Final Styling
        sheet.getColumn(1).font = { bold: true };

        res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        res.setHeader("Content-Disposition", `attachment; filename=StarFarms-Bill-${orderId}.xlsx`);
        
        await workbook.xlsx.write(res);
        res.end();
    });
});
// --- NEW API: SAVE CONTACT MESSAGE FROM HOME PAGE ---
app.post('/api/contact', (req, res) => {
    const { name, email, message } = req.body;

    if (!name || !email || !message) {
        return res.status(400).json({ error: "All fields are required" });
    }

    const query = "INSERT INTO contacts (name, email, message, created_at) VALUES (?, ?, ?, NOW())";
    
    db.execute(query, [name, email, message], (err, result) => {
        if (err) {
            console.error("Save Contact Error:", err);
            return res.status(500).json({ error: "Failed to save message" });
        }
        res.json({ message: "Message sent successfully!", id: result.insertId });
    });
});
// --- NEW API: GET ALL CONTACT MESSAGES ---
app.get('/api/admin/get-all-contacts', (req, res) => {
    // This query fetches messages from your contacts table
    const query = "SELECT id, name, email, message, created_at FROM contacts ORDER BY created_at DESC";
    
    db.query(query, (err, results) => {
        if (err) {
            console.error("Contacts DB Error:", err);
            return res.status(500).json({ error: "Failed to fetch contacts" });
        }
        // If successful, send the list of messages to the admin panel
        res.json(results);
    });
});


// Serve HTML pages
app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'login.html')));
app.get('/signup', (req, res) => res.sendFile(path.join(__dirname, 'public', 'signup.html')));
app.get('/admin', (req, res) => res.sendFile(path.join(__dirname, 'public', 'admin.html')));

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`));