const express = require('express');
const XLSX = require('xlsx');
const path = require('path');
const bodyParser = require('body-parser');
const app = express();
const PORT = 3000;

// إعداد ملفات static والمجلدات
app.use(express.static('public'));
app.set('view engine', 'ejs');
app.use(bodyParser.urlencoded({ extended: true }));

// مسار الملف Excel
const filePath = path.join(__dirname, 'data', '444.xlsx');

// دالة لتحميل البيانات من الملف Excel
function loadData() {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(sheet);
  return data;
}

// دالة لتحديث حالة الاشتراك
function updateSubscriptionStatus(data) {
  const currentDate = new Date();
  data.forEach(subscription => {
    const endDate = new Date(subscription['تاريخ الانتهاء']);
    const delta = Math.floor((endDate - currentDate) / (1000 * 60 * 60 * 24));

    if (isNaN(delta)) {
      subscription['حالة الاشتراك'] = 'تاريخ غير صحيح';
    } else if (delta > 0) {
      subscription['حالة الاشتراك'] = `باقٍ ${delta} يوم`;
    } else if (delta === 0) {
      subscription['حالة الاشتراك'] = 'ينتهي اليوم';
    } else {
      subscription['حالة الاشتراك'] = `منتهٍ منذ ${-delta} يوم`;
    }
  });
  return data;
}

// الصفحة الرئيسية
app.get('/', (req, res) => {
  res.render('index');
});

// عرض الاشتراكات
app.get('/subscriptions', (req, res) => {
  let data = loadData();
  data = updateSubscriptionStatus(data);
  res.render('subscriptions', { subscriptions: data });
});

// عرض الاشتراكات المنتهية
app.get('/expired_subscriptions', (req, res) => {
  let data = loadData();
  data = updateSubscriptionStatus(data);
  const expiredData = data.filter(sub => sub['حالة الاشتراك'].includes('منتهٍ'));
  res.render('expired_subscriptions', { expiredSubscriptions: expiredData });
});

// إضافة اشتراك
app.post('/add_subscription', (req, res) => {
  const { name, subscription_date, end_date, phone, subscription_type } = req.body;
  let data = loadData();
  const newSubscription = {
    'الاســــــم': name,
    'تاريخ الاشتراك': subscription_date,
    'تاريخ الانتهاء': end_date,
    'رقم الهاتف': phone,
    'نوع الاشتراك': subscription_type,
    'حالة الاشتراك': 'باقٍ'
  };
  data.push(newSubscription);
  data = updateSubscriptionStatus(data);

  // كتابة البيانات في الملف Excel
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, filePath);

  res.redirect('/');
});

// حذف اشتراك
app.post('/delete_subscription', (req, res) => {
  const { name } = req.body;
  let data = loadData();
  data = updateSubscriptionStatus(data);
  const index = data.findIndex(sub => sub['الاســــــم'].includes(name));
  if (index !== -1) {
    data.splice(index, 1);
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filePath);
    res.redirect('/');
  } else {
    res.send('الاسم غير موجود في النظام!');
  }
});

// البحث عن الاشتراك
app.post('/search_subscription', (req, res) => {
  const searchQuery = req.body.search_query;
  let data = loadData();
  const result = data.filter(sub => sub['الاســــــم'].includes(searchQuery));
  if (result.length > 0) {
    res.render('search_subscription', { searchQuery, subscriptions: result });
  } else {
    res.send('لا توجد اشتراكات تتطابق مع البحث!');
  }
});

app.listen(PORT, () => {
  console.log(`Server is running on http://localhost:${PORT}`);
});
