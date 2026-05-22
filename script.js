let salesData = []; 
let forecastTrendChart;  
let yearlyForecastChart; 
let salesChart, stockChart, priceTrendChart, forecastChart;

// DOM Elements
const fileInput = document.getElementById('fileUpload');
const previewTableBody = document.querySelector('#dataPreviewTable tbody');

document.querySelectorAll('.nav-item').forEach(item => {
    item.addEventListener('click', (e) => {
        e.preventDefault();
        const tabId = item.dataset.tab;
        document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
        document.querySelectorAll('.nav-item').forEach(nav => nav.classList.remove('active'));
        item.classList.add('active');
        document.getElementById(`${tabId}-tab`).classList.add('active');
        
        if(tabId === 'analytics' && salesData.length) updateAllCharts();
        if(tabId === 'alerts') renderStockAlerts();
        if(tabId === 'insights') renderAIInsights();
        if(tabId === 'overstock') renderStockHealthTab();
        
        // GABUNGKAN FORECAST JADI SATU
        if(tabId === 'forecast') {
            runForecastingAnalysis();
            updateDemandPerProduct();     // demand per produk
            updateYearlyForecast();       // prediksi 12 bulan
        }
    });
});

const menuToggle = document.querySelector('.menu-toggle');
const sidebar = document.querySelector('.sidebar');

let overlay = document.querySelector('.sidebar-overlay');
if(!overlay) {
    overlay = document.createElement('div');
    overlay.className = 'sidebar-overlay';
    document.body.appendChild(overlay);
}

if(menuToggle) {
    menuToggle.addEventListener('click', () => {
        sidebar.classList.toggle('open');
        overlay.classList.toggle('active');
    });
}

// Tutup sidebar saat klik overlay
overlay.addEventListener('click', () => {
    sidebar.classList.remove('open');
    overlay.classList.remove('active');
});

// ========== FILE UPLOAD & PARSING ==========
fileInput.addEventListener('change', handleFileUpload);

function handleFileUpload(e) {
    const file = e.target.files[0];
    if(!file) return;
    
    const extension = file.name.split('.').pop().toLowerCase();
    const reader = new FileReader();
    
    reader.onload = function(evt) {
        if(extension === 'csv') {
            parseCSV(evt.target.result);
        } else if(['xlsx', 'xls'].includes(extension)) {
            const data = new Uint8Array(evt.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const sheet = workbook.Sheets[workbook.SheetNames[0]];
            const json = XLSX.utils.sheet_to_json(sheet);
            processData(json);
        }
    };
    
    if(extension === 'csv') reader.readAsText(file);
    else reader.readAsArrayBuffer(file);
}

function parseCSV(csvText) {
    const lines = csvText.split('\n');
    const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
    const dataRows = [];
    for(let i=1; i<lines.length; i++) {
        if(!lines[i].trim()) continue;
        const values = lines[i].split(',');
        let obj = {};
        headers.forEach((h,idx) => obj[h] = values[idx]);
        dataRows.push(obj);
    }
    processData(dataRows);
}

function processData(jsonData) {
    // Group data berdasarkan produk, ambil stok dari tanggal terbaru
    const productMap = new Map();
    
    jsonData.forEach(row => {
        const productName = row.nama_produk || row.produk || row.nama || 'Unknown';
        const harga = parseFloat(row.harga_jual || row.harga || 0);
        const stok = parseFloat(row.stok || 0);
        const terjual = parseFloat(row.terjual || row.terjual_harian || 0);
        const tanggal = row.tanggal || new Date().toISOString().split('T')[0];
        
        if(isNaN(harga) || isNaN(stok) || isNaN(terjual)) return;
        
        if(!productMap.has(productName)) {
            productMap.set(productName, {
                nama_produk: productName,
                harga_jual: harga,
                stok: stok,
                terjual: terjual,
                tanggal: tanggal
            });
        } else {
            const existing = productMap.get(productName);
            const existingDate = new Date(existing.tanggal);
            const newDate = new Date(tanggal);
            
            // Jika tanggal baru LEBIH BARU, update data (TIDAK DIJUMLAH!)
            if(newDate > existingDate) {
                existing.harga_jual = harga;
                existing.stok = stok;
                existing.terjual = terjual;
                existing.tanggal = tanggal;
            }
        }
    });
    
    // Konversi Map ke array
    salesData = Array.from(productMap.values());
    salesData.sort((a,b) => a.nama_produk.localeCompare(b.nama_produk));
    
    if(salesData.length === 0) {
        document.getElementById('uploadStatus').innerHTML = '<span style="color:red;">❌ Format data tidak sesuai</span>';
        return;
    }
    
    const totalStokKeseluruhan = salesData.reduce((sum, p) => sum + p.stok, 0);
    
    document.getElementById('uploadStatus').innerHTML = `
        <span style="color:green;">✅ Data berhasil diupload!</span><br>
        📊 ${salesData.length} produk unik ditemukan<br>
        📦 Total stok keseluruhan: ${totalStokKeseluruhan.toLocaleString()} unit<br>
        
    `;
    
    renderPreviewTable();
    updateDashboardStats();
    runForecastingAnalysis();
    updateForecastTrendChart();
    renderStockAlerts();
    updateAllCharts();
    updatePerformanceStats();  // CUKUP TAMBAHKAN 1 BARIS INI
    renderAIInsights();
    initPriceWarSimulator();
    updateYearlyForecast();
    updateDemandPerProduct();
    populateYearlyProductDropdown();
}

// Preview Tabel
function renderPreviewTable() {
    previewTableBody.innerHTML = '';
    salesData.slice(0, 10).forEach(p => {
        const row = `<tr>
            <td>${p.nama_produk}</td>
            <td>Rp ${p.harga_jual.toLocaleString()}</td>
            <td>${p.stok}</td>
            <td>${p.terjual}</td>
            <td>${p.tanggal}</td>
        </tr>`;
        previewTableBody.insertAdjacentHTML('beforeend', row);
    });
}

// ========== INI BUAT STATISTIK CARD YAAAAAAAAAAAAAAAA =================
function updateDashboardStats() {
    const totalPenjualan = salesData.reduce((sum, p) => sum + (p.harga_jual * p.terjual), 0);
    let topProduct = salesData.reduce((max, p) => p.terjual > (max.terjual || 0) ? p : max, {nama_produk:'-', terjual:0});
    const totalProfit = totalPenjualan * 0.25; // estimasi margin 25%
    const lowStockCount = salesData.filter(p => p.stok < 30 && p.stok > 0).length;
    const predictedDemand = Math.round(salesData.reduce((sum,p) => sum + p.terjual,0) / salesData.length * 1.1);
    
    // Rekomendasi harga AI sederhana
    let rekomHarga = 'Stabil';
    const highDemandLowStock = salesData.filter(p => p.terjual > 50 && p.stok < 20).length;
    if(highDemandLowStock > 0) rekomHarga = 'Naik 5-10%';
    else if(salesData.filter(p => p.terjual < 10 && p.stok > 50).length > 0) rekomHarga = 'Turun 5%';
    
    document.getElementById('totalSales').innerText = `Rp ${Math.round(totalPenjualan).toLocaleString()}`;
    document.getElementById('topProduct').innerText = topProduct.nama_produk;
    document.getElementById('totalProfit').innerText = `Rp ${Math.round(totalProfit).toLocaleString()}`;
    document.getElementById('lowStockCount').innerText = lowStockCount;
    document.getElementById('predictedDemand').innerText = predictedDemand;
    document.getElementById('priceReco').innerHTML = rekomHarga;
}

// ========== FORECASTING DENGAN WEIGHTED MOVING AVERAGE (WMA) ==========
// ========== FORECASTING DENGAN WEIGHTED MOVING AVERAGE (WMA) ==========
function runForecastingAnalysis() {
    console.log('runForecastingAnalysis dipanggil, salesData length:', salesData.length);
    
    if(!salesData.length) {
        document.getElementById('demandPredictionText').innerHTML = 'Belum ada data. Upload file terlebih dahulu.';
        return;
    }

    try {
        // Urutkan data berdasarkan tanggal
        const sortedData = [...salesData].sort((a, b) => new Date(a.tanggal) - new Date(b.tanggal));
        
        // Group data per produk
        const productsMap = new Map();
        sortedData.forEach(item => {
            if(!productsMap.has(item.nama_produk)) {
                productsMap.set(item.nama_produk, []);
            }
            productsMap.get(item.nama_produk).push(item);
        });
        
        // Weighted Moving Average Function
        function weightedMovingAverage(data, weights = [0.5, 0.3, 0.2]) {
            if(!data.length) return 0;
            if(data.length < 3) {
                const avg = data.reduce((s, p) => s + p.terjual, 0) / data.length;
                return Math.round(avg);
            }
            const lastThree = data.slice(-3).map(p => p.terjual);
            let weightedSum = 0;
            let weightSum = 0;
            for(let i = 0; i < lastThree.length; i++) {
                weightedSum += lastThree[i] * weights[i];
                weightSum += weights[i];
            }
            return Math.round(weightedSum / weightSum);
        }
        
        // Detect Trend Function
        function detectTrend(data) {
            if(data.length < 5) return { trend: "➡️ Data Tidak Cukup", percent: 0 };
            const firstThree = data.slice(0, 3).reduce((s, p) => s + p.terjual, 0);
            const lastThree = data.slice(-3).reduce((s, p) => s + p.terjual, 0);
            if(firstThree === 0) return { trend: "➡️ Stabil", percent: 0 };
            const percentChange = ((lastThree - firstThree) / firstThree) * 100;
            let trendText = "Stabil";
            if(percentChange > 30) trendText = "🚀 Meningkat Pesat";
            else if(percentChange > 15) trendText = "📈 Meningkat";
            else if(percentChange > 5) trendText = "⬆️ Cenderung Naik";
            else if(percentChange < -30) trendText = "💀 Menurun Drastis";
            else if(percentChange < -15) trendText = "📉 Menurun";
            else if(percentChange < -5) trendText = "⬇️ Cenderung Turun";
            return { trend: trendText, percent: percentChange };
        }
        
        // Hitung forecasting
        const allSales = sortedData.map(p => ({ terjual: p.terjual, tanggal: p.tanggal }));
        const forecastPerDay = weightedMovingAverage(allSales);
        const forecast7Days = forecastPerDay * 7;
        const overallTrend = detectTrend(allSales);
        
        // Produk Butuh Restock
        function getProductsNeedRestock() {
            const needRestock = [];
            for(const [productName, productData] of productsMap) {
                const lastItem = productData[productData.length - 1];
                const stok = lastItem.stok;
                const terjual = lastItem.terjual || 1;
                if(stok < 40) {
                    needRestock.push({ nama: productName, stok: stok, terjual: terjual });
                }
            }
            needRestock.sort((a, b) => a.stok - b.stok);
            if(needRestock.length === 0) return '<li>✅ Semua stok aman</li>';
            return needRestock.map(p => `<li>🟡 <strong>${p.nama}</strong> (stok: ${p.stok}, terjual: ${p.terjual}/hari)</li>`).join('');
        }
        
        // Produk Slow Moving
        function getSlowMovingProducts() {
            const slowProducts = [];
            for(const [productName, productData] of productsMap) {
                const totalSales = productData.reduce((s, p) => s + p.terjual, 0);
                const avgSales = totalSales / productData.length;
                const lastStok = productData[productData.length - 1].stok;
                if(avgSales < 10 && lastStok > 30) {
                    slowProducts.push({ nama: productName, rataTerjual: avgSales.toFixed(1), stok: lastStok });
                }
            }
            if(slowProducts.length === 0) return '<li>✅ Tidak ada produk slow moving</li>';
            return slowProducts.map(p => `<li>📦 <strong>${p.nama}</strong> (terjual: ${p.rataTerjual}/hari, stok: ${p.stok})</li>`).join('');
        }
        
        // Update DOM
        const demandText = document.getElementById('demandPredictionText');
        if(demandText) {
            demandText.innerHTML = `
                <div style="margin-bottom: 16px;">
                    <div style="font-size: 48px; font-weight: 700; color: #3b82f6;">${forecastPerDay} <span style="font-size: 20px; color: #64748b;">unit/hari</span></div>
                    <div style="margin-top: 8px;">
                        <span style="background: ${overallTrend.percent >= 10 ? '#10b981' : (overallTrend.percent <= -10 ? '#ef4444' : '#f59e0b')}20; padding: 4px 12px; border-radius: 20px; color: ${overallTrend.percent >= 10 ? '#10b981' : (overallTrend.percent <= -10 ? '#ef4444' : '#f59e0b')};">
                            ${overallTrend.trend} ${overallTrend.percent >= 0 ? '+' : ''}${Math.abs(overallTrend.percent).toFixed(1)}%
                        </span>
                    </div>
                </div>
                <div style="display: flex; justify-content: space-between; gap: 12px;">
                    <div style="flex: 1; background: #3b82f6; border-radius: 12px; padding: 12px; text-align: center;">
                        <div style="font-size: 11px; color: #ffff;">HARI INI</div>
                        <div style="font-size: 18px; font-weight: 600;">${Math.max(1, Math.round(forecastPerDay * 0.8))}-${forecastPerDay}</div>
                    </div>
                    <div style="flex: 1; background: #3b82f6; border-radius: 12px; padding: 12px; text-align: center;">
                        <div style="font-size: 11px; color: #ffff;">7 HARI</div>
                        <div style="font-size: 18px; font-weight: 600;">${forecast7Days}</div>
                    </div>
                    <div style="flex: 1; background: #3b82f6; border-radius: 12px; padding: 12px; text-align: center;">
                        <div style="font-size: 11px; color: #ffff;">30 HARI</div>
                        <div style="font-size: 18px; font-weight: 600;">${forecastPerDay * 30}</div>
                    </div>
                </div>
            `;
        }
        
        // Update Restock List
        const restockList = document.getElementById('restockList');
        if(restockList) restockList.innerHTML = getProductsNeedRestock();
        
        // Update Slow Moving List
        const slowList = document.getElementById('slowMovingList');
        if(slowList) slowList.innerHTML = getSlowMovingProducts();
        
        // Update Counter
        if(document.getElementById('restockCount')) {
            const restockItems = getProductsNeedRestock();
            const count = (restockItems.match(/<li>/g) || []).length;
            document.getElementById('restockCount').innerText = count;
        }
        if(document.getElementById('slowMovingCount')) {
            const slowItems = getSlowMovingProducts();
            const count = (slowItems.match(/<li>/g) || []).length;
            document.getElementById('slowMovingCount').innerText = count;
        }
        
        // Update card prediksi demand
        const predictedElement = document.getElementById('predictedDemand');
        if(predictedElement) predictedElement.innerText = forecastPerDay;
        
        // Update grafik trend
        updateForecastTrendChart();
        
        console.log('Forecasting berhasil:', { forecastPerDay, forecast7Days });
        
    } catch(error) {
        console.error('Error di forecasting:', error);
        document.getElementById('demandPredictionText').innerHTML = 'Error: ' + error.message;
    }
}

// ========== GRAFIK FORECASTING TREND ==========
function updateForecastTrendChart() {
    console.log('updateForecastTrendChart dipanggil');
    
    if(!salesData.length) {
        console.log('Tidak ada data');
        return;
    }
    
    // Hitung prediksi 7 hari ke depan
    const avgDemand = salesData.reduce((s,p) => s + p.terjual, 0) / salesData.length;
    const forecast7Days = [];
    const labels = [];
    
    for(let i = 1; i <= 7; i++) {
        let forecast = avgDemand * (1 + (i * 0.03));
        forecast7Days.push(Math.round(forecast));
        labels.push(`H+${i}`);
    }
    
    const canvas = document.getElementById('forecastTrendChart');
    if(!canvas) {
        console.log('Canvas forecastTrendChart tidak ditemukan');
        return;
    }
    
    // Hapus chart lama jika ada (dengan cara aman)
    if(forecastTrendChart && typeof forecastTrendChart.destroy === 'function') {
        forecastTrendChart.destroy();
    }
    
    const ctx = canvas.getContext('2d');
    forecastTrendChart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: 'Prediksi Demand (WMA)',
                data: forecast7Days,
                borderColor: '#10b981',
                backgroundColor: 'rgba(16,185,129,0.1)',
                fill: true,
                tension: 0.4,
                pointRadius: 6,
                pointBackgroundColor: '#10b981',
                pointBorderColor: 'white',
                pointBorderWidth: 2
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { position: 'top' },
                tooltip: { callbacks: { label: (ctx) => `${ctx.raw} unit` } }
            },
            scales: {
                y: { 
                    title: { display: true, text: 'Demand (unit)' },
                    beginAtZero: true
                },
                x: { title: { display: true, text: 'Hari ke-' } }
            }
        }
    });
    
    console.log('Grafik berhasil dibuat:', forecast7Days);
}
    
    // ========== FUNGSI WEIGHTED MOVING AVERAGE ==========
    function weightedMovingAverage(data, weights = [0.5, 0.3, 0.2]) {
        // weights: bobot untuk 3 data terakhir (terbaru mendapat bobot tertinggi)
        if(data.length < 3) {
            // Jika data kurang dari 3, pakai simple average
            const avg = data.reduce((s,p) => s + p.terjual, 0) / data.length;
            return Math.round(avg);
        }
        
        const lastThree = data.slice(-3).map(p => p.terjual);
        let weightedSum = 0;
        let weightSum = 0;
        
        for(let i = 0; i < lastThree.length; i++) {
            weightedSum += lastThree[i] * weights[i];
            weightSum += weights[i];
        }
        
        return Math.round(weightedSum / weightSum);
    }
    
    // ========== FUNGSI DETEKSI TREND ==========
    function detectTrend(data) {
        if(data.length < 5) {
            return { trend: "➡️ Data Tidak Cukup", percent: 0, icon: "➡️" };
        }
        
        // Bandingkan 3 data awal vs 3 data akhir
        const firstThree = data.slice(0, 3).reduce((s,p) => s + p.terjual, 0);
        const lastThree = data.slice(-3).reduce((s,p) => s + p.terjual, 0);
        
        if(firstThree === 0) return { trend: "➡️ Stabil", percent: 0, icon: "➡️" };
        
        const percentChange = ((lastThree - firstThree) / firstThree) * 100;
        let trendIcon = "➡️";
        let trendText = "Stabil";
        
        if(percentChange > 30) { trendIcon = "🚀"; trendText = "Meningkat Pesat"; }
        else if(percentChange > 15) { trendIcon = "📈"; trendText = "Meningkat"; }
        else if(percentChange > 5) { trendIcon = "⬆️"; trendText = "Cenderung Naik"; }
        else if(percentChange < -30) { trendIcon = "💀"; trendText = "Menurun Drastis"; }
        else if(percentChange < -15) { trendIcon = "📉"; trendText = "Menurun"; }
        else if(percentChange < -5) { trendIcon = "⬇️"; trendText = "Cenderung Turun"; }
        
        return {
            trend: `${trendIcon} ${trendText}`,
            percent: percentChange.toFixed(1),
            icon: trendIcon
        };
    }
    
    // ========== FORECASTING KESELURUHAN ==========
    // Ambil semua data penjualan berurutan
    const allSales = sortedData.map(p => ({ terjual: p.terjual, tanggal: p.tanggal }));
    const forecastPerDay = weightedMovingAverage(allSales);
    const forecast7Days = forecastPerDay * 7;
    const overallTrend = detectTrend(allSales);
    
    // ========== PRODUK BUTUH RESTOK (Berdasarkan WMA) ==========
    function getProductsNeedRestock() {
        const productsWithForecast = [];
        
        for(const [productName, productData] of productsMap) {
            if(productData.length < 2) continue;
            
            const sortedProduct = [...productData].sort((a,b) => 
                new Date(a.tanggal) - new Date(b.tanggal)
            );
            
            const forecast = weightedMovingAverage(sortedProduct);
            const lastStock = sortedProduct[sortedProduct.length - 1].stok;
            const lastSales = sortedProduct[sortedProduct.length - 1].terjual;
            
            // Hitung rasio stok vs permintaan
            const stockToDemandRatio = lastStock / (lastSales || 1);
            
            productsWithForecast.push({
                nama: productName,
                stok: lastStock,
                terjual: lastSales,
                forecast: forecast,
                ratio: stockToDemandRatio
            });
        }
        
        // Filter produk yang butuh restock (stok < 15 ATAU rasio < 2)
        const needRestock = productsWithForecast.filter(p => 
            p.stok < 15 || (p.ratio < 2 && p.stok < 30)
        );
        
        // Urutkan berdasarkan paling kritis
        needRestock.sort((a,b) => a.stok - b.stok);
        
        if(needRestock.length === 0) {
            return '<li>✅ Semua stok aman</li>';
        }
        
        return needRestock.slice(0, 8).map(p => {
            let urgency = "";
            if(p.stok <= 0) urgency = "🔴 HABIS!";
            else if(p.stok < 5) urgency = "🔴 KRITIS!";
            else if(p.stok < 15) urgency = "🟡 MENIPIS";
            else urgency = "🟠 PERHATIAN";
            
            return `<li>${urgency} <strong>${p.nama}</strong> (stok: ${p.stok}, terjual: ${p.terjual}, prediksi: ${p.forecast})</li>`;
        }).join('');
    }
    
    // ========== PRODUK SLOW MOVING ==========
    function getSlowMovingProducts() {
        const slowProducts = [];
        
        for(const [productName, productData] of productsMap) {
            const totalSales = productData.reduce((s,p) => s + p.terjual, 0);
            const avgSales = totalSales / productData.length;
            const latestStock = productData[productData.length - 1].stok;
            
            // Slow moving if: avg sales < 10 AND stock > 30
            if(avgSales < 10 && latestStock > 30) {
                slowProducts.push({
                    nama: productName,
                    rataTerjual: avgSales.toFixed(1),
                    stok: latestStock
                });
            }
        }
        
        if(slowProducts.length === 0) {
            return '<li>✅ Tidak ada produk slow moving</li>';
        }
        
        return slowProducts.map(p => 
            `<li>📦 <strong>${p.nama}</strong> (rata-rata terjual: ${p.rataTerjual}/hari, stok: ${p.stok})</li>`
        ).join('');
    }
    
    // ========== UPDATE DOM ==========
    // Update text prediksi demand
    const trendResult = overallTrend;
    document.getElementById('demandPredictionText').innerHTML = `
        <div style="margin-bottom: 16px;">
            <div style="font-size: 48px; font-weight: 700; color: #3b82f6;">${forecastPerDay} <span style="font-size: 20px; color: #64748b;">unit/hari</span></div>
            <div style="margin-top: 8px;">
                <span class="${trendResult.percent >= 10 ? 'trend-up' : (trendResult.percent <= -10 ? 'trend-down' : 'trend-stable')}">
                    ${trendResult.trend} ${trendResult.percent >= 0 ? '+' : ''}${trendResult.percent}%
                </span>
            </div>
        </div>
        <div style="display: flex; justify-content: space-between; gap: 12px; margin-bottom: 16px;">
            <div style="flex: 1; background: rgba(255,255,255,0.05); border-radius: 12px; padding: 12px; text-align: center;">
                <div style="font-size: 11px; color: #64748b;">HARI INI</div>
                <div style="font-size: 20px; font-weight: 600;">${Math.round(forecastPerDay * 0.8)}-${forecastPerDay}</div>
            </div>
            <div style="flex: 1; background: rgba(255,255,255,0.05); border-radius: 12px; padding: 12px; text-align: center;">
                <div style="font-size: 11px; color: #64748b;">7 HARI</div>
                <div style="font-size: 20px; font-weight: 600;">${forecast7Days}</div>
            </div>
            <div style="flex: 1; background: rgba(255,255,255,0.05); border-radius: 12px; padding: 12px; text-align: center;">
                <div style="font-size: 11px; color: #64748b;">30 HARI</div>
                <div style="font-size: 20px; font-weight: 600;">${forecastPerDay * 30}</div>
            </div>
        </div>
    `;
    
    // Update counter restok dan slow moving
    const restockItems = document.getElementById('restockList').innerHTML;
    const restockCount = (restockItems.match(/<li>/g) || []).length;
    document.getElementById('restockCount').innerText = restockCount === 0 ? 0 : restockCount;
    
    const slowMovingItems = document.getElementById('slowMovingList').innerHTML;
    const slowMovingCount = (slowMovingItems.match(/<li>/g) || []).length;
    document.getElementById('slowMovingCount').innerText = slowMovingCount === 0 ? 0 : slowMovingCount;
    
    // Update grafik trend
    updateForecastTrendChart();


// ========== SMART STOCK ALERT (HANYA YANG KRITIS & WARNING) ==========
function renderStockAlerts() {
    const container = document.getElementById('stockAlertContainer');
    if(!container) return;
    
    const criticalProducts = salesData.filter(p => p.stok <= 40);   // KRITIS: ≤10 unit
    const warningProducts = salesData.filter(p => p.stok > 10 && p.stok <= 40);  // WARNING: 11-30 unit
    const safeProducts = salesData.filter(p => p.stok > 40);
    
    // Update badge counts
    if(document.getElementById('criticalCount')) document.getElementById('criticalCount').innerText = criticalProducts.length;
    if(document.getElementById('warningCount')) document.getElementById('warningCount').innerText = warningProducts.length;
    if(document.getElementById('safeCount')) document.getElementById('safeCount').innerText = safeProducts.length;
    
    container.innerHTML = '';
    
    // Jika tidak ada alert
    if(criticalProducts.length === 0 && warningProducts.length === 0) {
        container.innerHTML = `
            <div style="grid-column: 1/-1; text-align: center; padding: 40px; background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%); border-radius: 20px;">
                <i class="fas fa-check-circle" style="font-size: 64px; color: #10b981; margin-bottom: 16px;"></i>
                <h3 style="color: #065f46;">✅ Semua Stok Aman!</h3>
                <p style="color: #047857;">Tidak ada produk dengan stok menipis. Stok semua produk > 40 unit.</p>
                <div style="margin-top: 16px; padding: 10px; background: white; border-radius: 12px; display: inline-block;">
                    <span style="color: #10b981;">🎉 Selamat! Manajemen stok Anda sehat.</span>
                </div>
            </div>
        `;
        return;
    }
    
    // Urutkan produk
    criticalProducts.sort((a,b) => a.stok - b.stok);
    warningProducts.sort((a,b) => a.stok - b.stok);
    
    // Render card KRITIS
    criticalProducts.forEach(p => {
        const daysLeft = Math.ceil(p.stok / (p.terjual || 1));
        container.innerHTML += `
            <div style="background: #fef2f2; border-radius: 16px; padding: 16px; border-left: 5px solid #ef4444; box-shadow: 0 1px 3px rgba(0,0,0,0.1); transition: transform 0.2s;">
                <div style="display: flex; justify-content: space-between; align-items: start;">
                    <div>
                        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                            <span style="background: #ef4444; color: white; padding: 2px 10px; border-radius: 20px; font-size: 11px;">KRITIS</span>
                            <strong style="font-size: 16px;">🔥 ${p.nama_produk}</strong>
                        </div>
                        <div style="margin-top: 12px;">
                            <div style="display: flex; gap: 20px; margin-bottom: 8px;">
                                <div><span style="color: #64748b;">📦 Stok</span><br><strong style="color: #ef4444; font-size: 20px;">${p.stok}</strong> <span style="font-size: 12px;">unit</span></div>
                                <div><span style="color: #64748b;">📈 Terjual</span><br><strong>${p.terjual}</strong> <span style="font-size: 12px;">unit/hari</span></div>
                                <div><span style="color: #64748b;">⏰ Habis</span><br><strong style="color: #ef4444;">${daysLeft}</strong> <span style="font-size: 12px;">hari</span></div>
                            </div>
                        </div>
                    </div>
                    <i class="fas fa-skull-crosswalk" style="color: #ef4444; font-size: 28px;"></i>
                </div>
                <button onclick="triggerRestock('${p.nama_produk}')" style="margin-top: 12px; width: 100%; background: #ef4444; color: white; border: none; padding: 8px; border-radius: 10px; cursor: pointer;">
                    🚨 Restock Sekarang
                </button>
            </div>
        `;
    });
    
    // Render card WARNING
    warningProducts.forEach(p => {
        container.innerHTML += `
            <div style="background: #fffbeb; border-radius: 16px; padding: 16px; border-left: 5px solid #f59e0b; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">
                <div style="display: flex; justify-content: space-between; align-items: start;">
                    <div>
                        <div style="display: flex; align-items: center; gap: 8px; margin-bottom: 8px;">
                            <span style="background: #f59e0b; color: white; padding: 2px 10px; border-radius: 20px; font-size: 11px;">WARNING</span>
                            <strong style="font-size: 16px;">⚠️ ${p.nama_produk}</strong>
                        </div>
                        <div style="margin-top: 12px;">
                            <div style="display: flex; gap: 20px;">
                                <div><span style="color: #64748b;">📦 Stok</span><br><strong style="color: #f59e0b; font-size: 20px;">${p.stok}</strong> <span style="font-size: 12px;">unit</span></div>
                                <div><span style="color: #64748b;">📈 Terjual</span><br><strong>${p.terjual}</strong> <span style="font-size: 12px;">unit/hari</span></div>
                            </div>
                        </div>
                    </div>
                    <i class="fas fa-exclamation-triangle" style="color: #f59e0b; font-size: 28px;"></i>
                </div>
                <button onclick="alert('Rencanakan restock ${p.nama_produk} dalam 3-4 hari')" style="margin-top: 12px; width: 100%; background: #f59e0b; color: white; border: none; padding: 8px; border-radius: 10px; cursor: pointer;">
                    📋 Rencanakan Restock
                </button>
            </div>
        `;
    });
}

// ========== DETEKSI OVERSTOCK & STOCKOUT (TAMBAHKAN INI) ==========
function detectStockIssues() {
    if(!salesData.length) return { overstock: [], stockout: [], warning: [] };
    
    const recommendations = {
        overstock: [],
        stockout: [],
        warning: []
    };
    
    salesData.forEach(product => {
        const stok = product.stok;
        const terjual = product.terjual || 1; // Hindari pembagian 0
        const daysToEmpty = stok / terjual;
        
        // OVERSTOCK (Stok kebanyakan) - daysToEmpty > 60 hari
        if(daysToEmpty > 60) {
            let level = '', diskon = '', rekomendasi = '';
            if(daysToEmpty > 180) {
                level = '🔴 PARAH';
                diskon = '50-70%';
                rekomendasi = 'FLASH SALE! Diskon besar-besaran, bundle dengan produk laris';
            } else if(daysToEmpty > 90) {
                level = '🟠 BERAT';
                diskon = '30-40%';
                rekomendasi = 'Diskon besar + Beli 1 gratis 1 + Promosi media sosial';
            } else {
                level = '🟡 SEDANG';
                diskon = '15-25%';
                rekomendasi = 'Diskon bundling, paket hemat, beli banyak lebih murah';
            }
            recommendations.overstock.push({
                nama: product.nama_produk,
                stok: stok,
                terjual: terjual,
                daysToEmpty: Math.round(daysToEmpty),
                level: level,
                diskon: diskon,
                rekomendasi: rekomendasi,
                harga: product.harga_jual
            });
        }
        // STOCKOUT (Stok menipis) - daysToEmpty < 5 hari
        else if(daysToEmpty < 5) {
            let level = '', rekomendasi = '';
            if(daysToEmpty <= 1) {
                level = '🔴 DARURAT';
                rekomendasi = 'RESTOK SEKARANG JUGA! Pesan via supplier terdekat';
            } else if(daysToEmpty <= 3) {
                level = '🟠 KRITIS';
                rekomendasi = 'Restok dalam 1-2 hari, order ke supplier sekarang';
            } else {
                level = '🟡 WARNING';
                rekomendasi = 'Siapkan pesanan restok dalam 3-4 hari';
            }
            recommendations.stockout.push({
                nama: product.nama_produk,
                stok: stok,
                terjual: terjual,
                daysToEmpty: Math.round(daysToEmpty),
                level: level,
                rekomendasi: rekomendasi,
                harga: product.harga_jual
            });
        }
        // WARNING (Mulai perlu perhatian)
        else if(daysToEmpty < 15) {
            recommendations.warning.push({
                nama: product.nama_produk,
                stok: stok,
                terjual: terjual,
                daysToEmpty: Math.round(daysToEmpty)
            });
        }
    });
    
    return recommendations;
}

// ========== RENDER STOCK HEALTH (UNTUK TAB BARU) ==========
function renderStockHealthTab() {
    const issues = detectStockIssues();
    
    // Cek apakah elemen di tab Stock Health ada
    if(!document.getElementById('totalProductsHealth')) return;
    
    if(!issues || salesData.length === 0) {
        document.getElementById('totalProductsHealth').innerText = '0';
        document.getElementById('stockoutCountHealth').innerText = '0';
        document.getElementById('overstockCountHealth').innerText = '0';
        document.getElementById('globalRecommendation').innerHTML = 'Upload data terlebih dahulu untuk melihat analisis';
        return;
    }
    
    // Update summary cards
    document.getElementById('totalProductsHealth').innerText = salesData.length;
    document.getElementById('stockoutCountHealth').innerText = issues.stockout.length;
    document.getElementById('overstockCountHealth').innerText = issues.overstock.length;
    
    // Render Overstock Table
    const overstockBody = document.getElementById('overstockTableBody');
    if(overstockBody) {
        if(issues.overstock.length === 0) {
            overstockBody.innerHTML = '<tr><td colspan="6">✅ Tidak ada produk overstock</td></tr>';
        } else {
            overstockBody.innerHTML = issues.overstock.map(p => `
                <tr style="background: #f3f4ff;">
                    <td><strong>${p.nama}</strong></td>
                    <td>${p.stok.toLocaleString()} unit</td>
                    <td>${p.terjual} unit/hari</td>
                    <td style="color: #8b5cf6; font-weight: bold;">${p.daysToEmpty} hari</td>
                    <td>${p.diskon}</td>
                    <td>📉 Turunkan harga ${p.diskon} untuk percepat penjualan</td>
                </tr>
            `).join('');
        }
    }
    
    // Render Stockout Table
    const stockoutBody = document.getElementById('stockoutTableBody');
    if(stockoutBody) {
        if(issues.stockout.length === 0) {
            stockoutBody.innerHTML = '<tr><td colspan="6">✅ Tidak ada produk stockout</td></tr>';
        } else {
            stockoutBody.innerHTML = issues.stockout.map(p => `
                <tr style="background: #fef2f2;">
                    <td><strong>${p.nama}</strong></td>
                    <td style="color: #ef4444; font-weight: bold;">${p.stok} unit</td>
                    <td>${p.terjual} unit/hari</td>
                    <td style="color: #ef4444; font-weight: bold;">${p.daysToEmpty} hari</td>
                    <td>${p.level}</td>
                    <td>📈 Restok segera! Harga bisa naik 10-15%</td>
                </tr>
            `).join('');
        }
    }
    
    // Global Recommendation
    let globalReco = '';
    if(issues.overstock.length > 0 && issues.stockout.length > 0) {
        globalReco = '⚠️ Anda memiliki produk OVERSTOCK dan STOCKOUT. Fokus: Diskon produk overstock, restock produk stockout!';
    } else if(issues.overstock.length > 0) {
        globalReco = `📦 Terdapat ${issues.overstock.length} produk overstock. Segera lakukan promo diskon bundling untuk mengurangi stok!`;
    } else if(issues.stockout.length > 0) {
        globalReco = `🔥 Terdapat ${issues.stockout.length} produk stockout. Segera hubungi supplier untuk restock!`;
    } else {
        globalReco = '✅ Semua stok dalam kondisi sehat! Pertahankan manajemen stok Anda.';
    }
    
    const globalRecoElement = document.getElementById('globalRecommendation');
    if(globalRecoElement) globalRecoElement.innerHTML = globalReco;
}

// ===== TAMBAH JUGA FUNGSI UNTUK RESTOK DARI AI INSIGHTS =====
function triggerRestock(productName) {
    if(confirm(`Apakah Anda ingin memesan ${productName} sekarang?`)) {
        // Simulasi pemesanan
        alert(`✅ Pesanan untuk ${productName} telah dikirim ke supplier!`);
        
        // Optional: Update stok di UI (simulasi)
        const productIndex = salesData.findIndex(p => p.nama_produk === productName);
        if(productIndex !== -1) {
            salesData[productIndex].stok += 100; // tambah stok 100
            renderStockAlerts(); // refresh alert
            updateDashboardStats(); // refresh statistik
        }
    }
}

// ========== GRAFIK CHART.JS YANG DIPERBAIKI UNTUK DATA BESAR ==========
function updateAllCharts() {
    if(!salesData.length) return;
    
    // ===== SOLUSI 1: GROUPING PRODUK (TOP 10 SAJA) =====
    // Hitung total penjualan per produk
    const productSales = new Map();
    salesData.forEach(p => {
        const current = productSales.get(p.nama_produk) || { terjual: 0, stok: 0, harga: 0 };
        current.terjual += p.terjual;
        current.stok += p.stok;
        current.harga = p.harga_jual;
        productSales.set(p.nama_produk, current);
    });
    
    // Konversi ke array dan urutkan berdasarkan penjualan tertinggi
    const sortedProducts = Array.from(productSales.entries())
        .map(([nama, data]) => ({ nama, terjual: data.terjual, stok: data.stok, harga: data.harga }))
        .sort((a,b) => b.terjual - a.terjual);
    
    // Ambil TOP 12 produk terlaris untuk grafik (biar rapi)
    const topProducts = sortedProducts.slice(0, 12);
    const labels = topProducts.map(p => p.nama.length > 15 ? p.nama.slice(0,12) + '...' : p.nama);
    const penjualanData = topProducts.map(p => p.terjual);
    const stokData = topProducts.map(p => p.stok);
    const hargaData = topProducts.map(p => p.harga);
    
    // ===== GRAFIK 1: Penjualan Harian (TOP 12 PRODUK) =====
    if(salesChart) salesChart.destroy();
    const ctxSales = document.getElementById('salesChart').getContext('2d');
    salesChart = new Chart(ctxSales, { 
        type: 'bar', 
        data: { 
            labels: labels, 
            datasets: [{ 
                label: 'Total Terjual (unit)', 
                data: penjualanData, 
                backgroundColor: '#3b82f6',
                borderRadius: 8,
                barPercentage: 0.7
            }] 
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { position: 'top' },
                tooltip: { callbacks: { label: (ctx) => `${ctx.raw.toLocaleString()} unit` } }
            },
            scales: {
                x: { ticks: { rotation: -45, autoSkip: true, maxRotation: 45, minRotation: 45 } }
            }
        }
    });
    
    // ===== GRAFIK 2: Stok Produk (TOP 12 PRODUK) =====
    if(stockChart) stockChart.destroy();
    const ctxStock = document.getElementById('stockChart').getContext('2d');
    stockChart = new Chart(ctxStock, { 
        type: 'bar', 
        data: { 
            labels: labels, 
            datasets: [{ 
                label: 'Stok Tersisa (unit)', 
                data: stokData, 
                backgroundColor: '#10b981',
                borderRadius: 8
            }] 
        },
        options: {
            responsive: true,
            plugins: { legend: { position: 'top' } },
            scales: { x: { ticks: { rotation: -45, autoSkip: true } } }
        }
    });
    
    // ===== GRAFIK 3: Tren Harga vs Penjualan (SCATTER PLOT) =====
    if(priceTrendChart) priceTrendChart.destroy();
    const ctxPrice = document.getElementById('priceTrendChart').getContext('2d');
    priceTrendChart = new Chart(ctxPrice, { 
        type: 'scatter', 
        data: { 
            datasets: [{ 
                label: 'Harga vs Penjualan', 
                data: topProducts.map(p => ({ x: p.harga, y: p.terjual })),
                backgroundColor: '#f59e0b',
                pointRadius: 8,
                pointHoverRadius: 12
            }] 
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'top' },
                tooltip: { 
                    callbacks: { 
                        label: (ctx) => {
                            const product = topProducts[ctx.dataIndex];
                            return [`${product.nama}`, `Harga: Rp ${ctx.parsed.x.toLocaleString()}`, `Terjual: ${ctx.parsed.y.toLocaleString()} unit`];
                        }
                    }
                }
            },
            scales: {
                x: { title: { display: true, text: 'Harga Jual (Rp)' }, ticks: { callback: (v) => `Rp ${v.toLocaleString()}` } },
                y: { title: { display: true, text: 'Jumlah Terjual (unit)' } }
            }
        }
    });
    
    // ===== GRAFIK 4: Demand Forecasting (LINE CHART dengan trend) =====
    if(forecastChart) forecastChart.destroy();
    const ctxFore = document.getElementById('forecastChart').getContext('2d');
    
    // Hitung data time series per tanggal (agregat harian)
    const dailyData = new Map();
    salesData.forEach(p => {
        const date = p.tanggal;
        if(!dailyData.has(date)) dailyData.set(date, 0);
        dailyData.set(date, dailyData.get(date) + p.terjual);
    });
    
    // Urutkan berdasarkan tanggal
    const sortedDates = Array.from(dailyData.keys()).sort();
    const actualSales = sortedDates.map(d => dailyData.get(d));
    const dateLabels = sortedDates.map(d => d.slice(5)); // ambil MM-DD saja
    
    // Hitung trend line (linear regression)
    const n = actualSales.length;
    const indices = Array.from({length: n}, (_, i) => i);
    const sumX = indices.reduce((a,b) => a + b, 0);
    const sumY = actualSales.reduce((a,b) => a + b, 0);
    const sumXY = indices.reduce((sum, x, i) => sum + x * actualSales[i], 0);
    const sumXX = indices.reduce((sum, x) => sum + x * x, 0);
    const slope = (n * sumXY - sumX * sumY) / (n * sumXX - sumX * sumX);
    const intercept = (sumY - slope * sumX) / n;
    const trendLine = indices.map(x => slope * x + intercept);
    
    // Prediksi 7 hari ke depan
    const futureIndices = Array.from({length: n + 7}, (_, i) => i);
    const futureTrend = futureIndices.map(x => slope * x + intercept);
    const futureLabels = [...dateLabels, ...Array(7).fill('H+...')];
    
    forecastChart = new Chart(ctxFore, { 
        type: 'line', 
        data: { 
            labels: futureLabels, 
            datasets: [
                { 
                    label: 'Penjualan Aktual', 
                    data: [...actualSales, ...Array(7).fill(null)], 
                    borderColor: '#3b82f6', 
                    backgroundColor: 'rgba(59,130,246,0.1)',
                    fill: true,
                    tension: 0.3
                },
                { 
                    label: 'Trend Line (Regresi)', 
                    data: futureTrend, 
                    borderColor: '#ef4444', 
                    borderDash: [5, 5],
                    fill: false,
                    tension: 0
                }
            ] 
        },
        options: {
            responsive: true,
            plugins: {
                legend: { position: 'top' },
                tooltip: { callbacks: { label: (ctx) => `${ctx.raw?.toFixed(0) || 0} unit` } }
            },
            scales: {
                x: { title: { display: true, text: 'Waktu' }, ticks: { maxRotation: 45, autoSkip: true, maxTicksLimit: 10 } },
                y: { title: { display: true, text: 'Demand (unit)' } }
            }
        }
    });
}

// ========== GRAFIK FORECASTING TREND ==========
function updateForecastTrendChart() {
    if(!salesData.length) return;
    
    // Hitung prediksi 7 hari ke depan
    const avgDemand = salesData.reduce((s,p) => s + p.terjual, 0) / salesData.length;
    const forecast7Days = [];
    const labels = [];
    
    for(let i = 1; i <= 7; i++) {
        let forecast = avgDemand * (1 + (i * 0.03)); // Asumsi naik 3% per hari
        forecast7Days.push(Math.round(forecast));
        labels.push(`H+${i}`);
    }
    
    const ctx = document.getElementById('forecastTrendChart');
    if(!ctx) return;
    
    if(forecastTrendChart) forecastTrendChart.destroy();
    
    forecastTrendChart = new Chart(ctx.getContext('2d'), {
        type: 'line',
        data: {
            labels: labels,
            datasets: [
                {
                    label: 'Prediksi Demand (WMA)',
                    data: forecast7Days,
                    borderColor: '#3b82f6',
                    backgroundColor: 'rgba(59, 130, 246, 0.1)',
                    fill: true,
                    tension: 0.4,
                    pointRadius: 6,
                    pointBackgroundColor: '#3b82f6',
                    pointBorderColor: 'white',
                    pointBorderWidth: 2
                }
            ]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: { position: 'top' },
                tooltip: {
                    callbacks: {
                        label: (ctx) => `${ctx.raw} unit`
                    }
                }
            },
            scales: {
                y: {
                    title: { display: true, text: 'Demand (unit)', color: '#64748b' },
                    grid: { color: '#e2e8f0' }
                },
                x: {
                    title: { display: true, text: 'Hari ke-', color: '#64748b' },
                    grid: { display: false }
                }
            }
        }
    });
}

// ========== AI INSIGHTS GENERATE ==========
// ========== AI INSIGHTS CLEAN MODERN VERSION ==========
function renderAIInsights() {
    if(!salesData.length) {
        document.getElementById('healthScoreClean').innerHTML = '--';
        document.getElementById('growthRateClean').innerHTML = '--';
        document.getElementById('turnoverClean').innerHTML = '--';
        document.getElementById('revenueClean').innerHTML = 'Rp 0';
        document.getElementById('executiveSummaryClean').innerHTML = 'Upload data terlebih dahulu untuk melihat ringkasan bisnis Anda.';
        document.getElementById('topProductsClean').innerHTML = '<div style="color:#94a3b8;text-align:center;padding:20px;">Belum ada data</div>';
        document.getElementById('bottomProductsClean').innerHTML = '<div style="color:#94a3b8;text-align:center;padding:20px;">Belum ada data</div>';
        document.getElementById('recommendationsClean').innerHTML = '<div style="color:#94a3b8;text-align:center;padding:20px;">Upload data untuk melihat rekomendasi AI</div>';
        document.getElementById('performanceClean').innerHTML = '<div style="color:#94a3b8;text-align:center;padding:20px;">Belum ada data</div>';
        return;
    }

    // Hitung metrics
    const totalProducts = salesData.length;
    const totalStock = salesData.reduce((sum, p) => sum + p.stok, 0);
    const totalSales = salesData.reduce((sum, p) => sum + (p.harga_jual * p.terjual), 0);
    const avgDaily = salesData.reduce((sum, p) => sum + p.terjual, 0) / totalProducts;
    const turnover = salesData.reduce((sum, p) => sum + (p.stok / (p.terjual || 1)), 0) / totalProducts;

    // Growth rate
    const sortedData = [...salesData].sort((a,b) => new Date(a.tanggal) - new Date(b.tanggal));
    const firstAvg = sortedData.slice(0, 3).reduce((s,p) => s + p.terjual, 0) / 3;
    const lastAvg = sortedData.slice(-3).reduce((s,p) => s + p.terjual, 0) / 3;
    const growthRate = firstAvg === 0 ? 0 : ((lastAvg - firstAvg) / firstAvg) * 100;

    // Top & Bottom products
    const topProducts = [...salesData].sort((a,b) => b.terjual - a.terjual).slice(0, 5);
    const bottomProducts = [...salesData].filter(p => p.terjual > 0).sort((a,b) => a.terjual - b.terjual).slice(0, 5);

    // Health score
    let healthScore = 100;
    const lowStockCount = salesData.filter(p => p.stok < 30).length;
    const overstockCount = salesData.filter(p => p.stok > 100 && p.terjual < 10).length;
    if(lowStockCount > 3) healthScore -= 15;
    if(overstockCount > 2) healthScore -= 20;
    if(growthRate < -10) healthScore -= 15;
    if(turnover > 60) healthScore -= 10;
    healthScore = Math.max(0, Math.min(100, healthScore));

    let healthStatus = healthScore >= 80 ? 'Excellent' : (healthScore >= 60 ? 'Fair' : 'Critical');
    let healthColor = healthScore >= 80 ? '#10b981' : (healthScore >= 60 ? '#f59e0b' : '#ef4444');

    // Update DOM - Cards
    document.getElementById('healthScoreClean').innerHTML = healthScore;
    document.getElementById('healthStatusClean').innerHTML = healthStatus;
    document.getElementById('healthBarClean').style.width = healthScore + '%';
    document.getElementById('healthBarClean').style.background = healthColor;
    document.getElementById('growthRateClean').innerHTML = `${growthRate >= 0 ? '+' : ''}${growthRate.toFixed(1)}%`;
    document.getElementById('growthRateClean').style.color = growthRate >= 0 ? '#10b981' : '#ef4444';
    document.getElementById('turnoverClean').innerHTML = Math.round(turnover);
    document.getElementById('revenueClean').innerHTML = `Rp ${Math.round(totalSales / 1000)}K`;

    // Executive Summary
    document.getElementById('executiveSummaryClean').innerHTML = `
        <strong>📈 ${salesData.length} products</strong> generating <strong>Rp ${Math.round(totalSales).toLocaleString()}</strong> in sales. 
        Average daily demand: <strong>${Math.round(avgDaily)} units</strong>. 
        ${growthRate > 5 ? '📈 Positive growth trend (+' + growthRate.toFixed(1) + '%) - consider expanding stock.' : 
          growthRate < -5 ? '📉 Sales declining (' + growthRate.toFixed(1) + '%) - review pricing strategy.' : 
          '📊 Stable trend - maintain current strategy.'}
        Stock lasts <strong>${Math.round(turnover)} days</strong> on average.
    `;

    // Top Products
    let topHtml = '';
    const rankIcons = ['🥇', '🥈', '🥉'];
    topProducts.forEach((p, i) => {
        const rank = i < 3 ? rankIcons[i] : `#${i+1}`;
        const percentAbove = ((p.terjual / avgDaily) - 1) * 100;
        topHtml += `
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid #f1f5f9;">
                <div style="display: flex; align-items: center; gap: 12px;">
                    <span style="font-size: 24px;">${rank}</span>
                    <div>
                        <div style="font-weight: 600; color: #1e293b;">${p.nama_produk}</div>
                        <div style="font-size: 11px; color: #64748b;">${p.terjual} units/day | Stock: ${p.stok}</div>
                    </div>
                </div>
                <div style="background: #d1fae5; padding: 4px 12px; border-radius: 30px; font-size: 12px; color: #065f46;">+${Math.round(percentAbove)}%</div>
            </div>
        `;
    });
    document.getElementById('topProductsClean').innerHTML = topHtml;

    // Bottom Products
    let bottomHtml = '';
    bottomProducts.forEach(p => {
        const percentBelow = ((avgDaily - p.terjual) / avgDaily) * 100;
        bottomHtml += `
            <div style="display: flex; align-items: center; justify-content: space-between; padding: 12px 0; border-bottom: 1px solid #f1f5f9;">
                <div>
                    <div style="font-weight: 600; color: #1e293b;">${p.nama_produk}</div>
                    <div style="font-size: 11px; color: #64748b;">${p.terjual} units/day | Stock: ${p.stok}</div>
                </div>
                <div style="background: #fee2e2; padding: 4px 12px; border-radius: 30px; font-size: 12px; color: #991b1b;">-${Math.round(percentBelow)}%</div>
            </div>
        `;
    });
    document.getElementById('bottomProductsClean').innerHTML = bottomHtml || '<div style="text-align:center;padding:20px;color:#94a3b8;">All products performing well!</div>';

    // AI Recommendations
    let recommendations = [];
    if(lowStockCount > 0) recommendations.push({ icon: '⚠️', bg: '#fef2f2', color: '#ef4444', text: `${lowStockCount} product(s) low on stock. Restock: ${salesData.filter(p => p.stok < 30).slice(0,3).map(p => p.nama_produk).join(', ')}` });
    if(overstockCount > 0) recommendations.push({ icon: '📦', bg: '#f3f4ff', color: '#8b5cf6', text: `${overstockCount} product(s) overstocked. Run discounts: ${salesData.filter(p => p.stok > 100 && p.terjual < 10).slice(0,3).map(p => p.nama_produk).join(', ')}` });
    if(growthRate > 15) recommendations.push({ icon: '🚀', bg: '#ecfdf5', color: '#10b981', text: `Rapid growth ${growthRate.toFixed(1)}%! Increase stock by 30%.` });
    else if(growthRate < -15) recommendations.push({ icon: '📉', bg: '#fef2f2', color: '#ef4444', text: `Sales dropping ${Math.abs(growthRate).toFixed(1)}%. Consider promotions or price adjustment.` });
    if(turnover > 50) recommendations.push({ icon: '⏱️', bg: '#fffbeb', color: '#f59e0b', text: `Stock sits ${Math.round(turnover)} days. Reduce slow-moving inventory.` });
    if(topProducts[0] && topProducts[0].stok < 30) recommendations.push({ icon: '🔥', bg: '#fef2f2', color: '#ef4444', text: `Best seller "${topProducts[0].nama_produk}" running low! Restock immediately.` });
    if(recommendations.length === 0) recommendations.push({ icon: '✅', bg: '#ecfdf5', color: '#10b981', text: 'All metrics healthy! Your inventory management is on point.' });

    let recHtml = '';
    recommendations.forEach(rec => {
        recHtml += `
            <div style="background: ${rec.bg}; padding: 14px 18px; border-radius: 14px; display: flex; align-items: center; gap: 14px; border-left: 3px solid ${rec.color};">
                <span style="font-size: 22px;">${rec.icon}</span>
                <span style="flex: 1; color: #1e293b; font-size: 14px;">${rec.text}</span>
            </div>
        `;
    });
    document.getElementById('recommendationsClean').innerHTML = recHtml;

    // Performance vs Target
    const targetPerDay = avgDaily * 1.1;
    const perfData = [...salesData].sort((a,b) => b.terjual - a.terjual).slice(0, 8);
    let perfHtml = '';
    perfData.forEach(p => {
        const percent = Math.min(100, (p.terjual / targetPerDay) * 100);
        let barColor = percent >= 100 ? '#10b981' : (percent >= 70 ? '#f59e0b' : '#ef4444');
        perfHtml += `
            <div>
                <div style="display: flex; justify-content: space-between; margin-bottom: 6px;">
                    <span style="font-size: 13px; font-weight: 500; color: #1e293b;">${p.nama_produk}</span>
                    <span style="font-size: 12px; color: #64748b;">${p.terjual} / ${Math.round(targetPerDay)}</span>
                </div>
                <div style="background: #f1f5f9; border-radius: 10px; height: 8px; overflow: hidden;">
                    <div style="width: ${percent}%; background: ${barColor}; height: 8px; border-radius: 10px;"></div>
                </div>
            </div>
        `;
    });
    document.getElementById('performanceClean').innerHTML = perfHtml;
}

// ========== CHATBOT SIMULASI ==========
function sendChatMessage() {
    const input = document.getElementById('chatInput');
    const msg = input.value.trim();
    if(!msg) return;
    const chatDiv = document.getElementById('chatMessages');
    chatDiv.innerHTML += `<div class="user-message">${msg}</div>`;
    input.value = '';
    setTimeout(() => {
        let botReply = "Maaf, data belum tersedia. Silakan upload file terlebih dahulu.";
        if(salesData.length) {
            const lowerMsg = msg.toLowerCase();
            if(lowerMsg.includes('cabai')) {
                const cabai = salesData.find(p => p.nama_produk.toLowerCase().includes('cabai'));
                botReply = cabai ? `Cabai tersedia ${cabai.stok} kg dengan harga Rp ${cabai.harga_jual.toLocaleString()}/kg` : "Cabai tidak ditemukan dalam data.";
            } else if(lowerMsg.includes('stok')) {
                botReply = `Total produk dengan stok menipis: ${salesData.filter(p => p.stok < 10).length} item.`;
            } else if(lowerMsg.includes('harga')) {
                botReply = `Rekomendasi AI: ${document.getElementById('priceReco').innerText}`;
            } else {
                botReply = `Saya bisa bantu info stok, harga, dan rekomendasi. Coba tanya 'stok cabai' atau 'rekomendasi harga'.`;
            }
        }
        chatDiv.innerHTML += `<div class="bot-message">🤖 ${botReply}</div>`;
        chatDiv.scrollTop = chatDiv.scrollHeight;
    }, 500);
}

// ========== AI PRICE WAR SIMULATOR ==========

// Inisialisasi dropdown produk untuk price simulator
// ========== FUNGSI UNTUK PRICE WAR SIMULATOR ==========
function initPriceWarSimulator() {
    const selectEl = document.getElementById('priceWarProductSelect');
    if(!selectEl) return;
    
    selectEl.innerHTML = '<option value="">-- Pilih Produk --</option>';
    salesData.forEach(product => {
        selectEl.innerHTML += `<option value="${product.nama_produk}">${product.nama_produk} (Rp ${product.harga_jual.toLocaleString()} | laku ${product.terjual}/hari)</option>`;
    });
    
    selectEl.addEventListener('change', (e) => {
        if(e.target.value) {
            runPriceWarSimulation(e.target.value);
        }
    });
}

// ========== FUNGSI UNTUK DROPDOWN FORECASTING 12 BULAN ==========
function populateYearlyProductDropdown() {
    const selectEl = document.getElementById('yearlyProductSelect');
    if(!selectEl) return;
    
    selectEl.innerHTML = '<option value="all">📊 Semua Produk (Rata-rata)</option>';
    salesData.forEach(product => {
        selectEl.innerHTML += `<option value="${product.nama_produk}">📦 ${product.nama_produk}</option>`;
    });
    
    // Tambahkan event listener
    selectEl.addEventListener('change', function() {
        updateYearlyForecast();
    });
}

// Core simulation engine
function runPriceWarSimulation(productName) {
    const product = salesData.find(p => p.nama_produk === productName);
    if(!product) return;
    
    // Sembunyikan loading, tampilkan hasil
    document.getElementById('simulationLoading').style.display = 'none';
    document.getElementById('simulationResult').style.display = 'block';
    document.getElementById('selectedProductDisplay').innerHTML = `🔫 ${product.nama_produk}`;
    
    // Hitung elastisitas harga dari data historis (simulasi berdasarkan pola)
    const currentPrice = product.harga_jual;
    const currentDemand = product.terjual;
    const currentProfit = currentPrice * currentDemand * 0.25; // margin 25%
    
    // Estimasi elastisitas berdasarkan demand dan stok
    let elasticity = -1.2; // default elastis (demand sensitif ke harga)
    if(product.stok < 20 && product.terjual > 50) elasticity = -0.6; // inelastis (demand tetap walau harga naik)
    if(product.stok > 100 && product.terjual < 10) elasticity = -2.5; // sangat elastis (diskon besar pengaruh besar)
    
    // Tampilkan data real produk
    document.getElementById('currentProductStats').innerHTML = `
        <div style="text-align: center; padding: 12px; background: #f8fafc; border-radius: 12px;">
            <div style="font-size: 11px; color: #64748b;">Harga Saat Ini</div>
            <div style="font-size: 24px; font-weight: 700;">Rp ${currentPrice.toLocaleString()}</div>
        </div>
        <div style="text-align: center; padding: 12px; background: #f8fafc; border-radius: 12px;">
            <div style="font-size: 11px; color: #64748b;">Demand Per Hari</div>
            <div style="font-size: 24px; font-weight: 700;">${currentDemand} unit</div>
        </div>
        <div style="text-align: center; padding: 12px; background: #f8fafc; border-radius: 12px;">
            <div style="font-size: 11px; color: #64748b;">Profit Per Hari</div>
            <div style="font-size: 24px; font-weight: 700;">Rp ${Math.round(currentProfit).toLocaleString()}</div>
        </div>
    `;
    
    // Skenario 1: Turun 12%
    const downPercent = 12;
    const downPrice = currentPrice * (1 - downPercent/100);
    const downDemandChange = Math.abs(elasticity) * downPercent;
    const downDemand = Math.round(currentDemand * (1 + downDemandChange/100));
    const downProfit = downPrice * downDemand * 0.25;
    const downProfitDiff = ((downProfit - currentProfit) / currentProfit * 100);
    
    document.getElementById('scenarioDownDetail').innerHTML = `
        <div style="font-size: 24px; font-weight: 700; color: #22c55e;">Rp ${Math.round(downPrice).toLocaleString()}</div>
        <div style="font-size: 13px; margin-top: 8px;">📈 Demand: <strong class="${downDemand > currentDemand ? 'simulation-value-up' : 'simulation-value-down'}">${downDemand} unit</strong> (${downDemandChange > 0 ? '+' : ''}${downDemandChange}%)</div>
        <div style="font-size: 13px;">💰 Profit: <strong class="${downProfitDiff >= 0 ? 'simulation-value-up' : 'simulation-value-down'}">Rp ${Math.round(downProfit).toLocaleString()}</strong> (${downProfitDiff >= 0 ? '+' : ''}${downProfitDiff.toFixed(1)}%)</div>
        <div style="margin-top: 12px; padding: 8px; background: white; border-radius: 10px; font-size: 11px;">
            ${downProfitDiff >= 0 ? '✅ Profit meningkat! Cocok untuk jangka panjang' : '⚠️ Profit turun, tapi market share naik'}
        </div>
    `;
    
    // Skenario 2: Stabil (0%)
    document.getElementById('scenarioStableDetail').innerHTML = `
        <div style="font-size: 24px; font-weight: 700; color: #3b82f6;">Rp ${currentPrice.toLocaleString()}</div>
        <div style="font-size: 13px; margin-top: 8px;">📈 Demand: <strong>${currentDemand} unit</strong> (0%)</div>
        <div style="font-size: 13px;">💰 Profit: <strong>Rp ${Math.round(currentProfit).toLocaleString()}</strong> (0%)</div>
        <div style="margin-top: 12px; padding: 8px; background: white; border-radius: 10px; font-size: 11px;">
            🟡 Posisi aman tanpa risiko. Pertahankan strategi.
        </div>
    `;
    
    // Skenario 3: Naik 10%
    const upPercent = 10;
    const upPrice = currentPrice * (1 + upPercent/100);
    const upDemandChange = Math.abs(elasticity) * upPercent;
    const upDemand = Math.round(currentDemand * (1 - upDemandChange/100));
    const upProfit = upPrice * upDemand * 0.25;
    const upProfitDiff = ((upProfit - currentProfit) / currentProfit * 100);
    
    document.getElementById('scenarioUpDetail').innerHTML = `
        <div style="font-size: 24px; font-weight: 700; color: #f97316;">Rp ${Math.round(upPrice).toLocaleString()}</div>
        <div style="font-size: 13px; margin-top: 8px;">📈 Demand: <strong class="${upDemand < currentDemand ? 'simulation-value-down' : 'simulation-value-up'}">${upDemand} unit</strong> (${upDemandChange > 0 ? '-' : ''}${upDemandChange}%)</div>
        <div style="font-size: 13px;">💰 Profit: <strong class="${upProfitDiff >= 0 ? 'simulation-value-up' : 'simulation-value-down'}">Rp ${Math.round(upProfit).toLocaleString()}</strong> (${upProfitDiff >= 0 ? '+' : ''}${upProfitDiff.toFixed(1)}%)</div>
        <div style="margin-top: 12px; padding: 8px; background: white; border-radius: 10px; font-size: 11px;">
            ${upProfitDiff >= 0 ? '✅ Profit meningkat! Cocok untuk stok terbatas' : '⚠️ Profit turun, demand turun signifikan'}
        </div>
    `;
    
    // ========== AI RECOMMENDATION (Paling Keren) ==========
    let bestScenario = '';
    let bestReason = '';
    let bestColor = '';
    let bestAction = '';
    
    // Bandingkan profit tertinggi
    const profits = [
        { name: 'TURUN', profit: downProfit, change: downProfitDiff, price: downPrice, demand: downDemand },
        { name: 'STABIL', profit: currentProfit, change: 0, price: currentPrice, demand: currentDemand },
        { name: 'NAIK', profit: upProfit, change: upProfitDiff, price: upPrice, demand: upDemand }
    ];
    
    profits.sort((a,b) => b.profit - a.profit);
    const winner = profits[0];
    
    if(winner.name === 'TURUN') {
        bestColor = '#22c55e';
        bestReason = `💎 Turunkan harga ${downPercent}% ke Rp ${Math.round(downPrice).toLocaleString()}`;
        bestAction = `Strategi ini akan meningkatkan demand ${downDemandChange}% menjadi ${downDemand} unit/hari, dan profit akan ${downProfitDiff >= 0 ? 'naik' : 'turun'} ${Math.abs(downProfitDiff).toFixed(1)}%. ${downProfitDiff >= 0 ? 'Ini langkah agresif yang cerdas untuk merebut pangsa pasar dari kompetitor!' : 'Meski profit sedikit turun, market share Anda akan melonjak drastis!'}`;
    } else if(winner.name === 'NAIK') {
        bestColor = '#f97316';
        bestReason = `💰 Naikkan harga ${upPercent}% ke Rp ${Math.round(upPrice).toLocaleString()}`;
        bestAction = `Strategi ini memanfaatkan demand yang tinggi (${currentDemand} unit/hari) dengan stok terbatas (${product.stok} unit). Profit akan ${upProfitDiff >= 0 ? 'naik' : 'turun'} ${Math.abs(upProfitDiff).toFixed(1)}%. ${upProfitDiff >= 0 ? 'Keputusan tepat! Stok terbatas, maksimalkan keuntungan!' : 'Perlu dipertimbangkan ulang, profit malah turun.'}`;
    } else {
        bestColor = '#3b82f6';
        bestReason = `⚖️ Pertahankan harga di Rp ${currentPrice.toLocaleString()}`;
        bestAction = `Saat ini adalah posisi teraman. Demand stabil di ${currentDemand} unit/hari dengan profit Rp ${Math.round(currentProfit).toLocaleString()}/hari. ${product.stok < 20 ? 'Stok Anda terbatas, bersiap untuk naikkan harga bertahap.' : 'Stok masih aman, pantau terus pergerakan pasar.'}`;
    }
    
    document.getElementById('aiRecommendation').innerHTML = `
        <div style="background: ${bestColor}20; border-radius: 16px; padding: 20px;">
            <div style="display: flex; align-items: center; gap: 16px; flex-wrap: wrap; justify-content: space-between;">
                <div>
                    <div style="font-size: 14px; opacity: 0.8;">🎯 REKOMENDASI TINDAKAN</div>
                    <div style="font-size: 20px; font-weight: 700; margin-top: 4px;">${bestReason}</div>
                    <div style="margin-top: 12px; font-size: 14px; opacity: 0.9;">${bestAction}</div>
                </div>
                <div style="text-align: center; background: rgba(255,255,255,0.1); padding: 12px 20px; border-radius: 16px;">
                    <div style="font-size: 11px;">PREDIKSI PROFIT</div>
                    <div style="font-size: 28px; font-weight: 700;">${winner.change >= 0 ? '+' : ''}${winner.change.toFixed(1)}%</div>
                </div>
            </div>
        </div>
        <div style="margin-top: 16px; padding: 12px; background: rgba(255,255,255,0.05); border-radius: 12px; font-size: 12px;">
            <i class="fas fa-chart-line"></i> Berdasarkan analisis elastisitas harga (${elasticity.toFixed(1)}), produk ini memiliki sensitivitas ${Math.abs(elasticity) > 1 ? 'tinggi' : 'rendah'} terhadap perubahan harga.
        </div>
    `;
    
    // Tambahkan event click ke card scenario untuk highlight
    document.querySelectorAll('#scenarioDown, #scenarioStable, #scenarioUp').forEach(el => {
        el.addEventListener('click', () => {
            document.querySelectorAll('#scenarioDown, #scenarioStable, #scenarioUp').forEach(e => {
                e.classList.remove('scenario-active');
            });
            el.classList.add('scenario-active');
            
            // Scroll ke rekomendasi
            document.getElementById('aiRecommendationBox').scrollIntoView({ behavior: 'smooth', block: 'center' });
        });
    });
}

// Call function setelah data diproses
// Tambahkan di dalam processData() setelah renderAIInsights():
// initPriceWarSimulator();

// Update tab navigation untuk include pricewar
document.querySelectorAll('.nav-item').forEach(item => {
    if(item.dataset.tab === 'pricewar') {
        // Sudah ada di HTML? Jika belum, tambahkan di sidebar
    }
});

// ========== PROFIL PERUSAHAAN & KINERJA ==========

// Pastikan DOM sudah siap
document.addEventListener('DOMContentLoaded', function() {
    const profileAvatar = document.getElementById('profileAvatar');
    const profileModal = document.getElementById('companyProfileModal');
    const closeProfileBtn = document.getElementById('closeProfileModal');
    
    // Debug: cek apakah elemen ditemukan
    console.log('profileAvatar:', profileAvatar);
    console.log('profileModal:', profileModal);
    
    if(profileAvatar) {
        profileAvatar.addEventListener('click', function(e) {
            e.stopPropagation();
            console.log('Avatar diklik!');
            if(profileModal) {
                updatePerformanceStats(); // update data dulu
                profileModal.style.display = 'flex';
            } else {
                alert('Modal tidak ditemukan!');
            }
        });
    } else {
        console.log('Avatar dengan id "profileAvatar" tidak ditemukan');
    }
    
    if(closeProfileBtn) {
        closeProfileBtn.addEventListener('click', function() {
            if(profileModal) {
                profileModal.style.display = 'none';
            }
        });
    }
    
    // Tutup modal jika klik di luar modal
    window.addEventListener('click', function(e) {
        if(e.target === profileModal) {
            profileModal.style.display = 'none';
        }
    });
});

// Update statistik kinerja berdasarkan salesData
function updatePerformanceStats() {
    if(!salesData || salesData.length === 0) {
        const perfElements = ['perfTotalProducts', 'perfTotalStock', 'perfTotalSales', 'perfAvgDaily'];
        perfElements.forEach(id => {
            const el = document.getElementById(id);
            if(el) el.innerText = id === 'perfTotalSales' ? 'Rp 0' : '0';
        });
        const topEl = document.getElementById('perfTopProduct');
        if(topEl) topEl.innerHTML = '-';
        const lowestEl = document.getElementById('perfLowestProduct');
        if(lowestEl) lowestEl.innerHTML = '-';
        const aiEl = document.getElementById('aiAssessment');
        if(aiEl) aiEl.innerHTML = 'Upload file Excel terlebih dahulu untuk melihat analisis kinerja.';
        return;
    }
    
    // Hitung total produk unik
    const totalProducts = salesData.length;
    
    // Hitung total stok
    const totalStock = salesData.reduce((sum, p) => sum + (p.stok || 0), 0);
    
    // Hitung total penjualan
    const totalSales = salesData.reduce((sum, p) => sum + ((p.harga_jual || 0) * (p.terjual || 0)), 0);
    
    // Hitung rata-rata penjualan per hari
    const avgDaily = Math.round(salesData.reduce((sum, p) => sum + (p.terjual || 0), 0) / totalProducts);
    
    // Produk terlaris
    const topProduct = salesData.reduce((max, p) => (p.terjual || 0) > (max.terjual || 0) ? p : max, {nama_produk: '-', terjual: 0});
    
    // Produk terendah
    const lowestProduct = salesData.filter(p => (p.terjual || 0) > 0).reduce((min, p) => (p.terjual || 0) < (min.terjual || Infinity) ? p : min, {nama_produk: '-', terjual: 0});
    
    // Update DOM dengan aman (cek apakah element ada)
    const totalProductsEl = document.getElementById('perfTotalProducts');
    if(totalProductsEl) totalProductsEl.innerText = totalProducts;
    
    const totalStockEl = document.getElementById('perfTotalStock');
    if(totalStockEl) totalStockEl.innerText = totalStock.toLocaleString();
    
    const totalSalesEl = document.getElementById('perfTotalSales');
    if(totalSalesEl) totalSalesEl.innerText = `Rp ${Math.round(totalSales).toLocaleString()}`;
    
    const avgDailyEl = document.getElementById('perfAvgDaily');
    if(avgDailyEl) avgDailyEl.innerText = avgDaily;
    
    const topProductEl = document.getElementById('perfTopProduct');
    if(topProductEl) topProductEl.innerHTML = `${topProduct.nama_produk} <span style="color: #10b981;">(${topProduct.terjual} unit)</span>`;
    
    const lowestProductEl = document.getElementById('perfLowestProduct');
    if(lowestProductEl) lowestProductEl.innerHTML = `${lowestProduct.nama_produk} <span style="color: #ef4444;">(${lowestProduct.terjual} unit)</span>`;
    
    // AI Assessment
    const stockHealth = salesData.filter(p => (p.stok || 0) < 15).length;
    const overstockHealth = salesData.filter(p => (p.stok || 0) > 100 && (p.terjual || 0) < 10).length;
    const highDemand = salesData.filter(p => (p.terjual || 0) > 40).length;
    
    let score = 95;
    let assessment = '';
    
    if(stockHealth === 0 && overstockHealth === 0) {
        score = 95;
        assessment = '🎉 Luar biasa! Manajemen stok Anda sangat sehat. Tidak ada produk stok menipis atau overstock. Pertahankan strategi ini!';
    } else if(stockHealth <= 2 && overstockHealth <= 1) {
        score = 80;
        assessment = '✅ Cukup baik. Hanya sedikit produk yang perlu perhatian. Segera restok produk stok menipis.';
    } else if(stockHealth <= 5) {
        score = 65;
        assessment = '⚠️ Perlu peningkatan. Beberapa produk stok menipis. Saran: review jadwal restok Anda.';
    } else {
        score = 45;
        assessment = '🔴 Butuh perhatian serius! Banyak produk stok menipis atau overstock. Segera evaluasi strategi pembelian dan restok.';
    }
    
    let insight = highDemand > 3 ? ` 📈 Plus: ${highDemand} produk memiliki demand tinggi - ini peluang besar!` : '';
    
    const aiEl = document.getElementById('aiAssessment');
    if(aiEl) {
        aiEl.innerHTML = `
            <div style="margin-bottom: 12px;">
                <span style="background: ${score >= 80 ? '#10b981' : (score >= 60 ? '#f59e0b' : '#ef4444')}; padding: 4px 12px; border-radius: 20px; font-size: 12px;">
                    Skor Kesehatan: ${score}/100
                </span>
            </div>
            <p style="margin: 0;">${assessment}${insight}</p>
        `;
    }
}

// ========== PREDIKSI 12 BULAN ==========

function updateYearlyForecast() {
    console.log('updateYearlyForecast dipanggil');
    
    if(!salesData.length) {
        console.log('Tidak ada data');
        return;
    }
    
    // Ambil nilai dari dropdown
    const productSelect = document.getElementById('yearlyProductSelect');
    let selectedProduct = 'all';
    if(productSelect) {
        selectedProduct = productSelect.value;
    }
    console.log('Produk dipilih:', selectedProduct);
    
    // Hitung rata-rata demand
    let avgDemand;
    let productName = 'Semua Produk';
    
    if(selectedProduct === 'all') {
        const totalTerjual = salesData.reduce((sum, p) => sum + p.terjual, 0);
        avgDemand = totalTerjual / salesData.length;
        productName = 'Semua Produk (Rata-rata)';
    } else {
        const product = salesData.find(p => p.nama_produk === selectedProduct);
        if(product) {
            avgDemand = product.terjual;
            productName = selectedProduct;
        } else {
            const totalTerjual = salesData.reduce((sum, p) => sum + p.terjual, 0);
            avgDemand = totalTerjual / salesData.length;
            productName = 'Semua Produk (Rata-rata)';
        }
    }
    
    console.log('Rata-rata demand:', avgDemand);
    
    const months = ['Jan', 'Feb', 'Mar', 'Apr', 'Mei', 'Jun', 'Jul', 'Ags', 'Sep', 'Okt', 'Nov', 'Des'];
    
    // Faktor musiman
    const seasonalFactor = [1.1, 1.0, 0.9, 0.95, 1.0, 1.05, 1.1, 1.15, 1.1, 1.2, 1.25, 1.3];
    
    // Hitung prediksi
    const predictions = [];
    for(let i = 0; i < 12; i++) {
        let predicted = Math.round(avgDemand * seasonalFactor[i]);
        predictions.push(predicted);
    }
    
    // Update Grafik - HANYA SATU DATASET (HAPUS SEMUA LEGENDA LAIN)
    const canvas = document.getElementById('yearlyForecastChart');
    if(canvas) {
        if(yearlyForecastChart) yearlyForecastChart.destroy();
        
        const ctx = canvas.getContext('2d');
        yearlyForecastChart = new Chart(ctx, {
            type: 'bar',
            data: {
                labels: months,
                datasets: [{
                    label: `${productName}`,
                    data: predictions,
                    backgroundColor: '#3b82f6',  // SATU WARNA UNTUK SEMUA
                    borderRadius: 8
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                plugins: {
                    legend: { 
                        position: 'top',
                        labels: { boxWidth: 12 }  // PERKECIL KOTAK LEGENDA
                    }
                },
                scales: {
                    y: { 
                        title: { display: true, text: 'Demand (unit)' },
                        beginAtZero: true
                    },
                    x: {
                        title: { display: true, text: 'Bulan' }
                    }
                }
            }
        });
        console.log('Grafik berhasil dibuat');
    }
    
    // Update Tabel
    const tableBody = document.getElementById('yearlyForecastBody');
    if(tableBody) {
        tableBody.innerHTML = '';
        for(let i = 0; i < 12; i++) {
            let trend = '';
            let recom = '';
            if(predictions[i] > avgDemand * 1.2) {
                trend = '📈 Naik Drastis';
                recom = 'Siapkan stok 2x lipat';
            } else if(predictions[i] > avgDemand) {
                trend = '📈 Naik';
                recom = 'Tambah stok 30%';
            } else if(predictions[i] < avgDemand * 0.8) {
                trend = '📉 Turun';
                recom = 'Kurangi stok';
            } else {
                trend = '➡️ Stabil';
                recom = 'Pertahankan';
            }
            
            tableBody.innerHTML += `
                <tr>
                    <td><strong>${months[i]}</strong></td>
                    <td><strong style="color: #3b82f6;">${predictions[i].toLocaleString()}</strong> unit</td>
                    <td>${trend}</td>
                    <td>${recom}</td>
                </tr>
            `;
        }
    }
    
    // Update Ringkasan
    const totalYear = predictions.reduce((a,b) => a + b, 0);
    const summaryDiv = document.getElementById('yearlySummary');
    if(summaryDiv) {
        summaryDiv.innerHTML = `
            <div style="display: flex; justify-content: space-around; flex-wrap: wrap; gap: 16px;">
                <div style="text-align: center; flex: 1;">
                    <div style="font-size: 12px; color: #64748b;">TOTAL TAHUNAN</div>
                    <div style="font-size: 24px; font-weight: 700; color: #8b5cf6;">${totalYear.toLocaleString()}</div>
                </div>
                <div style="text-align: center; flex: 1;">
                    <div style="font-size: 12px; color: #64748b;">RATA-RATA/BULAN</div>
                    <div style="font-size: 24px; font-weight: 700; color: #3b82f6;">${Math.round(totalYear/12).toLocaleString()}</div>
                </div>
                <div style="text-align: center; flex: 1;">
                    <div style="font-size: 12px; color: #64748b;">BULAN TERTINGGI</div>
                    <div style="font-size: 20px; font-weight: 700; color: #ef4444;">${months[predictions.indexOf(Math.max(...predictions))]}</div>
                </div>
                <div style="text-align: center; flex: 1;">
                    <div style="font-size: 12px; color: #64748b;">BULAN TERENDAH</div>
                    <div style="font-size: 20px; font-weight: 700; color: #10b981;">${months[predictions.indexOf(Math.min(...predictions))]}</div>
                </div>
            </div>
        `;
    }
}
// ========== DEMAND PER PRODUK ==========

function updateDemandPerProduct() {
    if(!salesData.length) {
        const tbody = document.getElementById('demandPerProductBody');
        if(tbody) tbody.innerHTML = '<td><td colspan="6">Upload data terlebih dahulu</td></tr>';
        return;
    }
    
    // Urutkan berdasarkan demand (terjual) tertinggi
    const sortedProducts = [...salesData].sort((a,b) => b.terjual - a.terjual);
    
    const tbody = document.getElementById('demandPerProductBody');
    if(!tbody) return;
    
    tbody.innerHTML = '';
    
    sortedProducts.forEach(product => {
        // Tentukan status demand
        let statusText = '';
        let statusColor = '';
        let statusIcon = '';
        
        if(product.terjual > 50) {
            statusText = '🔥 Sangat Tinggi';
            statusColor = '#ef4444';
            statusIcon = '🚀';
        } else if(product.terjual > 30) {
            statusText = '📈 Tinggi';
            statusColor = '#f97316';
            statusIcon = '⬆️';
        } else if(product.terjual > 15) {
            statusText = '➡️ Sedang';
            statusColor = '#3b82f6';
            statusIcon = '➡️';
        } else if(product.terjual > 5) {
            statusText = '🐢 Rendah';
            statusColor = '#8b5cf6';
            statusIcon = '🐢';
        } else {
            statusText = '💀 Sangat Rendah';
            statusColor = '#64748b';
            statusIcon = '💀';
        }
        
        // Tambahkan rekomendasi
        let rekomendasi = '';
        if(product.terjual > 50 && product.stok < 20) {
            rekomendasi = '⚠️ Restok segera!';
        } else if(product.terjual > 50) {
            rekomendasi = '✅ Pertahankan stok';
        } else if(product.terjual < 10 && product.stok > 50) {
            rekomendasi = '💡 Beri diskon!';
        } else {
            rekomendasi = '📊 Pantau terus';
        }
        
        tbody.innerHTML += `
            <tr style="border-bottom: 1px solid #f1f5f9;">
                <td style="padding: 12px 8px;"><strong>${product.nama_produk}</strong></td>
                <td style="padding: 12px 8px;">Rp ${product.harga_jual.toLocaleString()}</td>
                <td style="padding: 12px 8px;">${product.stok.toLocaleString()}</td>
                <td style="padding: 12px 8px;"><strong style="color: #3b82f6;">${product.terjual}</strong> unit</td>
                <td style="padding: 12px 8px;">
                    <span style="background: ${statusColor}20; color: ${statusColor}; padding: 4px 10px; border-radius: 20px; font-size: 12px;">
                        ${statusIcon} ${statusText}
                    </span>
                </td>
                <td style="padding: 12px 8px; font-size: 12px;">${rekomendasi}</td>
            </tr>
        `;
    });
    
    // Tambahkan ringkasan demand
    const highDemand = salesData.filter(p => p.terjual > 50).length;
    const midDemand = salesData.filter(p => p.terjual > 15 && p.terjual <= 50).length;
    const lowDemand = salesData.filter(p => p.terjual <= 15).length;
    
    const summaryHtml = `
        <div style="margin-top: 16px; padding: 12px; background: #f8fafc; border-radius: 12px; display: flex; justify-content: space-around; flex-wrap: wrap; gap: 12px;">
            <div style="text-align: center;">
                <span style="background: #ef4444; width: 12px; height: 12px; display: inline-block; border-radius: 50%;"></span>
                <span style="margin-left: 6px;">Sangat Tinggi: <strong>${highDemand}</strong></span>
            </div>
            <div style="text-align: center;">
                <span style="background: #3b82f6; width: 12px; height: 12px; display: inline-block; border-radius: 50%;"></span>
                <span style="margin-left: 6px;">Sedang: <strong>${midDemand}</strong></span>
            </div>
            <div style="text-align: center;">
                <span style="background: #64748b; width: 12px; height: 12px; display: inline-block; border-radius: 50%;"></span>
                <span style="margin-left: 6px;">Rendah: <strong>${lowDemand}</strong></span>
            </div>
        </div>
    `;
    
    // Tambahkan ringkasan di bawah tabel (opsional)
    const existingSummary = document.getElementById('demandSummary');
    if(!existingSummary) {
        const tableContainer = document.querySelector('#demandPerProductTable')?.closest('div');
        if(tableContainer && !document.getElementById('demandSummary')) {
            const summaryDiv = document.createElement('div');
            summaryDiv.id = 'demandSummary';
            summaryDiv.innerHTML = summaryHtml;
            tableContainer.appendChild(summaryDiv);
        }
    } else {
        existingSummary.innerHTML = summaryHtml;
    }
}

// ========== EVENT LISTENER UNTUK DROPDOWN PRODUK ==========
document.addEventListener('DOMContentLoaded', function() {
    const productSelect = document.getElementById('yearlyProductSelect');
    if(productSelect) {
        productSelect.addEventListener('change', function() {
            updateYearlyForecast();  // panggil ulang dengan produk yang dipilih
        });
    }
});

// LESGOOOO SELESAI
window.addEventListener('load', () => {
    console.log('AI Smart Retail Dashboard Ready');
});

// ========== FUNGSI DOWNLOAD CONTOH FILE ==========
function downloadContohData() {
    // Data contoh lengkap
    const dataContoh = `Produk,Harga_Jual,Stok,Terjual,Tanggal
Pempek Lenjer Pack,62572,78,38,2026-04-01
Keripik Singkong Pedas,14463,118,20,2026-04-01
Cokelat Kurma,60503,144,25,2026-04-01
Kopi Bubuk Robusta,43440,102,9,2026-04-01
Siomay Ikan Frozen,42097,81,34,2026-04-01
Cokelat Kurma,56240,59,30,2026-04-01
Kacang Mete Goreng,47031,121,23,2026-04-01
Pempek Lenjer Pack,45555,117,42,2026-04-02
Kopi Bubuk Robusta,53943,138,21,2026-04-02
Manisan Mangga,37892,34,9,2026-04-02
Dendeng Sapi Balado,81594,90,39,2026-04-02
Keripik Singkong Pedas,15020,99,24,2026-04-02
Kacang Mete Goreng,64335,74,4,2026-04-02
Rendang Daging Kemasan,84671,55,25,2026-04-02
Kacang Mete Goreng,45645,88,24,2026-04-03
Bakso Sapi Frozen,43424,76,32,2026-04-03
Sambal Roa Botol,48388,49,7,2026-04-03
Kue Semprong,23380,100,17,2026-04-03
Keripik Singkong Pedas,15847,74,32,2026-04-03
Siomay Ikan Frozen,34732,30,3,2026-04-03
Emping Melinjo,32154,124,35,2026-04-04
Bumbu Rendang Instan,23698,78,15,2026-04-04
Bumbu Rendang Instan,23748,104,13,2026-04-04
Kue Semprong,29593,130,42,2026-04-04
Siomay Ikan Frozen,42867,101,14,2026-04-04
Telur Asin Brebes,6611,134,41,2026-04-04
Selai Nanas Homemade,29613,23,14,2026-04-05
Lapis Legit Slice,22544,49,31,2026-04-05
Cokelat Kurma,48661,47,26,2026-04-05
Sambal Roa Botol,46679,30,17,2026-04-05
Dendeng Sapi Balado,72029,96,29,2026-04-05
Sale Pisang Jari,15290,130,10,2026-04-06
Keripik Singkong Pedas,18785,88,11,2026-04-07
Kerupuk Ikan Palembang,27736,94,9,2026-04-07
Keripik Singkong Pedas,14750,113,26,2026-04-07
Abon Sapi Solo,64368,106,38,2026-04-07
Bakso Sapi Frozen,58893,102,6,2026-04-08
Keripik Tempe Sagu,22431,60,22,2026-04-08
Kue Kering Nastar,88097,23,7,2026-04-08
Cokelat Kurma,46027,84,10,2026-04-09
Pempek Lenjer Pack,66373,121,3,2026-04-09
Dendeng Sapi Balado,83628,30,22,2026-04-09
Kue Semprong,26995,81,14,2026-04-09
Dendeng Sapi Balado,65556,134,36,2026-04-09
Nugget Ayam Homemade,43644,25,21,2026-04-09
Kue Kering Nastar,88594,66,14,2026-04-09
Kue Kering Nastar,91635,90,20,2026-04-09
Lapis Legit Slice,21565,115,15,2026-04-09
Kacang Mete Goreng,49898,88,43,2026-04-10
Selai Nanas Homemade,31311,100,34,2026-04-10
Pempek Lenjer Pack,55852,131,9,2026-04-10
Kerupuk Ikan Palembang,26005,77,6,2026-04-10
Keripik Singkong Pedas,13314,71,24,2026-04-11
Abon Sapi Solo,55191,85,18,2026-04-11
Emping Melinjo,33772,78,23,2026-04-11
Emping Melinjo,38242,49,5,2026-04-11
Telur Asin Brebes,6863,24,20,2026-04-11
Selai Nanas Homemade,31190,128,44,2026-04-11
Dendeng Sapi Balado,65681,103,23,2026-04-11
Keripik Tempe Sagu,24223,92,37,2026-04-12
Sambal Roa Botol,43673,98,36,2026-04-12
Nugget Ayam Homemade,35017,21,6,2026-04-12
Kopi Bubuk Robusta,40436,58,14,2026-04-13
Sale Pisang Jari,17724,105,37,2026-04-13
Kerupuk Ikan Palembang,30704,94,21,2026-04-13
Selai Nanas Homemade,30591,53,42,2026-04-13
Emping Melinjo,32563,101,41,2026-04-13
Keripik Tempe Sagu,18971,109,15,2026-04-13
Bumbu Rendang Instan,19266,117,28,2026-04-14
Cokelat Kurma,49726,99,14,2026-04-14
Kacang Mete Goreng,57005,128,8,2026-04-14
Bumbu Rendang Instan,22607,41,2,2026-04-14
Siomay Ikan Frozen,34433,91,27,2026-04-14
Lapis Legit Slice,16010,98,12,2026-04-14
Siomay Ikan Frozen,42672,90,14,2026-04-14
Sale Pisang Jari,17928,100,2,2026-04-15
Dendeng Sapi Balado,81334,119,19,2026-04-15
Cokelat Kurma,45761,145,12,2026-04-15
Selai Nanas Homemade,30166,139,18,2026-04-15
Nugget Ayam Homemade,44909,104,24,2026-04-15
Kue Kering Nastar,78627,136,31,2026-04-15
Dendeng Sapi Balado,71443,70,6,2026-04-15
Keripik Singkong Pedas,16182,123,11,2026-04-15
Sale Pisang Jari,18906,100,7,2026-04-16
Telur Asin Brebes,6757,44,20,2026-04-16
Madu Hutan Alami,113896,73,30,2026-04-16
Keripik Singkong Pedas,17477,76,33,2026-04-16
Bakso Sapi Frozen,40351,139,8,2026-04-16
Nugget Ayam Homemade,49061,81,24,2026-04-16
Pempek Lenjer Pack,66739,23,13,2026-04-16
Pempek Lenjer Pack,46826,112,12,2026-04-17
Lapis Legit Slice,24727,66,36,2026-04-17
Pempek Lenjer Pack,63744,30,18,2026-04-17
Kue Kering Nastar,89069,65,4,2026-04-17
Kue Kering Nastar,89311,132,10,2026-04-17
Keripik Singkong Pedas,18394,89,24,2026-04-17
Emping Melinjo,43122,83,23,2026-04-17
Rendang Daging Kemasan,93537,133,3,2026-04-17
Kue Kering Nastar,93709,69,32,2026-04-18
Kacang Mete Goreng,56667,41,32,2026-04-18
Manisan Mangga,36065,64,44,2026-04-18
Emping Melinjo,41386,41,29,2026-04-18
Cokelat Kurma,48150,82,32,2026-04-18
Dendeng Sapi Balado,72932,108,24,2026-04-19
Bumbu Rendang Instan,23463,64,2,2026-04-19
Manisan Mangga,30101,84,40,2026-04-19
Keripik Tempe Sagu,19267,58,32,2026-04-19
Kopi Bubuk Robusta,48799,117,28,2026-04-20
Sambal Roa Botol,42201,64,38,2026-04-20
Dendeng Sapi Balado,69860,91,25,2026-04-20
Keripik Singkong Pedas,18819,149,31,2026-04-20
Kacang Mete Goreng,63994,102,24,2026-04-20
Kue Semprong,28757,44,32,2026-04-21
Lapis Legit Slice,15692,30,16,2026-04-21
Rendang Daging Kemasan,92624,30,15,2026-04-21
Emping Melinjo,42821,29,6,2026-04-21
Dendeng Sapi Balado,81434,42,24,2026-04-21
Kue Semprong,29812,138,4,2026-04-22
Selai Nanas Homemade,29479,107,36,2026-04-22
Lapis Legit Slice,24512,103,7,2026-04-22
Kacang Mete Goreng,52393,30,20,2026-04-23
Madu Hutan Alami,100211,97,40,2026-04-23
Bakso Sapi Frozen,57732,107,33,2026-04-23
Emping Melinjo,38859,24,16,2026-04-23
Sale Pisang Jari,20326,41,36,2026-04-23
Pempek Lenjer Pack,51381,119,30,2026-04-23
Kue Semprong,21157,77,36,2026-04-24
Pempek Lenjer Pack,56305,138,36,2026-04-24
Madu Hutan Alami,116897,115,7,2026-04-24
Abon Sapi Solo,70615,56,28,2026-04-24
Kopi Bubuk Robusta,35831,68,32,2026-04-24
Pempek Lenjer Pack,60834,84,9,2026-04-24
Rendang Daging Kemasan,102698,110,4,2026-04-24
Pempek Lenjer Pack,47085,138,10,2026-04-25
Telur Asin Brebes,6917,55,2,2026-04-25
Telur Asin Brebes,5618,147,44,2026-04-25
Sale Pisang Jari,15734,90,14,2026-04-25
Abon Sapi Solo,57713,29,10,2026-04-25
Sambal Roa Botol,36400,137,4,2026-04-26
Sale Pisang Jari,21601,87,26,2026-04-26
Sambal Roa Botol,39920,122,4,2026-04-26
Abon Sapi Solo,66098,138,13,2026-04-26
Abon Sapi Solo,65523,95,38,2026-04-26
Sambal Roa Botol,37689,89,38,2026-04-27
Kue Semprong,23372,87,31,2026-04-27
Bakso Sapi Frozen,59707,60,40,2026-04-27
Madu Hutan Alami,124570,28,13,2026-04-27
Abon Sapi Solo,68622,148,16,2026-04-27
Keripik Singkong Pedas,19962,43,2,2026-04-27
Keripik Tempe Sagu,24365,120,12,2026-04-28
Telur Asin Brebes,6844,35,21,2026-04-28
Sambal Roa Botol,41513,53,19,2026-04-28
Kue Kering Nastar,84288,30,4,2026-04-28
Abon Sapi Solo,58883,96,10,2026-04-28
Keripik Tempe Sagu,16159,38,9,2026-04-28
Bakso Sapi Frozen,52920,92,4,2026-04-28
Kacang Mete Goreng,54173,77,18,2026-04-28
Siomay Ikan Frozen,40457,74,42,2026-04-28
Kopi Bubuk Robusta,53843,72,22,2026-04-28
Kerupuk Ikan Palembang,21134,85,7,2026-04-29
Kopi Bubuk Robusta,44394,135,40,2026-04-29
Kerupuk Ikan Palembang,26728,143,41,2026-04-29
Siomay Ikan Frozen,44527,88,42,2026-04-29
Nugget Ayam Homemade,35777,113,30,2026-04-29
Dendeng Sapi Balado,70749,86,1,2026-04-29
Rendang Daging Kemasan,105591,54,38,2026-04-29
Keripik Tempe Sagu,15752,25,9,2026-04-29
Sale Pisang Jari,19425,74,10,2026-04-29
Emping Melinjo,30478,134,7,2026-04-30
Emping Melinjo,35645,142,5,2026-04-30
Sale Pisang Jari,19488,81,23,2026-04-30
Pempek Lenjer Pack,46774,43,41,2026-04-30
Sale Pisang Jari,21744,149,43,2026-04-30
Keripik Tempe Sagu,16785,138,42,2026-04-30
Kue Semprong,26924,57,41,2026-04-30`;

    // Proses download
    const blob = new Blob([dataContoh], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'contoh_data_pasar.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    
    // Kasih notifikasi
    alert('Download dimulai! File "contoh_data_pasar.csv" akan tersimpan di folder download Anda.');
}

// ========== DOWNLOAD EXCEL PASTI JALAN ==========
document.getElementById('excelButton').addEventListener('click', function() {
    // Data dalam bentuk array 2D
    const data = [
        ["nama_produk", "harga_jual", "stok", "terjual", "tanggal"],
        ["Pempek Lenjer Pack", 62572, 78, 38, "2026-04-01"],
        ["Keripik Singkong Pedas", 14463, 118, 20, "2026-04-01"],
        ["Cokelat Kurma", 60503, 144, 25, "2026-04-01"],
        ["Kopi Bubuk Robusta", 43440, 102, 9, "2026-04-01"],
        ["Siomay Ikan Frozen", 42097, 81, 34, "2026-04-01"],
        ["Telur Asin Brebes", 6611, 134, 41, "2026-04-04"],
        ["Kue Semprong", 29593, 130, 42, "2026-04-04"],
        ["Dendeng Sapi Balado", 81594, 90, 39, "2026-04-02"],
        ["Rendang Daging Kemasan", 84671, 55, 25, "2026-04-02"],
        ["Abon Sapi Solo", 64368, 106, 38, "2026-04-07"]
    ];
    
    // Buat worksheet
    var ws = XLSX.utils.aoa_to_sheet(data);
    
    // Atur lebar kolom
    ws['!cols'] = [
        {wch: 25},
        {wch: 12},
        {wch: 8},
        {wch: 10},
        {wch: 12}
    ];
    
    // Buat workbook dan download
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Data Penjualan");
    XLSX.writeFile(wb, "contoh_data.xlsx");
});

// Tombol menu untuk HP
const menuToggle = document.querySelector('.menu-toggle');
const sidebar = document.querySelector('.sidebar');

if (menuToggle) {
    menuToggle.addEventListener('click', function() {
        sidebar.classList.toggle('open');
    });
}

// Tutup sidebar saat klik di luar (opsional)
document.addEventListener('click', function(e) {
    if (sidebar && sidebar.classList.contains('open')) {
        if (!sidebar.contains(e.target) && !menuToggle.contains(e.target)) {
            sidebar.classList.remove('open');
        }
    }
});