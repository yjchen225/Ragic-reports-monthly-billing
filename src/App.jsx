import React, { useState, useEffect, useRef } from 'react';
import { Upload, FileText, AlertCircle, Printer, Calculator, Calendar } from 'lucide-react';

/**
 * 月結請款單總表生成器
 * V1.7 修正版
 * - 移除 colgroup 內的註解以解決 DOM Nesting 警告
 * - V1.6 功能維持不變：新增「付款備註」(w-36)，調整日期寬度 (w-22)
 */

// ----------------------------------------------------------------------
// 1. 工具函式與常數定義
// ----------------------------------------------------------------------

// 載入 SheetJS
const useSheetJS = () => {
  const [loaded, setLoaded] = useState(false);
  useEffect(() => {
    if (window.XLSX) {
      setLoaded(true);
      return;
    }
    const script = document.createElement('script');
    script.src = "https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js";
    script.async = true;
    script.onload = () => setLoaded(true);
    document.body.appendChild(script);
  }, []);
  return loaded;
};

// 格式化金額 (千分位)
const formatCurrency = (amount) => {
  if (amount === undefined || amount === null || isNaN(amount)) return '0';
  return Number(amount).toLocaleString('en-US');
};

// 清洗數字 (移除 $ , 等符號)
const cleanNumber = (val) => {
  if (typeof val === 'number') return val;
  if (!val) return 0;
  const str = String(val).replace(/[^0-9.-]+/g, "");
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
};

// 解析流水號 (SO-yyyymmdd-xxxx) 用於排序
const parseSequence = (str, type) => {
  if (!str) return 0;
  const parts = str.split('-');
  const seq = parts[parts.length - 1];
  const num = parseInt(seq, 10);
  return isNaN(num) ? 0 : num;
};

// 客戶編號排序 (Cxxxxxxx)
const sortCustomerIds = (a, b) => {
  const numA = parseInt(a.replace(/[^0-9]/g, ''), 10) || 0;
  const numB = parseInt(b.replace(/[^0-9]/g, ''), 10) || 0;
  return numA - numB;
};

// Excel 日期格式化工具
const formatExcelDate = (val) => {
  if (!val) return '';
  
  // 情況 1: Excel 序列號 (數值)
  if (typeof val === 'number') {
    if (val > 20000) {
      const date = new Date(Math.round((val - 25569) * 86400 * 1000));
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}/${m}/${d}`;
    }
    return String(val);
  }

  // 情況 2: 字串
  const str = String(val).trim();
  if (str.includes('-')) {
    return str.replace(/-/g, '/');
  }
  
  const date = new Date(str);
  if (!isNaN(date.getTime()) && str.includes('/')) {
     const y = date.getFullYear();
     const m = String(date.getMonth() + 1).padStart(2, '0');
     const d = String(date.getDate()).padStart(2, '0');
     return `${y}/${m}/${d}`;
  }

  return str;
};

// ----------------------------------------------------------------------
// 2. 主元件
// ----------------------------------------------------------------------

export default function App() {
  const isXLSXLoaded = useSheetJS();
  
  // State
  const [startDate, setStartDate] = useState('');
  const [endDate, setEndDate] = useState('');
  const [salesFile, setSalesFile] = useState(null);
  const [rmaFile, setRmaFile] = useState(null);
  const [errorMsg, setErrorMsg] = useState(null);
  const [reportData, setReportData] = useState(null);
  const [isProcessing, setIsProcessing] = useState(false);

  // 用於重置 file input
  const salesInputRef = useRef(null);
  const rmaInputRef = useRef(null);

  // 更新：加入「付款備註」
  const SALES_REQUIRED_COLUMNS = [
    '客戶編號', '客戶姓名', '銷貨單編號', '付款編號', '付款金額', '總公司名稱', '出貨日期', '銷貨單簽核日期', '付款備註'
  ];
  const RMA_REQUIRED_COLUMNS = [
    '客戶編號', '客戶姓名', '總公司名稱', '退換貨單號', '退換貨金額合計', '退換貨單備註'
  ];

  // --------------------------------------------------------------------
  // 核心邏輯
  // --------------------------------------------------------------------
  const readExcel = (file) => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const data = new Uint8Array(e.target.result);
          const workbook = window.XLSX.read(data, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const jsonData = window.XLSX.utils.sheet_to_json(worksheet, { defval: "" }); 
          resolve(jsonData);
        } catch (err) {
          reject(err);
        }
      };
      reader.onerror = (err) => reject(err);
      reader.readAsArrayBuffer(file);
    });
  };

  const validateHeaders = (data, requiredColumns, fileName) => {
    if (!data || data.length === 0) throw new Error(`檔案「${fileName}」中沒有資料。`);
    const headers = Object.keys(data[0]);
    const missing = requiredColumns.filter(col => !headers.includes(col));
    if (missing.length > 0) {
      throw new Error(`資料欄位有誤 (${fileName})，缺少必要欄位：${missing.join(', ')}，請確認後再試一次。`);
    }
  };

  const handleGenerateReport = async () => {
    setErrorMsg(null);
    setReportData(null);
    setIsProcessing(true);

    try {
      if (!startDate || !endDate) throw new Error("請輸入完整的報表日期區間。");
      if (!salesFile) throw new Error("請上傳銷貨單付款明細 Excel。");
      if (!rmaFile) throw new Error("請上傳銷退單 Excel。");
      if (!window.XLSX) throw new Error("系統元件 (SheetJS) 尚未載入完成，請稍候再試。");

      const salesRaw = await readExcel(salesFile);
      const rmaRaw = await readExcel(rmaFile);

      validateHeaders(salesRaw, SALES_REQUIRED_COLUMNS, "銷貨單付款明細");
      validateHeaders(rmaRaw, RMA_REQUIRED_COLUMNS, "銷退單");

      const allRows = [...salesRaw, ...rmaRaw];
      const companies = new Set(allRows.map(r => r['總公司名稱']).filter(n => n && String(n).trim() !== ''));
      if (companies.size > 1) {
        throw new Error(`Excel 資料異常：偵測到多個不同的總公司名稱 (${Array.from(companies).join(', ')})，請檢查資料一致性。`);
      }
      const companyName = companies.size === 1 ? Array.from(companies)[0] : "未顯示總公司名稱";

      const customerIds = new Set([
        ...salesRaw.map(r => r['客戶編號']),
        ...rmaRaw.map(r => r['客戶編號'])
      ]);
      const sortedCustomerIds = Array.from(customerIds).filter(Boolean).sort(sortCustomerIds);

      const processedCustomers = sortedCustomerIds.map(custId => {
        const clientSales = salesRaw.filter(r => r['客戶編號'] === custId);
        const clientRma = rmaRaw.filter(r => r['客戶編號'] === custId);
        const clientName = clientSales[0]?.['客戶姓名'] || clientRma[0]?.['客戶姓名'] || '未知客戶';

        const sortedSales = [...clientSales].sort((a, b) => {
           return parseSequence(a['銷貨單編號']) - parseSequence(b['銷貨單編號']);
        });

        const subtotalA = sortedSales.reduce((sum, row) => sum + cleanNumber(row['付款金額']), 0);

        const sortedRma = [...clientRma].sort((a, b) => {
          return parseSequence(a['退換貨單號']) - parseSequence(b['退換貨單號']);
        });

        const subtotalB = sortedRma.reduce((sum, row) => sum + cleanNumber(row['退換貨金額合計']), 0);
        const clientTotal = subtotalA - subtotalB;

        return {
          id: custId,
          name: clientName,
          sales: sortedSales,
          rma: sortedRma,
          subtotalA,
          subtotalB,
          clientTotal
        };
      });

      const grandTotal = processedCustomers.reduce((sum, c) => sum + c.clientTotal, 0);

      setReportData({
        header: {
          startDate,
          endDate,
          companyName,
          generatedAt: new Date().toLocaleString('zh-TW', { hour12: false, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
        },
        customers: processedCustomers,
        grandTotal
      });

    } catch (err) {
      console.error(err);
      setErrorMsg(err.message);
    } finally {
      setIsProcessing(false);
    }
  };

  const handlePrint = () => {
    window.print();
  };

  const handleReset = () => {
    setStartDate('');
    setEndDate('');
    setSalesFile(null);
    setRmaFile(null);
    setReportData(null);
    setErrorMsg(null);
    if (salesInputRef.current) salesInputRef.current.value = '';
    if (rmaInputRef.current) rmaInputRef.current.value = '';
  };

  return (
    <div className="min-h-screen bg-gray-50 font-sans text-gray-900">
      <style>{`
        @media print {
          @page { size: A4; margin: 10mm; }
          body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
          .no-print { display: none !important; }
          .print-container { 
            display: block !important; 
            width: 100%; 
            position: absolute; 
            top: 0; 
            left: 0; 
            background: white;
            margin: 0;
            padding: 0;
          }
          table { page-break-inside: auto; width: 100%; border-collapse: collapse; font-size: 10pt; }
          tr { page-break-inside: avoid; page-break-after: auto; }
          thead { display: table-header-group; }
          tfoot { display: table-row-group; }
          .page-break { page-break-before: always; }
          .avoid-break { page-break-inside: avoid; }
        }
      `}</style>

      {/* ------------------------------------------------------------------
          操作面板 (列印時隱藏)
         ------------------------------------------------------------------ */}
      <div className="no-print max-w-4xl mx-auto p-6">
        <div className="bg-blue-600 p-4 flex items-center gap-3 rounded-t-xl">
          <Calculator className="text-white w-6 h-6" />
          <h1 className="text-xl font-bold text-white">月結請款單總表生成器</h1>
        </div>
        <div className="bg-white p-6 shadow-lg border border-t-0 border-gray-200 rounded-b-xl space-y-8">
            {/* 1. 日期設定 */}
            <section>
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center text-sm">1</span>
                設定報表區間
              </h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">開始日期</label>
                  <input
                    type="date"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500"
                  />
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-1">結束日期</label>
                  <input
                    type="date"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 rounded-md focus:ring-2 focus:ring-blue-500"
                  />
                </div>
              </div>
            </section>

            {/* 2. 檔案上傳 */}
            <section>
              <h2 className="text-lg font-semibold text-gray-800 mb-4 flex items-center gap-2">
                <span className="w-6 h-6 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center text-sm">2</span>
                上傳 Excel 檔案
              </h2>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className={`border-2 border-dashed rounded-lg p-6 text-center transition-colors ${salesFile ? 'border-green-500 bg-green-50' : 'border-gray-300 hover:border-blue-400'}`}>
                  <FileText className={`w-10 h-10 mx-auto mb-2 ${salesFile ? 'text-green-600' : 'text-gray-400'}`} />
                  <h3 className="font-medium text-gray-900">銷貨單付款明細</h3>
                  <p className="text-xs text-gray-500 mb-4">包含欄位：付款編號、付款備註、銷貨單編號、簽核日期...</p>
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    ref={salesInputRef}
                    onChange={(e) => setSalesFile(e.target.files[0])}
                    className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
                  />
                  {salesFile && <p className="mt-2 text-sm text-green-700 font-medium">已選取：{salesFile.name}</p>}
                </div>
                <div className={`border-2 border-dashed rounded-lg p-6 text-center transition-colors ${rmaFile ? 'border-green-500 bg-green-50' : 'border-gray-300 hover:border-blue-400'}`}>
                  <FileText className={`w-10 h-10 mx-auto mb-2 ${rmaFile ? 'text-green-600' : 'text-gray-400'}`} />
                  <h3 className="font-medium text-gray-900">銷退單</h3>
                  <p className="text-xs text-gray-500 mb-4">包含欄位：退換貨單號、退換貨金額合計...</p>
                  <input
                    type="file"
                    accept=".xlsx, .xls"
                    ref={rmaInputRef}
                    onChange={(e) => setRmaFile(e.target.files[0])}
                    className="block w-full text-sm text-gray-500 file:mr-4 file:py-2 file:px-4 file:rounded-full file:border-0 file:text-sm file:font-semibold file:bg-blue-50 file:text-blue-700 hover:file:bg-blue-100"
                  />
                  {rmaFile && <p className="mt-2 text-sm text-green-700 font-medium">已選取：{rmaFile.name}</p>}
                </div>
              </div>
            </section>

            {errorMsg && (
              <div className="bg-red-50 border-l-4 border-red-500 p-4 flex items-start gap-3">
                <AlertCircle className="text-red-500 w-5 h-5 flex-shrink-0 mt-0.5" />
                <div>
                  <h3 className="font-bold text-red-700">無法生成報表</h3>
                  <p className="text-sm text-red-600 mt-1">{errorMsg}</p>
                </div>
              </div>
            )}

            <div className="flex gap-4 pt-4 border-t border-gray-100">
              <button
                onClick={handleGenerateReport}
                disabled={isProcessing || !isXLSXLoaded}
                className="flex-1 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-6 rounded-lg shadow transition-colors flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                {isProcessing ? '資料處理中...' : '生成報表預覽'}
              </button>
              
              {reportData && (
                <button
                  onClick={handlePrint}
                  className="bg-gray-800 hover:bg-gray-900 text-white font-bold py-3 px-6 rounded-lg shadow transition-colors flex items-center justify-center gap-2"
                >
                  <Printer className="w-5 h-5" />
                  列印 / 存為 PDF
                </button>
              )}

              <button
                 onClick={handleReset}
                 className="px-6 py-3 border border-gray-300 rounded-lg hover:bg-gray-50 text-gray-700 font-medium"
              >
                清除重置
              </button>
            </div>
            {!isXLSXLoaded && <p className="text-center text-xs text-gray-400">正在初始化 Excel 解析核心...</p>}
        </div>
      </div>

      {/* ------------------------------------------------------------------
          報表預覽區 (列印時的主體)
         ------------------------------------------------------------------ */}
      {reportData && (
        <div className="print-container mt-8 max-w-[210mm] mx-auto bg-white shadow-2xl min-h-[297mm] p-10 md:p-0">
          <div className="p-8 md:p-12">
            
            {/* 報表表頭 (Report Header) */}
            <header className="mb-6 border-b-2 border-gray-800 pb-2">
              <h1 className="text-2xl font-bold text-center text-gray-900 mb-4">月結請款單 - 總表</h1>
              <div className="grid grid-cols-2 gap-y-1 text-sm text-gray-700">
                <div className="flex items-baseline">
                  <span className="font-bold w-20 flex-shrink-0">日期區間：</span>
                  <span>{reportData.header.startDate.replace(/-/g, '/')} ~ {reportData.header.endDate.replace(/-/g, '/')}</span>
                </div>
                <div className="flex justify-end items-baseline">
                  <span className="font-bold w-20 text-right pr-2 flex-shrink-0">製表時間：</span>
                  <span>{reportData.header.generatedAt}</span>
                </div>
                <div className="col-span-2 flex items-baseline">
                  <span className="font-bold w-20 flex-shrink-0">總公司：</span>
                  <span className="text-base font-semibold">{reportData.header.companyName}</span>
                </div>
              </div>
            </header>

            {/* 報表內容：客戶迴圈 */}
            <div className="space-y-6">
              {reportData.customers.length === 0 ? (
                <p className="text-center text-gray-500 py-10">查無相關客戶資料</p>
              ) : (
                reportData.customers.map((customer, cIndex) => (
                  <div key={customer.id} className="avoid-break mb-6 pb-2 border-b border-gray-300 last:border-0">
                    
                    {/* 客戶標題 */}
                    <h2 className="text-base font-bold bg-gray-100 p-1.5 mb-2 border-l-4 border-blue-600">
                      客戶：{customer.name} ({customer.id})
                    </h2>

                    {/* A. 銷貨單付款明細表 */}
                    <div className="mb-2">
                      <h3 className="text-xs font-bold text-gray-700 mb-1 pl-1">【銷貨單付款明細】</h3>
                      {customer.sales.length > 0 ? (
                        <table className="w-full text-xs text-left border border-gray-300 mb-1 table-fixed">
                          <colgroup>
                            <col className="w-36" />
                            <col className="w-36" />
                            <col className="w-22" />
                            <col className="w-22" />
                            <col className="w-36" />
                            <col className="w-auto" />
                          </colgroup>
                          <thead className="bg-gray-50 text-gray-700 font-semibold border-b border-gray-300">
                            <tr>
                              <th className="px-1.5 py-1 border-r border-gray-300">銷貨單編號</th>
                              <th className="px-1.5 py-1 border-r border-gray-300">付款編號</th>
                              <th className="px-1.5 py-1 border-r border-gray-300">簽核日期</th>
                              <th className="px-1.5 py-1 border-r border-gray-300">出貨日期</th>
                              <th className="px-1.5 py-1 border-r border-gray-300">付款備註</th>
                              <th className="px-1.5 py-1 text-right">付款金額</th>
                            </tr>
                          </thead>
                          <tbody>
                            {customer.sales.map((row, idx) => {
                              const showSO = idx === 0 || row['銷貨單編號'] !== customer.sales[idx - 1]['銷貨單編號'];
                              return (
                                <tr key={idx} className="border-b border-gray-200">
                                  <td className="px-1.5 py-1 border-r border-gray-300 align-top break-words">
                                    {showSO ? row['銷貨單編號'] : ''}
                                  </td>
                                  <td className="px-1.5 py-1 border-r border-gray-300 align-top break-words">{row['付款編號']}</td>
                                  <td className="px-1.5 py-1 border-r border-gray-300 align-top break-words">
                                    {showSO ? formatExcelDate(row['銷貨單簽核日期']) : ''}
                                  </td>
                                  <td className="px-1.5 py-1 border-r border-gray-300 align-top break-words">
                                    {showSO ? formatExcelDate(row['出貨日期']) : ''}
                                  </td>
                                  <td className="px-1.5 py-1 border-r border-gray-300 align-top break-words">
                                    {row['付款備註']}
                                  </td>
                                  <td className="px-1.5 py-1 text-right align-top break-words">{formatCurrency(row['付款金額'])}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                          <tfoot>
                             <tr className="bg-gray-50 font-bold border-t-2 border-gray-300">
                               <td colSpan={5} className="px-1.5 py-1 text-right">銷貨付款小計：</td>
                               <td className="px-1.5 py-1 text-right">{formatCurrency(customer.subtotalA)}</td>
                             </tr>
                          </tfoot>
                        </table>
                      ) : (
                        <p className="text-xs text-gray-500 italic pl-2 mb-1">此期間無銷貨付款紀錄。</p>
                      )}
                    </div>

                    {/* B. 銷退單明細表 */}
                    <div className="mb-2">
                      <h3 className="text-xs font-bold text-gray-700 mb-1 pl-1">【銷退單明細】</h3>
                      {customer.rma.length > 0 ? (
                        <table className="w-full text-xs text-left border border-gray-300 mb-1 table-fixed">
                          <colgroup>
                            <col className="w-40" />
                            <col className="w-auto" />
                            <col className="w-28" />
                          </colgroup>
                          <thead className="bg-gray-50 text-gray-700 font-semibold border-b border-gray-300">
                            <tr>
                              <th className="px-1.5 py-1 border-r border-gray-300">退換貨單號</th>
                              <th className="px-1.5 py-1 border-r border-gray-300">備註</th>
                              <th className="px-1.5 py-1 text-right">退換貨金額</th>
                            </tr>
                          </thead>
                          <tbody>
                            {customer.rma.map((row, idx) => {
                              const showRMA = idx === 0 || row['退換貨單號'] !== customer.rma[idx - 1]['退換貨單號'];
                              return (
                                <tr key={idx} className="border-b border-gray-200">
                                  <td className="px-1.5 py-1 border-r border-gray-300 align-top break-words">
                                    {showRMA ? row['退換貨單號'] : ''}
                                  </td>
                                  <td className="px-1.5 py-1 border-r border-gray-300 align-top break-words">
                                    {row['退換貨單備註']}
                                  </td>
                                  <td className="px-1.5 py-1 text-right align-top break-words">{formatCurrency(row['退換貨金額合計'])}</td>
                                </tr>
                              );
                            })}
                          </tbody>
                          <tfoot>
                             <tr className="bg-gray-50 font-bold border-t-2 border-gray-300">
                               <td colSpan={2} className="px-1.5 py-1 text-right">退換貨小計：</td>
                               <td className="px-1.5 py-1 text-right">{formatCurrency(customer.subtotalB)}</td>
                             </tr>
                          </tfoot>
                        </table>
                      ) : (
                        <p className="text-xs text-gray-500 italic pl-2 mb-1">此期間無銷退紀錄。</p>
                      )}
                    </div>

                    {/* C. 本次請款金額列 */}
                    <div className="flex justify-end items-center bg-blue-50 p-1.5 rounded border border-blue-200">
                      <span className="text-xs font-bold text-gray-800 mr-2">
                        客戶 {customer.name} 本次請款金額：
                      </span>
                      <span className={`text-base font-bold ${customer.clientTotal < 0 ? 'text-red-600' : 'text-blue-700'}`}>
                        $ {formatCurrency(customer.clientTotal)}
                      </span>
                    </div>

                  </div>
                ))
              )}
            </div>

            {/* Footer: 總表總金額 */}
            <footer className="mt-8 border-t-4 border-gray-800 pt-4 avoid-break">
              <div className="flex justify-between items-end">
                 <div className="text-xs text-gray-500">
                    <p>備註：本報表為系統自動生成，若有疑問請聯繫財務部。</p>
                 </div>
                 <div className="text-right">
                    <p className="text-gray-600 font-bold mb-1 text-sm">本期請款總金額</p>
                    <p className={`text-3xl font-extrabold ${reportData.grandTotal < 0 ? 'text-red-600' : 'text-gray-900'}`}>
                      $ {formatCurrency(reportData.grandTotal)}
                    </p>
                 </div>
              </div>
            </footer>

          </div>
        </div>
      )}
    </div>
  );
}