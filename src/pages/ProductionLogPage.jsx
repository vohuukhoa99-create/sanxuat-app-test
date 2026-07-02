import { useMemo, useState } from 'react'
import { CustomerFilterCombobox } from '../components/CustomerFilterCombobox.jsx'
import { customerCatalog } from '../data/customerCatalog.js'

const emptyText = '-'
const customerOptions = Array.from(new Set((customerCatalog || []).map((customer) => customer.customerName).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'vi', { numeric: true }))
const getOrderLotCode = (order = {}) => order.lot || order.lotCode || order.orderCode || order.id || emptyText

function isChemicalGroup(group = '') {
  const value = String(group).toLowerCase()
  return value.includes('hóa chế')
    || value.includes('hóa chất')
    || value.includes('hoa che')
    || value.includes('hoa chat')
    || value.includes('chemical')
    || value.includes('hóa')
}

function isSolidGroup(group = '') {
  const value = String(group).toLowerCase()
  return value.includes('nguyên liệu rắn')
    || value.includes('nl rắn')
    || value.includes('nguyen lieu ran')
    || value.includes('nl ran')
    || value.includes('solid')
    || value.includes('rắn')
}

function formatKg(value) {
  const number = Number(value)
  return Number.isFinite(number) ? number : ''
}

function displayKg(value) {
  const number = formatKg(value)
  return number === '' ? emptyText : `${number.toLocaleString('vi-VN')} kg`
}

function diffMinutes(start, end) {
  const startDate = new Date(String(start).replace(' ', 'T'))
  const endDate = new Date(String(end).replace(' ', 'T'))
  if (Number.isNaN(startDate.getTime()) || Number.isNaN(endDate.getTime())) return emptyText
  const minutes = Math.max(0, Math.round((endDate - startDate) / 60000))
  const hours = Math.floor(minutes / 60)
  const rest = minutes % 60
  return hours > 0 ? `${hours}h ${rest}m` : `${rest}m`
}

function normalizeIngredient(item, orderId, index, groupLabel) {
  const requiredKg = Number(item.requiredKg ?? item.materialPerLot ?? 0)
  const actualWeight = item.actualWeight ?? item.actual ?? ''
  const actualNumber = Number(actualWeight)
  const variance = actualWeight === '' || !Number.isFinite(actualNumber)
    ? ''
    : Number((actualNumber - requiredKg).toFixed(3))

  return {
    orderId,
    stt: index + 1,
    group: groupLabel,
    materialCode: item.materialCode || '',
    materialName: item.materialName || item.name || '',
    requiredKg,
    actualWeight,
    variance,
    toleranceKg: Number(item.toleranceKg ?? 0),
    qrScanned: item.qrScanned || '',
    qrStatus: item.qrStatus || 'Chờ quét',
    weighStatus: item.weighStatus || 'Chờ cân',
    operator: item.operator || item.weighedBy || '',
    confirmedAt: item.confirmedAt || '',
    note: item.note || '',
  }
}

function normalizeOrder(order) {
  const ingredients = order.ingredients || []
  const chemicalRows = ingredients
    .filter((item) => isChemicalGroup(item.materialGroup))
    .map((item, index) => normalizeIngredient(item, getOrderLotCode(order), index, 'Hóa chế'))
  const solidRows = ingredients
    .filter((item) => isSolidGroup(item.materialGroup))
    .map((item, index) => normalizeIngredient(item, getOrderLotCode(order), index, 'Nguyên liệu rắn'))
  const actualWeight = Number(order.mixing?.finalWeightKg ?? order.actualWeightKg ?? '')
  const requestedWeight = Number(order.quantityKg ?? 0)
  const lossRate = Number.isFinite(actualWeight) && actualWeight > 0 && requestedWeight > 0
    ? Number((((requestedWeight - actualWeight) / requestedWeight) * 100).toFixed(2))
    : ''
  const status = order.status || order.mixing?.status || order.stage || 'Pending'

  return {
    ...order,
    productionDate: order.productionDate || String(order.createdAt || '').slice(0, 10),
    customer: order.customer || '',
    formula: order.formula || '',
    lot: order.lot || '',
    quantityKg: requestedWeight,
    actualWeightKg: Number.isFinite(actualWeight) ? actualWeight : '',
    lossRate,
    batchCount: order.batchCount || '',
    machineNo: order.mixing?.mixerNo || order.machineNo || '',
    status,
    owner: order.owner || order.mixing?.operator || '',
    qcOwner: order.qc?.operator || order.qcOwner || '',
    note: order.note || order.mixing?.note || '',
    createdAt: order.createdAt || '',
    receivedAt: order.receivedAt || order.createdAt || '',
    chemicalStartedAt: order.chemicalStartedAt || '',
    chemicalCompletedAt: order.chemicalCompletedAt || '',
    solidStartedAt: order.solidStartedAt || '',
    solidCompletedAt: order.solidCompletedAt || '',
    readyAt: order.mixing?.readyAt || '',
    mixingStartedAt: order.mixing?.startedAt || '',
    mixingCompletedAt: order.mixing?.completedAt || '',
    qcTransferredAt: order.qcTransferredAt || order.mixing?.completedAt || '',
    qcCheckedAt: order.qc?.checkedAt || order.qcCheckedAt || '',
    completedAt: order.completedAt || (order.stage === 'completed' ? order.updatedAt : ''),
    chemicalRows,
    solidRows,
  }
}

function getSummaryRows(orders) {
  return orders.map((order, index) => ({
    STT: index + 1,
    Ngày: order.productionDate || emptyText,
    'Mã lô': getOrderLotCode(order),
    'Khách hàng': order.customer || emptyText,
    'Sản phẩm': order.product || emptyText,
    'Khối lượng yêu cầu': displayKg(order.quantityKg),
    'Khối lượng thực tế': displayKg(order.actualWeightKg),
    'Tỷ lệ hao hụt': order.lossRate === '' ? emptyText : `${order.lossRate}%`,
    'Trạng thái': order.status || emptyText,
    'Tổng thời gian hoàn thành': diffMinutes(order.createdAt, order.completedAt || order.qcCheckedAt || order.mixingCompletedAt),
  }))
}

function getWeighRows(orders, type) {
  return orders.flatMap((order) =>
    (type === 'chemical' ? order.chemicalRows : order.solidRows).map((item) => ({
      'Mã lô': getOrderLotCode(order),
      'Mã vật tư': item.materialCode,
      [type === 'chemical' ? 'Tên hóa chất' : 'Tên nguyên liệu rắn']: item.materialName,
      'Khối lượng yêu cầu': displayKg(item.requiredKg),
      'Khối lượng thực tế': displayKg(item.actualWeight),
      'Sai lệch': displayKg(item.variance),
      'Dung sai': displayKg(item.toleranceKg),
      'QR quét được': item.qrScanned || emptyText,
      'Trạng thái QR': item.qrStatus,
      'Trạng thái cân': item.weighStatus,
      'Người cân': item.operator || emptyText,
      'Thời gian cân': item.confirmedAt || emptyText,
      'Ghi chú': item.note || emptyText,
    })),
  )
}

function getMixQcRows(orders) {
  return orders.map((order) => ({
    'Mã lô': getOrderLotCode(order),
    'Người phối trộn': order.mixing?.operator || emptyText,
    'Máy trộn số': order.mixing?.mixerNo || emptyText,
    'Thời gian bắt đầu': order.mixingStartedAt || emptyText,
    'Thời gian hoàn thành': order.mixingCompletedAt || emptyText,
    'Khối lượng sau phối trộn': displayKg(order.actualWeightKg),
    'Ghi chú phối trộn': order.mixing?.note || emptyText,
    'Trạng thái chuyển QC': order.mixing?.confirmQc ? 'Đã chuyển QC' : order.stage === 'qc' ? 'QC' : emptyText,
    'QC phụ trách': order.qcOwner || emptyText,
    'Thời gian QC kiểm tra': order.qcCheckedAt || emptyText,
    'Thời gian hoàn thành lệnh': order.completedAt || emptyText,
  }))
}

function tableToHtml(title, rows) {
  const headers = Object.keys(rows[0] || { 'Không có dữ liệu': '' })
  const body = rows.length ? rows : [{ 'Không có dữ liệu': '' }]
  return `
    <h2>${title}</h2>
    <table>
      <thead><tr>${headers.map((header) => `<th>${header}</th>`).join('')}</tr></thead>
      <tbody>${body.map((row) => `<tr>${headers.map((header) => `<td>${row[header] ?? ''}</td>`).join('')}</tr>`).join('')}</tbody>
    </table>
  `
}

function downloadFile(filename, mimeType, content) {
  const blob = new Blob([content], { type: mimeType })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}

function exportExcel(orders) {
  const sheets = [
    ['Tổng hợp mã lô', getSummaryRows(orders)],
    ['Chi tiết cân hóa', getWeighRows(orders, 'chemical')],
    ['Chi tiết cân rắn', getWeighRows(orders, 'solid')],
    ['Phối trộn và QC', getMixQcRows(orders)],
  ]
  const html = `
    <html>
      <head><meta charset="UTF-8"></head>
      <body>${sheets.map(([title, rows]) => tableToHtml(title, rows)).join('<br/>')}</body>
    </html>
  `
  downloadFile('nhat-ky-san-xuat.xls', 'application/vnd.ms-excel;charset=utf-8', html)
}

function exportPdf(orders) {
  const html = `
    <html>
      <head>
        <meta charset="UTF-8">
        <title>NHẬT KÝ SẢN XUẤT</title>
        <style>
          body { font-family: Arial, sans-serif; color: #123b2b; }
          h1 { text-align: center; }
          table { width: 100%; border-collapse: collapse; margin-bottom: 22px; font-size: 11px; }
          th, td { border: 1px solid #9bb5aa; padding: 6px; text-align: left; }
          th { background: #e8f4ef; }
          @media print { button { display: none; } }
        </style>
      </head>
      <body>
        <button onclick="window.print()">Lưu PDF</button>
        <h1>NHẬT KÝ SẢN XUẤT</h1>
        ${tableToHtml('Thông tin tổng hợp mã lô', getSummaryRows(orders))}
        ${tableToHtml('Bảng chi tiết cân hóa', getWeighRows(orders, 'chemical'))}
        ${tableToHtml('Bảng chi tiết cân rắn', getWeighRows(orders, 'solid'))}
        ${tableToHtml('Bảng phối trộn / QC', getMixQcRows(orders))}
        <script>window.onload = () => window.print()</script>
      </body>
    </html>
  `
  const printWindow = window.open('', '_blank')
  printWindow.document.write(html)
  printWindow.document.close()
}

function OrderDetail({ order, logs, onClose }) {
  const [tab, setTab] = useState('info')
  const relatedLogs = logs.filter((log) => String(log.entry || '').includes(order.id))

  return (
    <section className="panel production-log-detail">
      <div className="section-heading-row">
        <div>
          <h3>Chi tiết nhật ký {getOrderLotCode(order)}</h3>
          <p className="panel-text">{order.product}</p>
        </div>
        <button className="secondary-button" onClick={onClose}>Đóng chi tiết</button>
      </div>

      <div className="log-tabs">
        <button className={tab === 'info' ? 'active' : ''} onClick={() => setTab('info')}>Thông tin lệnh</button>
        <button className={tab === 'chemical' ? 'active' : ''} onClick={() => setTab('chemical')}>Nhật ký cân hóa</button>
        <button className={tab === 'solid' ? 'active' : ''} onClick={() => setTab('solid')}>Nhật ký cân rắn</button>
        <button className={tab === 'mixqc' ? 'active' : ''} onClick={() => setTab('mixqc')}>Phối trộn / QC</button>
      </div>

      {tab === 'info' && (
        <div className="production-log-grid">
          {[
            ['Mã lô', getOrderLotCode(order)],
            ['Ngày sản xuất', order.productionDate],
            ['Khách hàng', order.customer],
            ['Tên sản phẩm', order.product],
            ['Mã công thức', order.formula],
            ['Khối lượng yêu cầu', displayKg(order.quantityKg)],
            ['Khối lượng thực tế', displayKg(order.actualWeightKg)],
            ['Tỷ lệ hao hụt', order.lossRate === '' ? emptyText : `${order.lossRate}%`],
            ['Số mẻ', order.batchCount],
            ['Máy số', order.machineNo],
            ['Trạng thái hiện tại', order.status],
            ['Người phụ trách', order.owner],
            ['QC phụ trách', order.qcOwner],
            ['Ghi chú', order.note],
            ['Thời gian tạo lệnh', order.createdAt],
            ['Thời gian nhận lệnh SX', order.receivedAt],
            ['Bắt đầu cân hóa', order.chemicalStartedAt],
            ['Hoàn thành cân hóa', order.chemicalCompletedAt],
            ['Bắt đầu cân rắn', order.solidStartedAt],
            ['Hoàn thành cân rắn', order.solidCompletedAt],
            ['Ready phối trộn', order.readyAt],
            ['Bắt đầu phối trộn', order.mixingStartedAt],
            ['Hoàn thành phối trộn', order.mixingCompletedAt],
            ['Chuyển QC', order.qcTransferredAt],
            ['QC kiểm tra', order.qcCheckedAt],
            ['Hoàn thành lệnh', order.completedAt],
          ].map(([label, value]) => (
            <div key={label}><span>{label}</span><strong>{value || emptyText}</strong></div>
          ))}
        </div>
      )}

      {tab === 'chemical' && <WeighingTable rows={order.chemicalRows} type="chemical" />}
      {tab === 'solid' && <WeighingTable rows={order.solidRows} type="solid" />}
      {tab === 'mixqc' && (
        <div className="production-log-stack">
          <div className="production-log-grid">
            {Object.entries(getMixQcRows([order])[0]).map(([label, value]) => (
              <div key={label}><span>{label}</span><strong>{value || emptyText}</strong></div>
            ))}
          </div>
          <h3>Sự kiện productionLogs</h3>
          <ul className="log-list full-log">
            {relatedLogs.map((log) => (
              <li key={log.id}><span className="log-time">{log.time}</span><span>{log.entry}</span></li>
            ))}
            {relatedLogs.length === 0 && <li>Chưa có sự kiện ghi nhận riêng cho lệnh này.</li>}
          </ul>
        </div>
      )}
    </section>
  )
}

function WeighingTable({ rows, type }) {
  return (
    <div className="table-wrapper">
      <table className="production-log-wide-table">
        <thead>
          <tr>
            <th>Mã lô</th>
            <th>Mã vật tư</th>
            <th>{type === 'chemical' ? 'Tên hóa chất' : 'Tên nguyên liệu rắn'}</th>
            <th>Khối lượng yêu cầu</th>
            <th>Khối lượng thực tế</th>
            <th>Sai lệch</th>
            <th>Dung sai</th>
            <th>QR quét được</th>
            <th>Trạng thái QR</th>
            <th>Trạng thái cân</th>
            <th>Người cân</th>
            <th>Thời gian cân</th>
            <th>Ghi chú</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={`${row.orderId}-${row.materialCode}`}>
              <td>{row.orderId}</td>
              <td>{row.materialCode}</td>
              <td>{row.materialName}</td>
              <td>{displayKg(row.requiredKg)}</td>
              <td>{displayKg(row.actualWeight)}</td>
              <td>{displayKg(row.variance)}</td>
              <td>{displayKg(row.toleranceKg)}</td>
              <td>{row.qrScanned || emptyText}</td>
              <td><span className={`flow-pill ${row.qrStatus === 'PASS' ? 'pass' : row.qrStatus === 'FAIL' ? 'fail' : 'locked'}`}>{row.qrStatus}</span></td>
              <td><span className={`flow-pill ${row.weighStatus === 'PASS' ? 'pass' : row.weighStatus === 'FAIL' ? 'fail' : 'locked'}`}>{row.weighStatus}</span></td>
              <td>{row.operator || emptyText}</td>
              <td>{row.confirmedAt || emptyText}</td>
              <td>{row.note || emptyText}</td>
            </tr>
          ))}
          {rows.length === 0 && <tr><td colSpan="13" className="empty-row">Không có dữ liệu.</td></tr>}
        </tbody>
      </table>
    </div>
  )
}

export function ProductionLogPage({ orders = [], logs = [] }) {
  const normalizedOrders = useMemo(() => orders.map(normalizeOrder), [orders])
  const [filters, setFilters] = useState({
    fromDate: '',
    toDate: '',
    orderId: '',
    customer: '',
    product: '',
    status: '',
    actor: '',
  })
  const [selectedOrderId, setSelectedOrderId] = useState('')
  const filteredOrders = useMemo(() => normalizedOrders.filter((order) => {
    const date = order.productionDate || ''
    const actorText = `${order.owner} ${order.mixing?.operator || ''} ${order.qcOwner}`.toLowerCase()
    return (!filters.fromDate || date >= filters.fromDate)
      && (!filters.toDate || date <= filters.toDate)
      && (!filters.orderId || order.id.toLowerCase().includes(filters.orderId.toLowerCase()))
      && (!filters.customer || order.customer === filters.customer)
      && (!filters.product || String(order.product).toLowerCase().includes(filters.product.toLowerCase()))
      && (!filters.status || String(order.status).toLowerCase().includes(filters.status.toLowerCase()))
      && (!filters.actor || actorText.includes(filters.actor.toLowerCase()))
  }), [normalizedOrders, filters])
  const selectedOrder = filteredOrders.find((order) => order.id === selectedOrderId)

  const updateFilter = (field, value) => setFilters((current) => ({ ...current, [field]: value }))

  return (
    <div className="page-content production-log-page">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Nhật ký sản xuất</h2>
            <p className="panel-text">Lưu trữ toàn bộ lịch sử lệnh sản xuất, cân hóa, cân rắn, phối trộn và QC.</p>
          </div>
          <div className="action-row">
            <button className="secondary-button" onClick={() => exportExcel(filteredOrders)}>Tải Excel</button>
            <button className="primary-button" onClick={() => exportPdf(filteredOrders)}>Tải PDF</button>
          </div>
        </div>

        <div className="production-log-filters">
          <label>Từ ngày<input type="date" value={filters.fromDate} onChange={(event) => updateFilter('fromDate', event.target.value)} /></label>
          <label>Đến ngày<input type="date" value={filters.toDate} onChange={(event) => updateFilter('toDate', event.target.value)} /></label>
          <label>Mã lô<input value={filters.orderId} onChange={(event) => updateFilter('orderId', event.target.value)} /></label>
          <label>Khách hàng
            <CustomerFilterCombobox
              options={customerOptions}
              value={filters.customer}
              emptyValue=""
              onChange={(customer) => updateFilter('customer', customer)}
            />
          </label>
          <label>Sản phẩm<input value={filters.product} onChange={(event) => updateFilter('product', event.target.value)} /></label>
          <label>Trạng thái<input value={filters.status} onChange={(event) => updateFilter('status', event.target.value)} /></label>
          <label>Người thực hiện<input value={filters.actor} onChange={(event) => updateFilter('actor', event.target.value)} /></label>
        </div>

        <div className="table-wrapper">
          <table className="production-log-summary-table">
            <thead>
              <tr>
                <th>STT</th>
                <th>Ngày</th>
                <th>Mã lô</th>
                <th>Khách hàng</th>
                <th>Sản phẩm</th>
                <th>Khối lượng yêu cầu</th>
                <th>Khối lượng thực tế</th>
                <th>Tỷ lệ hao hụt</th>
                <th>Trạng thái</th>
                <th>Tổng thời gian hoàn thành</th>
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {filteredOrders.map((order, index) => (
                <tr key={order.id}>
                  <td>{index + 1}</td>
                  <td>{order.productionDate || emptyText}</td>
                  <td>{getOrderLotCode(order)}</td>
                  <td>{order.customer || emptyText}</td>
                  <td>{order.product || emptyText}</td>
                  <td>{displayKg(order.quantityKg)}</td>
                  <td>{displayKg(order.actualWeightKg)}</td>
                  <td>{order.lossRate === '' ? emptyText : `${order.lossRate}%`}</td>
                  <td><span className="status-pill">{order.status || emptyText}</span></td>
                  <td>{diffMinutes(order.createdAt, order.completedAt || order.qcCheckedAt || order.mixingCompletedAt)}</td>
                  <td><button className="primary-button" onClick={() => setSelectedOrderId(order.id)}>Xem chi tiết</button></td>
                </tr>
              ))}
              {filteredOrders.length === 0 && <tr><td colSpan="12" className="empty-row">Không có nhật ký phù hợp bộ lọc.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>

      {selectedOrder && <OrderDetail order={selectedOrder} logs={logs} onClose={() => setSelectedOrderId('')} />}
    </div>
  )
}
