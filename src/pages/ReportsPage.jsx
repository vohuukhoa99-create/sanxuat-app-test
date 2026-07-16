import { useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import { CustomerFilterCombobox } from '../components/CustomerFilterCombobox.jsx'
import { customerCatalog } from '../data/customerCatalog.js'

const stages = ['Lệnh sản xuất', 'Cân hóa', 'Cân rắn', 'Phối trộn', 'QC', 'Hoàn thành']
const customerOptions = Array.from(new Set((customerCatalog || []).map((customer) => customer.customerName).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'vi', { numeric: true }))
const getOrderLotCode = (order = {}) => order.lot || order.lotCode || order.orderCode || order.id || '-'

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

function parseDateTime(value) {
  const date = new Date(String(value || '').replace(' ', 'T'))
  return Number.isNaN(date.getTime()) ? null : date
}

function minutesBetween(start, end) {
  const startDate = parseDateTime(start)
  const endDate = parseDateTime(end)
  if (!startDate || !endDate) return 0
  return Math.max(0, Math.round((endDate - startDate) / 60000))
}

function formatDuration(minutes) {
  if (!minutes) return '-'
  const hours = Math.floor(minutes / 60)
  const rest = minutes % 60
  return hours ? `${hours}h ${rest}m` : `${rest}m`
}

function formatKg(value) {
  const number = Number(value)
  return Number.isFinite(number) ? Number(number.toFixed(3)) : 0
}

function displayKg(value) {
  return `${formatKg(value).toLocaleString('vi-VN')} kg`
}

function normalizeIngredient(item, order, stage) {
  const requiredKg = formatKg(item.requiredKg ?? item.materialPerLot)
  const actualKg = item.actualWeight === '' || item.actualWeight == null
    ? 0
    : formatKg(item.actualWeight ?? item.actual)
  const variance = actualKg ? Number((actualKg - requiredKg).toFixed(3)) : 0
  return {
    orderId: getOrderLotCode(order),
    date: order.productionDate || String(order.createdAt || '').slice(0, 10),
    stage,
    materialCode: item.materialCode || '',
    materialName: item.materialName || item.name || '',
    requiredKg,
    actualKg,
    toleranceKg: formatKg(item.toleranceKg),
    variance,
    qrStatus: item.qrStatus || 'Chờ quét',
    weighStatus: item.weighStatus || 'Chờ cân',
    operator: item.operator || item.weighedBy || '',
    note: item.note || '',
    confirmedAt: item.confirmedAt || '',
  }
}

function normalizeOrder(order) {
  const ingredients = order.ingredients || []
  const chemicalRows = ingredients
    .filter((item) => isChemicalGroup(item.materialGroup))
    .map((item) => normalizeIngredient(item, order, 'Cân hóa'))
  const solidRows = ingredients
    .filter((item) => isSolidGroup(item.materialGroup))
    .map((item) => normalizeIngredient(item, order, 'Cân rắn'))
  const requiredKg = formatKg(order.quantityKg)
  const actualKg = formatKg(order.mixing?.finalWeightKg ?? order.actualWeightKg)
  const lossKg = actualKg ? Number((requiredKg - actualKg).toFixed(3)) : 0
  const lossRate = requiredKg && actualKg ? Number(((lossKg / requiredKg) * 100).toFixed(2)) : 0
  const status = order.status || order.mixing?.status || order.stage || 'Pending'
  const currentStage = getCurrentStage(order, status)

  return {
    ...order,
    productionDate: order.productionDate || String(order.createdAt || '').slice(0, 10),
    customer: order.customer || '',
    product: order.product || '',
    status,
    currentStage,
    requiredKg,
    actualKg,
    lossKg,
    lossRate,
    chemicalRows,
    solidRows,
    allWeighRows: [...chemicalRows, ...solidRows],
    actor: `${order.owner || ''} ${order.mixing?.operator || ''} ${order.qcOwner || order.qc?.operator || ''}`.trim(),
    totalMinutes: minutesBetween(order.createdAt, order.completedAt || order.qc?.checkedAt || order.qcCheckedAt || order.mixing?.completedAt || order.updatedAt),
  }
}

function getCurrentStage(order, status) {
  if (order.stage === 'completed' || status.includes('Hoàn thành')) return 'Hoàn thành'
  if (order.stage === 'qc' || status === 'QC') return 'QC'
  if (order.mixing?.status === 'Active' || order.mixing?.status === 'Ready' || order.stage === 'mixing') return 'Phối trộn'
  if (order.scaleStatus?.solid === 'Active' || order.stage === 'solid') return 'Cân rắn'
  if (order.scaleStatus?.chemical === 'Active' || order.stage === 'chemical') return 'Cân hóa'
  return 'Lệnh sản xuất'
}

function hasStage(order, stage) {
  if (stage === 'Lệnh sản xuất') return true
  if (stage === 'Cân hóa') return ['Active', 'Completed'].includes(order.scaleStatus?.chemical)
  if (stage === 'Cân rắn') return ['Active', 'Completed'].includes(order.scaleStatus?.solid)
  if (stage === 'Phối trộn') return Boolean(order.mixing?.status)
  if (stage === 'QC') return order.stage === 'qc' || order.stage === 'completed'
  if (stage === 'Hoàn thành') return order.stage === 'completed'
  return false
}

function getStageDuration(order, stage) {
  if (stage === 'Cân hóa') return minutesBetween(order.chemicalStartedAt, order.chemicalCompletedAt)
  if (stage === 'Cân rắn') return minutesBetween(order.solidStartedAt, order.solidCompletedAt)
  if (stage === 'Phối trộn') return minutesBetween(order.mixing?.startedAt, order.mixing?.completedAt)
  if (stage === 'QC') return minutesBetween(order.qcTransferredAt || order.mixing?.completedAt, order.qc?.checkedAt || order.qcCheckedAt)
  if (stage === 'Hoàn thành') return minutesBetween(order.createdAt, order.completedAt)
  return minutesBetween(order.createdAt, order.receivedAt || order.updatedAt)
}

function getStageRows(orders) {
  return stages.map((stage) => {
    const stageOrders = orders.filter((order) => hasStage(order, stage))
    const errors = stageOrders.flatMap((order) => order.allWeighRows)
      .filter((row) => (stage === 'Cân hóa' && row.stage === 'Cân hóa') || (stage === 'Cân rắn' && row.stage === 'Cân rắn'))
      .filter((row) => row.qrStatus === 'FAIL' || row.weighStatus === 'FAIL').length
    const totalMinutes = stageOrders.reduce((sum, order) => sum + getStageDuration(order, stage), 0)
    const completed = stageOrders.filter((order) => {
      if (stage === 'Cân hóa') return order.scaleStatus?.chemical === 'Completed'
      if (stage === 'Cân rắn') return order.scaleStatus?.solid === 'Completed'
      if (stage === 'Phối trộn') return order.mixing?.status === 'Completed'
      if (stage === 'QC') return order.stage === 'completed'
      if (stage === 'Hoàn thành') return order.stage === 'completed'
      return true
    }).length

    return {
      stage,
      count: stageOrders.length,
      volume: stageOrders.reduce((sum, order) => sum + order.requiredKg, 0),
      errors,
      avgMinutes: stageOrders.length ? Math.round(totalMinutes / stageOrders.length) : 0,
      completionRate: stageOrders.length ? Math.round((completed / stageOrders.length) * 100) : 0,
    }
  })
}

function getKpis(orders) {
  const totalRequired = orders.reduce((sum, order) => sum + order.requiredKg, 0)
  const totalActual = orders.reduce((sum, order) => sum + order.actualKg, 0)
  const qrErrors = orders.flatMap((order) => order.allWeighRows).filter((row) => row.qrStatus === 'FAIL').length
  const weightErrors = orders.flatMap((order) => order.allWeighRows).filter((row) => row.weighStatus === 'FAIL').length
  return {
    totalOrders: orders.length,
    totalRequired,
    totalActual,
    avgLossRate: orders.length ? Number((orders.reduce((sum, order) => sum + order.lossRate, 0) / orders.length).toFixed(2)) : 0,
    weighingOrders: orders.filter((order) => order.currentStage === 'Cân hóa' || order.currentStage === 'Cân rắn').length,
    readyMixing: orders.filter((order) => order.status.includes('Ready') || order.mixing?.status === 'Ready').length,
    qcOrders: orders.filter((order) => order.currentStage === 'QC').length,
    completedOrders: orders.filter((order) => order.currentStage === 'Hoàn thành').length,
    qrErrors,
    weightErrors,
  }
}

function exportExcel({ summaryRows, stageRows, lossRows, errorRows, progressRows }) {
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), 'Tong_hop')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(stageRows), 'Theo_cong_doan')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(lossRows), 'Hao_hut')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(errorRows), 'Loi_can')
  XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(progressRows), 'Tien_do_lenh')
  XLSX.writeFile(workbook, 'bao-cao-san-xuat.xlsx')
}

function exportPdf(data, filters) {
  const doc = new jsPDF({ orientation: 'landscape' })
  doc.setFontSize(16)
  doc.text('BAO CAO SAN XUAT', 14, 14)
  doc.setFontSize(10)
  doc.text(`Thoi gian loc: ${filters.fromDate || '-'} den ${filters.toDate || '-'}`, 14, 22)
  autoTable(doc, {
    startY: 28,
    head: [['KPI', 'Gia tri']],
    body: data.kpiRows.map((row) => [row.label, row.value]),
  })
  const tables = [
    ['Bao cao theo cong doan', data.stageRows],
    ['Bao cao hao hut', data.lossRows],
    ['Bao cao loi can', data.errorRows],
    ['Bao cao tien do lenh', data.progressRows],
  ]
  tables.forEach(([title, rows]) => {
    doc.addPage('landscape')
    doc.text(title, 14, 14)
    const headers = Object.keys(rows[0] || { 'Khong co du lieu': '' })
    autoTable(doc, {
      startY: 20,
      head: [headers],
      body: rows.map((row) => headers.map((header) => row[header])),
      styles: { fontSize: 7 },
    })
  })
  doc.save('bao-cao-san-xuat.pdf')
}

function barWidth(value, max) {
  return `${max ? Math.max(4, Math.round((value / max) * 100)) : 4}%`
}

export function ReportsPage({ orders }) {
  const normalizedOrders = useMemo(() => orders.map(normalizeOrder), [orders])
  const [filters, setFilters] = useState({
    fromDate: '',
    toDate: '',
    orderId: '',
    customer: '',
    product: '',
    status: '',
    actor: '',
    stage: '',
  })

  const filteredOrders = useMemo(() => normalizedOrders.filter((order) => {
    const date = order.productionDate || ''
    return (!filters.fromDate || date >= filters.fromDate)
      && (!filters.toDate || date <= filters.toDate)
      && (!filters.orderId || order.id.toLowerCase().includes(filters.orderId.toLowerCase()))
      && (!filters.customer || order.customer === filters.customer)
      && (!filters.product || order.product.toLowerCase().includes(filters.product.toLowerCase()))
      && (!filters.status || order.status.toLowerCase().includes(filters.status.toLowerCase()))
      && (!filters.actor || order.actor.toLowerCase().includes(filters.actor.toLowerCase()))
      && (!filters.stage || order.currentStage === filters.stage)
  }), [normalizedOrders, filters])

  const kpis = getKpis(filteredOrders)
  const stageRowsRaw = getStageRows(filteredOrders)
  const lossRowsRaw = filteredOrders
  const errorRowsRaw = filteredOrders.flatMap((order) => order.allWeighRows.filter((row) => row.qrStatus === 'FAIL' || row.weighStatus === 'FAIL'))
  const progressRowsRaw = filteredOrders
  const maxVolume = Math.max(...stageRowsRaw.map((row) => row.volume), 1)
  const statusCounts = ['Cân hóa', 'Cân rắn', 'Phối trộn', 'QC', 'Hoàn thành']
    .map((stage) => ({ stage, count: filteredOrders.filter((order) => order.currentStage === stage).length }))
  const passFail = {
    pass: filteredOrders.flatMap((order) => order.allWeighRows).filter((row) => row.qrStatus === 'PASS' && row.weighStatus === 'PASS').length,
    fail: errorRowsRaw.length,
  }

  const kpiRows = [
    ['Tổng số lệnh sản xuất', kpis.totalOrders],
    ['Tổng khối lượng yêu cầu', displayKg(kpis.totalRequired)],
    ['Tổng khối lượng thực tế', displayKg(kpis.totalActual)],
    ['Tỷ lệ hao hụt bình quân', `${kpis.avgLossRate}%`],
    ['Số lệnh đang cân', kpis.weighingOrders],
    ['Số lệnh Ready phối trộn', kpis.readyMixing],
    ['Số lệnh đang QC', kpis.qcOrders],
    ['Số lệnh hoàn thành', kpis.completedOrders],
    ['Số lỗi QR', kpis.qrErrors],
    ['Số lỗi cân ngoài dung sai', kpis.weightErrors],
  ].map(([label, value]) => ({ label, value }))

  const stageRows = stageRowsRaw.map((row) => ({
    'Công đoạn': row.stage,
    'Số lệnh': row.count,
    'Khối lượng': displayKg(row.volume),
    'Số lỗi': row.errors,
    'Thời gian xử lý bình quân': formatDuration(row.avgMinutes),
    'Tỷ lệ hoàn thành': `${row.completionRate}%`,
  }))
  const lossRows = lossRowsRaw.map((order) => ({
    'Mã lô SX': getOrderLotCode(order),
    'Sản phẩm': order.product,
    'Khối lượng yêu cầu': displayKg(order.requiredKg),
    'Khối lượng thực tế': order.actualKg ? displayKg(order.actualKg) : '-',
    'Hao hụt kg': order.actualKg ? displayKg(order.lossKg) : '-',
    'Tỷ lệ hao hụt %': order.actualKg ? `${order.lossRate}%` : '-',
    'Trạng thái': order.status,
    'Ghi chú': order.note || order.mixing?.completionNote || '-',
  }))
  const errorRows = errorRowsRaw.map((row) => ({
    Ngày: row.date || '-',
    'Mã lô SX': row.orderId,
    'Công đoạn': row.stage,
    'Mã vật tư': row.materialCode,
    'Tên nguyên liệu': row.materialName,
    'Khối lượng yêu cầu': displayKg(row.requiredKg),
    'Khối lượng thực tế': row.actualKg ? displayKg(row.actualKg) : '-',
    'Dung sai': displayKg(row.toleranceKg),
    'Sai lệch': row.actualKg ? displayKg(row.variance) : '-',
    'Trạng thái PASS/FAIL': row.qrStatus === 'FAIL' ? 'QR FAIL' : row.weighStatus,
    'Người cân': row.operator || '-',
    'Ghi chú': row.note || '-',
  }))
  const progressRows = progressRowsRaw.map((order) => ({
    'Mã lô SX': getOrderLotCode(order),
    'Khách hàng': order.customer || '-',
    'Sản phẩm': order.product,
    'Trạng thái hiện tại': order.status,
    'Cân hóa': order.scaleStatus?.chemical || 'Pending',
    'Cân rắn': order.scaleStatus?.solid || 'Pending',
    'Phối trộn': order.mixing?.status || 'Pending',
    QC: order.currentStage === 'QC' || order.currentStage === 'Hoàn thành' ? 'Đã chuyển' : 'Pending',
    'Hoàn thành': order.currentStage === 'Hoàn thành' ? 'Completed' : 'Pending',
    'Tổng thời gian xử lý': formatDuration(order.totalMinutes),
  }))

  const updateFilter = (field, value) => setFilters((current) => ({ ...current, [field]: value }))

  return (
    <div className="page-content reports-page">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Báo cáo sản xuất</h2>
            <p className="panel-text">Tổng hợp hiệu suất, khối lượng, trạng thái lệnh, hao hụt, thời gian xử lý và lỗi phát sinh.</p>
          </div>
          <div className="action-row">
            <button className="secondary-button" onClick={() => exportExcel({ summaryRows: kpiRows, stageRows, lossRows, errorRows, progressRows })}>Tải Excel</button>
            <button className="primary-button" onClick={() => exportPdf({ kpiRows, stageRows, lossRows, errorRows, progressRows }, filters)}>Tải PDF</button>
          </div>
        </div>

        <div className="report-filters">
          <label>Từ ngày<input type="date" value={filters.fromDate} onChange={(event) => updateFilter('fromDate', event.target.value)} /></label>
          <label>Đến ngày<input type="date" value={filters.toDate} onChange={(event) => updateFilter('toDate', event.target.value)} /></label>
          <label>Mã lô SX<input value={filters.orderId} onChange={(event) => updateFilter('orderId', event.target.value)} /></label>
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
          <label>Nhóm công đoạn<select value={filters.stage} onChange={(event) => updateFilter('stage', event.target.value)}><option value="">Tất cả</option>{stages.slice(1).map((stage) => <option key={stage}>{stage}</option>)}</select></label>
        </div>
      </section>

      <section className="report-kpi-grid">
        {kpiRows.map((item) => (
          <article className="report-kpi-card" key={item.label}>
            <span>{item.label}</span>
            <strong>{item.value}</strong>
          </article>
        ))}
      </section>

      <section className="report-chart-grid">
        <article className="panel report-chart-card">
          <h3>Khối lượng theo công đoạn</h3>
          {stageRowsRaw.map((row) => (
            <div className="css-bar-row" key={row.stage}>
              <span>{row.stage}</span>
              <div><i style={{ width: barWidth(row.volume, maxVolume) }} /></div>
              <strong>{displayKg(row.volume)}</strong>
            </div>
          ))}
        </article>
        <article className="panel report-chart-card">
          <h3>Trạng thái lệnh</h3>
          {statusCounts.map((row) => (
            <div className="css-bar-row status" key={row.stage}>
              <span>{row.stage}</span>
              <div><i style={{ width: barWidth(row.count, Math.max(...statusCounts.map((item) => item.count), 1)) }} /></div>
              <strong>{row.count}</strong>
            </div>
          ))}
        </article>
        <article className="panel report-chart-card">
          <h3>Lỗi PASS/FAIL</h3>
          <div className="pass-fail-chart">
            <div className="pass" style={{ height: barWidth(passFail.pass, Math.max(passFail.pass, passFail.fail, 1)) }}><strong>{passFail.pass}</strong><span>PASS</span></div>
            <div className="fail" style={{ height: barWidth(passFail.fail, Math.max(passFail.pass, passFail.fail, 1)) }}><strong>{passFail.fail}</strong><span>FAIL</span></div>
          </div>
        </article>
      </section>

      <ReportTable title="Báo cáo theo công đoạn" rows={stageRows} />
      <ReportTable title="Báo cáo hao hụt" rows={lossRows} />
      <ReportTable title="Báo cáo lỗi cân" rows={errorRows} emptyText="Không có lỗi cân trong bộ lọc hiện tại." />
      <ReportTable title="Báo cáo tiến độ lệnh" rows={progressRows} />
    </div>
  )
}

function ReportTable({ title, rows, emptyText = 'Không có dữ liệu.' }) {
  const headers = Object.keys(rows[0] || {})
  return (
    <section className="panel report-table-panel">
      <h3>{title}</h3>
      <div className="table-wrapper">
        <table className="report-wide-table">
          <thead>
            <tr>{headers.map((header) => <th key={header}>{header}</th>)}</tr>
          </thead>
          <tbody>
            {rows.map((row, index) => (
              <tr key={index}>{headers.map((header) => <td key={header}>{row[header]}</td>)}</tr>
            ))}
            {rows.length === 0 && <tr><td colSpan="12" className="empty-row">{emptyText}</td></tr>}
          </tbody>
        </table>
      </div>
    </section>
  )
}
