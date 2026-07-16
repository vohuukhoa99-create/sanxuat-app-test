import { useMemo, useState } from 'react'
import { CustomerFilterCombobox } from '../components/CustomerFilterCombobox.jsx'
import { customerCatalog } from '../data/customerCatalog.js'

const customerOptions = Array.from(new Set((customerCatalog || []).map((customer) => customer.customerName).filter(Boolean))).sort((a, b) => a.localeCompare(b, 'vi', { numeric: true }))

const pipelineStages = ['Nhận lệnh', 'Cân hóa', 'Cân rắn', 'Phối trộn', 'QC', 'Hoàn thành']
const statusLabels = ['Chờ cân', 'Đang cân', 'Ready phối trộn', 'Đang phối trộn', 'QC', 'Hoàn thành', 'Lỗi']

const demoOrders = Array.from({ length: 15 }, (_, index) => {
  const id = `DEMO-${String(index + 1).padStart(3, '0')}`
  const products = ['HSS 251.023', 'Sơn lót HB-01', 'Sơn phủ HB-22', 'Sơn chống thấm CT-9', 'Sơn công nghiệp IND-5']
  const stages = ['chemical', 'solid', 'mixing', 'qc', 'completed']
  const stage = stages[index % stages.length]
  const quantityKg = [500, 750, 1000, 1250, 1500][index % 5]
  const finalWeight = stage === 'completed' || stage === 'qc' ? quantityKg - (index % 4) * 12 : 0
  const hasQrFail = index === 2 || index === 9
  const hasWeightFail = index === 4 || index === 11
  const date = new Date()
  date.setDate(date.getDate() - (index % 10))
  const productionDate = date.toISOString().slice(0, 10)
  const group = index % 2 === 0 ? 'Sơn công nghiệp' : 'Sơn dân dụng'
  const customer = ['CT-KH A', 'CT-KH B', 'CT-KH C'][index % 3]
  const owner = ['Nguyễn An', 'Trần Bình', 'Lê Chi'][index % 3]

  return {
    id,
    productionDate,
    customer,
    product: products[index % products.length],
    productGroup: group,
    formula: `F-${index + 100}`,
    lot: `LOT-${index + 1}`,
    quantityKg,
    stage,
    status: stage === 'completed' ? 'Hoàn thành' : stage === 'qc' ? 'QC' : stage === 'mixing' ? 'Ready phối trộn' : 'Đang cân',
    owner,
    createdAt: `${productionDate} 08:00`,
    updatedAt: `${productionDate} ${10 + (index % 8)}:15`,
    scaleStatus: {
      chemical: ['Completed', 'Active', 'Completed'][index % 3],
      solid: ['Pending', 'Completed', 'Active', 'Completed'][index % 4],
    },
    mixing: stage === 'mixing' || stage === 'qc' || stage === 'completed'
      ? {
        status: stage === 'mixing' ? (index % 2 ? 'Ready' : 'Active') : 'Completed',
        readyAt: `${productionDate} 11:00`,
        startedAt: `${productionDate} 11:30`,
        completedAt: stage === 'mixing' ? '' : `${productionDate} 13:20`,
        operator: owner,
        finalWeightKg: finalWeight,
      }
      : null,
    qc: stage === 'completed' ? { result: index % 7 === 0 ? 'FAIL' : 'PASS', checkedAt: `${productionDate} 15:00` } : null,
    completedAt: stage === 'completed' ? `${productionDate} 16:00` : '',
    ingredients: [
      {
        id: `${id}-chem`,
        materialCode: 'PASTE 02',
        materialName: 'PASTE 02',
        materialGroup: 'Hóa chế',
        requiredKg: 2.31,
        actualWeight: hasWeightFail ? 2.5 : 2.31,
        toleranceKg: 0.01,
        qrStatus: hasQrFail ? 'FAIL' : 'PASS',
        weighStatus: hasWeightFail ? 'FAIL' : 'PASS',
        confirmedAt: `${productionDate} 09:00`,
        operator: owner,
        note: hasQrFail ? 'Sai QR' : hasWeightFail ? 'Ngoài dung sai' : '',
      },
      {
        id: `${id}-solid`,
        materialCode: 'R91',
        materialName: 'R91',
        materialGroup: 'Nguyên liệu rắn',
        requiredKg: 10,
        actualWeight: 10,
        toleranceKg: 0.1,
        qrStatus: 'PASS',
        weighStatus: 'PASS',
        confirmedAt: `${productionDate} 10:00`,
        operator: owner,
        note: '',
      },
    ],
  }
})

function parseDate(value) {
  const date = new Date(String(value || '').replace(' ', 'T'))
  return Number.isNaN(date.getTime()) ? null : date
}

function ageText(order) {
  const start = parseDate(order.updatedAt || order.createdAt)
  if (!start) return '-'
  const minutes = Math.max(0, Math.round((Date.now() - start.getTime()) / 60000))
  if (minutes < 60) return `${minutes} phút`
  const hours = Math.floor(minutes / 60)
  if (hours < 24) return `${hours} giờ`
  return `${Math.floor(hours / 24)} ngày`
}

function durationMinutes(start, end) {
  const a = parseDate(start)
  const b = parseDate(end)
  if (!a || !b) return 0
  return Math.max(0, Math.round((b - a) / 60000))
}

function formatDuration(minutes) {
  if (!minutes) return '-'
  const hours = Math.floor(minutes / 60)
  const rest = minutes % 60
  return hours ? `${hours}h ${rest}m` : `${rest}m`
}

function kg(value) {
  const number = Number(value)
  return Number.isFinite(number) ? number : 0
}

function kgText(value) {
  return `${kg(value).toLocaleString('vi-VN')} kg`
}

function normalizeOrder(order) {
  const ingredients = order.ingredients || []
  const qrFail = ingredients.filter((item) => item.qrStatus === 'FAIL').length
  const weighFail = ingredients.filter((item) => item.weighStatus === 'FAIL').length
  const actualKg = kg(order.mixing?.finalWeightKg ?? order.actualWeightKg)
  const requiredKg = kg(order.quantityKg)
  const lossKg = actualKg ? requiredKg - actualKg : 0
  const lossRate = actualKg && requiredKg ? Number(((lossKg / requiredKg) * 100).toFixed(2)) : 0
  const qcFail = order.qc?.result === 'FAIL' || String(order.qcStatus || '').toUpperCase() === 'FAIL'
  const currentStage = getCurrentStage(order, qrFail, weighFail, qcFail)

  return {
    ...order,
    productionDate: order.productionDate || String(order.createdAt || '').slice(0, 10),
    customer: order.customer || '',
    productGroup: order.productGroup || 'Sơn công nghiệp',
    owner: order.owner || order.mixing?.operator || '',
    requiredKg,
    actualKg,
    lossKg,
    lossRate,
    qrFail,
    weighFail,
    qcFail,
    currentStage,
    isCompleted: currentStage === 'Hoàn thành',
    totalMinutes: durationMinutes(order.createdAt, order.completedAt || order.qc?.checkedAt || order.mixing?.completedAt || order.updatedAt),
  }
}

function getCurrentStage(order, qrFail, weighFail, qcFail) {
  if (qrFail || weighFail || qcFail) return 'Lỗi'
  if (order.stage === 'completed') return 'Hoàn thành'
  if (order.stage === 'qc') return 'QC'
  if (order.mixing?.status === 'Active') return 'Đang phối trộn'
  if (order.mixing?.status === 'Ready' || String(order.status || '').includes('Ready')) return 'Ready phối trộn'
  if (order.scaleStatus?.chemical === 'Active' || order.scaleStatus?.solid === 'Active') return 'Đang cân'
  if (order.stage === 'chemical' || order.stage === 'solid') return 'Đang cân'
  return 'Chờ cân'
}

function pipelineStage(order) {
  if (order.currentStage === 'Hoàn thành') return 'Hoàn thành'
  if (order.currentStage === 'QC') return 'QC'
  if (order.currentStage === 'Đang phối trộn' || order.currentStage === 'Ready phối trộn') return 'Phối trộn'
  if (order.stage === 'solid' || order.scaleStatus?.solid === 'Active') return 'Cân rắn'
  if (order.stage === 'chemical' || order.scaleStatus?.chemical === 'Active') return 'Cân hóa'
  return 'Nhận lệnh'
}

function getAlerts(orders) {
  const alerts = []
  orders.forEach((order) => {
    if (order.totalMinutes > 360 && !order.isCompleted) alerts.push({ order, stage: order.currentStage, reason: 'Kẹt quá thời gian chuẩn', level: 'Cao' })
    if (order.weighFail) alerts.push({ order, stage: 'Cân', reason: 'Có dòng cân FAIL', level: 'Cao' })
    if (order.qrFail) alerts.push({ order, stage: 'QR', reason: 'Có QR FAIL', level: 'Cao' })
    if (order.lossRate > 2) alerts.push({ order, stage: 'Hao hụt', reason: `Hao hụt ${order.lossRate}%`, level: 'Trung bình' })
    if (order.currentStage === 'QC') alerts.push({ order, stage: 'QC', reason: 'QC chưa xử lý', level: 'Trung bình' })
    if (order.currentStage === 'Ready phối trộn') alerts.push({ order, stage: 'Phối trộn', reason: 'Ready nhưng chưa bắt đầu', level: 'Thấp' })
  })
  return alerts.slice(0, 8)
}

function groupByPeriod(orders, period) {
  const bucket = {}
  orders.forEach((order) => {
    const date = parseDate(order.productionDate)
    if (!date) return
    let key = order.productionDate
    if (period === 'week') key = `${date.getFullYear()}-W${Math.ceil((date.getDate() + 6) / 7)}`
    if (period === 'month') key = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`
    bucket[key] = (bucket[key] || 0) + order.requiredKg
  })
  return Object.entries(bucket).sort(([a], [b]) => a.localeCompare(b)).slice(-10).map(([label, value]) => ({ label, value }))
}

export function DashboardPage({ orders }) {
  const [viewMode, setViewMode] = useState('ceo')
  const [volumeMode, setVolumeMode] = useState('day')
  const [filters, setFilters] = useState({
    fromDate: '',
    toDate: '',
    shift: '',
    productGroup: '',
    customer: '',
    status: '',
    stage: '',
    owner: '',
  })
  const sourceOrders = orders.length >= 15 ? orders : [...orders, ...demoOrders].slice(0, 20)
  const normalized = useMemo(() => sourceOrders.map(normalizeOrder), [sourceOrders])
  const filtered = useMemo(() => normalized.filter((order) => {
    const date = order.productionDate || ''
    return (!filters.fromDate || date >= filters.fromDate)
      && (!filters.toDate || date <= filters.toDate)
      && (!filters.productGroup || order.productGroup === filters.productGroup)
      && (!filters.customer || order.customer === filters.customer)
      && (!filters.status || order.currentStage === filters.status)
      && (!filters.stage || pipelineStage(order) === filters.stage)
      && (!filters.owner || order.owner.toLowerCase().includes(filters.owner.toLowerCase()))
  }), [normalized, filters])

  const totalRequired = filtered.reduce((sum, order) => sum + order.requiredKg, 0)
  const totalActual = filtered.reduce((sum, order) => sum + order.actualKg, 0)
  const completed = filtered.filter((order) => order.isCompleted).length
  const processing = filtered.filter((order) => !order.isCompleted).length
  const stuck = getAlerts(filtered).filter((alert) => alert.reason.includes('Kẹt')).length
  const qrFail = filtered.reduce((sum, order) => sum + order.qrFail, 0)
  const weighFail = filtered.reduce((sum, order) => sum + order.weighFail, 0)
  const qcWaiting = filtered.filter((order) => order.currentStage === 'QC').length
  const qcFail = filtered.filter((order) => order.qcFail).length
  const avgLoss = filtered.length ? filtered.reduce((sum, order) => sum + order.lossRate, 0) / filtered.length : 0
  const alerts = getAlerts(filtered)
  const volumeSeries = groupByPeriod(filtered, volumeMode)
  const maxVolume = Math.max(...volumeSeries.map((item) => item.value), 1)
  const pipeline = pipelineStages.map((stage) => {
    const rows = filtered.filter((order) => pipelineStage(order) === stage)
    return { stage, count: rows.length, volume: rows.reduce((sum, order) => sum + order.requiredKg, 0), rate: filtered.length ? Math.round((rows.length / filtered.length) * 100) : 0 }
  })
  const statusCounts = statusLabels.map((status) => ({ status, count: filtered.filter((order) => order.currentStage === status).length }))
  const attentionRows = alerts.map((alert) => ({
    ...alert,
    issue: alert.reason,
    age: ageText(alert.order),
  }))
  const stageEfficiency = pipelineStages.map((stage) => {
    const rows = filtered.filter((order) => pipelineStage(order) === stage)
    const errors = rows.reduce((sum, order) => sum + order.qrFail + order.weighFail + (order.qcFail ? 1 : 0), 0)
    const done = rows.filter((order) => order.isCompleted || ['Cân hóa', 'Cân rắn', 'Phối trộn'].includes(stage)).length
    const avgMinutes = rows.length ? Math.round(rows.reduce((sum, order) => sum + order.totalMinutes, 0) / rows.length) : 0
    const completionRate = rows.length ? Math.round((done / rows.length) * 100) : 0
    return { stage, count: rows.length, volume: rows.reduce((sum, order) => sum + order.requiredKg, 0), avgMinutes, errors, completionRate, rating: errors > 2 ? 'Cảnh báo' : avgMinutes > 360 ? 'Cần theo dõi' : 'Tốt' }
  })
  const productRows = Object.values(filtered.reduce((acc, order) => {
    acc[order.product] ||= { product: order.product, count: 0, required: 0, actual: 0, loss: 0, qcPass: 0, qcTotal: 0 }
    acc[order.product].count += 1
    acc[order.product].required += order.requiredKg
    acc[order.product].actual += order.actualKg
    acc[order.product].loss += order.lossKg
    if (order.currentStage === 'QC' || order.isCompleted) acc[order.product].qcTotal += 1
    if (order.qc?.result === 'PASS' || order.isCompleted) acc[order.product].qcPass += 1
    return acc
  }, {}))

  const kpis = [
    ['Tổng lệnh sản xuất', filtered.length],
    ['Lệnh hoàn thành', completed],
    ['Lệnh đang xử lý', processing],
    ['Lệnh trễ / kẹt', stuck],
    ['Tổng khối lượng yêu cầu', kgText(totalRequired)],
    ['Tổng khối lượng thực tế', kgText(totalActual)],
    ['Tỷ lệ hoàn thành', `${filtered.length ? Math.round((completed / filtered.length) * 100) : 0}%`],
    ['Tỷ lệ hao hụt bình quân', `${avgLoss.toFixed(2)}%`],
    ['Số lỗi QR', qrFail],
    ['Số lỗi cân ngoài dung sai', weighFail],
    ['Số lệnh chờ QC', qcWaiting],
    ['Số lệnh QC không đạt', qcFail],
  ]

  const updateFilter = (field, value) => setFilters((current) => ({ ...current, [field]: value }))

  return (
    <div className="page-content executive-dashboard">
      <section className="panel dashboard-control-panel">
        <div className="section-heading-row">
          <div>
            <h2>Dashboard điều hành</h2>
            <p className="panel-text">Tổng hợp nhanh sản lượng, lệnh kẹt, lỗi, hao hụt và QC tồn đọng.</p>
          </div>
          <div className="view-switch">
            <button className={viewMode === 'ceo' ? 'active' : ''} onClick={() => setViewMode('ceo')}>Góc nhìn TGĐ</button>
            <button className={viewMode === 'production' ? 'active' : ''} onClick={() => setViewMode('production')}>Góc nhìn GĐ Sản xuất</button>
          </div>
        </div>
        <div className="executive-filters">
          <label>Từ ngày<input type="date" value={filters.fromDate} onChange={(event) => updateFilter('fromDate', event.target.value)} /></label>
          <label>Đến ngày<input type="date" value={filters.toDate} onChange={(event) => updateFilter('toDate', event.target.value)} /></label>
          <label>Ca sản xuất<select value={filters.shift} onChange={(event) => updateFilter('shift', event.target.value)}><option value="">Tất cả</option><option>Ca 1</option><option>Ca 2</option><option>Ca 3</option></select></label>
          <label>Nhóm sản phẩm<select value={filters.productGroup} onChange={(event) => updateFilter('productGroup', event.target.value)}><option value="">Tất cả</option><option>Sơn công nghiệp</option><option>Sơn dân dụng</option></select></label>
          <label>Khách hàng
            <CustomerFilterCombobox
              options={customerOptions}
              value={filters.customer}
              emptyValue=""
              onChange={(customer) => updateFilter('customer', customer)}
            />
          </label>
          <label>Trạng thái lệnh<select value={filters.status} onChange={(event) => updateFilter('status', event.target.value)}><option value="">Tất cả</option>{statusLabels.map((status) => <option key={status}>{status}</option>)}</select></label>
          <label>Công đoạn<select value={filters.stage} onChange={(event) => updateFilter('stage', event.target.value)}><option value="">Tất cả</option>{pipelineStages.map((stage) => <option key={stage}>{stage}</option>)}</select></label>
          <label>Người phụ trách<input value={filters.owner} onChange={(event) => updateFilter('owner', event.target.value)} /></label>
        </div>
      </section>

      <section className="executive-kpi-grid">
        {kpis.map(([label, value], index) => (
          <article className={`executive-kpi ${index === 3 || index >= 8 ? 'warn' : ''}`} key={label}>
            <span>{label}</span>
            <strong>{value}</strong>
          </article>
        ))}
      </section>

      <section className="dashboard-main-grid">
        <article className="panel alert-panel">
          <h3>Cảnh báo cần xử lý</h3>
          <div className="alert-list">
            {alerts.map((alert, index) => (
              <div className={`alert-item ${alert.level === 'Cao' ? 'high' : alert.level === 'Trung bình' ? 'medium' : 'low'}`} key={`${alert.order.id}-${index}`}>
                <strong>{alert.order.id}</strong>
                <span>{alert.stage}</span>
                <p>{alert.reason}</p>
                <small>{ageText(alert.order)}</small>
                <em>{alert.level}</em>
              </div>
            ))}
            {alerts.length === 0 && <div className="empty-alert">Không có cảnh báo nghiêm trọng.</div>}
          </div>
        </article>

        <article className="panel pipeline-panel">
          <h3>Tiến độ lệnh theo công đoạn</h3>
          <div className="pipeline-flow">
            {pipeline.map((item) => (
              <div className="pipeline-step" key={item.stage}>
                <span>{item.stage}</span>
                <strong>{item.count}</strong>
                <small>{kgText(item.volume)}</small>
                <em>{item.rate}%</em>
              </div>
            ))}
          </div>
        </article>
      </section>

      <section className="dashboard-chart-grid">
        <article className="panel dashboard-chart-card">
          <div className="section-heading-row">
            <h3>Khối lượng sản xuất</h3>
            <div className="mini-switch">
              <button className={volumeMode === 'day' ? 'active' : ''} onClick={() => setVolumeMode('day')}>Ngày</button>
              <button className={volumeMode === 'week' ? 'active' : ''} onClick={() => setVolumeMode('week')}>Tuần</button>
              <button className={volumeMode === 'month' ? 'active' : ''} onClick={() => setVolumeMode('month')}>Tháng</button>
            </div>
          </div>
          <div className="volume-bars">
            {volumeSeries.map((item) => (
              <div key={item.label}>
                <i style={{ height: `${Math.max(8, (item.value / maxVolume) * 100)}%` }} />
                <span>{item.label}</span>
                <strong>{kg(item.value)}</strong>
              </div>
            ))}
          </div>
        </article>

        <article className="panel dashboard-chart-card">
          <h3>Trạng thái lệnh</h3>
          <div className="status-stack">
            {statusCounts.map((item) => (
              <div key={item.status} style={{ flexGrow: Math.max(item.count, 0.4) }}><span>{item.status}</span><strong>{item.count}</strong></div>
            ))}
          </div>
        </article>

        <article className="panel dashboard-chart-card">
          <h3>Lỗi phát sinh</h3>
          <div className="error-bars">
            {[
              ['QR FAIL', qrFail],
              ['Cân FAIL', weighFail],
              ['QC FAIL', qcFail],
              ['Hao hụt vượt chuẩn', filtered.filter((order) => order.lossRate > 2).length],
            ].map(([label, value]) => (
              <div key={label}><span>{label}</span><i style={{ width: `${Math.max(6, (Number(value) / Math.max(qrFail, weighFail, qcFail, 1)) * 100)}%` }} /><strong>{value}</strong></div>
            ))}
          </div>
        </article>
      </section>

      <section className="dashboard-table-grid">
        <DashboardTable title="Top mã lô SX cần chú ý" headers={['Mã lô SX', 'Khách hàng', 'Sản phẩm', 'Khối lượng', 'Trạng thái', 'Công đoạn', 'Vấn đề', 'Thời gian tồn', 'Người phụ trách', 'Hành động']}>
          {attentionRows.map((row) => (
            <tr key={`${row.order.id}-${row.issue}`}>
              <td>{row.order.lot || row.order.orderCode || row.order.id}</td><td>{row.order.customer || '-'}</td><td>{row.order.product}</td><td>{kgText(row.order.requiredKg)}</td><td>{row.order.currentStage}</td><td>{row.stage}</td><td>{row.issue}</td><td>{row.age}</td><td>{row.order.owner || '-'}</td><td><button className="secondary-button">Xem chi tiết</button></td>
            </tr>
          ))}
        </DashboardTable>

        {viewMode === 'production' && (
          <DashboardTable title="Hiệu suất công đoạn" headers={['Công đoạn', 'Số lệnh xử lý', 'Khối lượng xử lý', 'Thời gian bình quân', 'Số lỗi', 'Tỷ lệ hoàn thành', 'Đánh giá']}>
            {stageEfficiency.map((row) => (
              <tr key={row.stage}>
                <td>{row.stage}</td><td>{row.count}</td><td>{kgText(row.volume)}</td><td>{formatDuration(row.avgMinutes)}</td><td>{row.errors}</td><td>{row.completionRate}%</td><td><span className={`flow-pill ${row.rating === 'Tốt' ? 'pass' : row.rating === 'Cảnh báo' ? 'fail' : 'active'}`}>{row.rating}</span></td>
              </tr>
            ))}
          </DashboardTable>
        )}

        <DashboardTable title="Sản lượng theo sản phẩm" headers={['Sản phẩm', 'Số lệnh', 'Khối lượng yêu cầu', 'Khối lượng thực tế', 'Hao hụt', 'Tỷ lệ đạt QC']}>
          {productRows.map((row) => (
            <tr key={row.product}>
              <td>{row.product}</td><td>{row.count}</td><td>{kgText(row.required)}</td><td>{kgText(row.actual)}</td><td>{kgText(row.loss)}</td><td>{row.qcTotal ? Math.round((row.qcPass / row.qcTotal) * 100) : 0}%</td>
            </tr>
          ))}
        </DashboardTable>
      </section>
    </div>
  )
}

function DashboardTable({ title, headers, children }) {
  return (
    <section className="panel dashboard-table-panel">
      <h3>{title}</h3>
      <div className="table-wrapper">
        <table className="dashboard-wide-table">
          <thead><tr>{headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
          <tbody>{children}</tbody>
        </table>
      </div>
    </section>
  )
}
