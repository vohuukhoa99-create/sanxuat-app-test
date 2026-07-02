import { useEffect, useMemo, useState } from 'react'

const priority = {
  Active: 1,
  Ready: 2,
  Pending: 3,
  Completed: 4,
}

function isChemicalGroup(group = '') {
  const value = group.toLowerCase()
  return value.includes('hóa chế')
    || value.includes('hóa chất')
    || value.includes('hoa che')
    || value.includes('hoa chat')
    || value.includes('chemical')
    || value.includes('hóa')
}

function isSolidGroup(group = '') {
  const value = group.toLowerCase()
  return value.includes('nguyên liệu rắn')
    || value.includes('nl rắn')
    || value.includes('nguyen lieu ran')
    || value.includes('nl ran')
    || value.includes('solid')
    || value.includes('rắn')
}

function isScalePass(item) {
  return item.qrStatus === 'PASS' && item.weighStatus === 'PASS'
}

function getScaledIngredients(order) {
  return (order.ingredients || [])
    .filter((item) => isChemicalGroup(item.materialGroup) || isSolidGroup(item.materialGroup))
    .map((item, index) => {
      const requiredKg = Number(item.requiredKg ?? item.materialPerLot ?? 0)
      const actualWeight = item.actualWeight ?? item.actual ?? ''
      return {
        ...item,
        no: index + 1,
        groupLabel: isChemicalGroup(item.materialGroup) ? 'Hóa chế' : 'Nguyên liệu rắn',
        requiredKg,
        actualWeight,
        variance: actualWeight === '' ? '' : Number((Number(actualWeight) - requiredKg).toFixed(3)),
      }
    })
}

function isReadyForMixing(order) {
  const scaledIngredients = getScaledIngredients(order)
  return order.scaleStatus?.chemical === 'Completed'
    && order.scaleStatus?.solid === 'Completed'
    && scaledIngredients.length > 0
    && scaledIngredients.every(isScalePass)
}

function getMixingStatus(order) {
  if (order.mixing?.status) return order.mixing.status
  if (order.stage === 'qc' || order.stage === 'completed') return 'Completed'
  if (order.stage === 'mixing' && isReadyForMixing(order)) return 'Ready'
  if (isReadyForMixing(order)) return 'Ready'
  return 'Pending'
}

function statusClass(status) {
  if (status === 'Ready' || status === 'Completed' || status === 'PASS') return 'pass'
  if (status === 'FAIL') return 'fail'
  if (status === 'Active') return 'active'
  return 'locked'
}

function formatKg(value) {
  const number = Number(value)
  return Number.isFinite(number) ? `${number.toLocaleString('vi-VN')} kg` : '-'
}

function StartMixingModal({ order, onClose, onSubmit }) {
  const [form, setForm] = useState({
    operator: order.mixing?.operator || '',
    mixerNo: order.mixing?.mixerNo || '',
    note: order.mixing?.note || '',
  })

  const updateField = (field, value) => setForm((current) => ({ ...current, [field]: value }))

  return (
    <div className="modal-backdrop" role="presentation">
      <div className="mixing-modal" role="dialog" aria-modal="true">
        <div className="modal-header">
          <div>
            <span className="section-kicker">Bắt đầu phối trộn</span>
            <h2>{order.id}</h2>
          </div>
          <button type="button" className="icon-button" onClick={onClose} aria-label="Đóng">×</button>
        </div>
        <div className="mixing-form-grid">
          <label>Người thực hiện<input value={form.operator} onChange={(event) => updateField('operator', event.target.value)} /></label>
          <label>Máy trộn số<input value={form.mixerNo} onChange={(event) => updateField('mixerNo', event.target.value)} /></label>
          <label className="wide-field">Ghi chú<textarea value={form.note} onChange={(event) => updateField('note', event.target.value)} rows="4" /></label>
        </div>
        <div className="modal-actions">
          <button type="button" className="secondary-button" onClick={onClose}>Hủy</button>
          <button type="button" className="primary-button" onClick={() => onSubmit(form)}>Xác nhận bắt đầu</button>
        </div>
      </div>
    </div>
  )
}

function CompleteMixingModal({ order, onClose, onSubmit }) {
  const [form, setForm] = useState({
    finalWeightKg: order.mixing?.finalWeightKg || '',
    completionNote: order.mixing?.completionNote || '',
    confirmQc: true,
  })

  const updateField = (field, value) => setForm((current) => ({ ...current, [field]: value }))

  return (
    <div className="modal-backdrop" role="presentation">
      <div className="mixing-modal" role="dialog" aria-modal="true">
        <div className="modal-header">
          <div>
            <span className="section-kicker">Hoàn thành phối trộn</span>
            <h2>{order.id}</h2>
          </div>
          <button type="button" className="icon-button" onClick={onClose} aria-label="Đóng">×</button>
        </div>
        <div className="mixing-form-grid">
          <label>Khối lượng sau phối trộn<input type="number" step="0.001" value={form.finalWeightKg} onChange={(event) => updateField('finalWeightKg', event.target.value)} /></label>
          <label className="checkbox-field"><input type="checkbox" checked={form.confirmQc} onChange={(event) => updateField('confirmQc', event.target.checked)} /> Xác nhận chuyển QC</label>
          <label className="wide-field">Ghi chú hoàn thành<textarea value={form.completionNote} onChange={(event) => updateField('completionNote', event.target.value)} rows="4" /></label>
        </div>
        <div className="modal-actions">
          <button type="button" className="secondary-button" onClick={onClose}>Hủy</button>
          <button type="button" className="primary-button" disabled={!form.confirmQc} onClick={() => onSubmit(form)}>Hoàn thành phối trộn</button>
        </div>
      </div>
    </div>
  )
}

export function MixingPage({ orders, onUpdateOrder }) {
  const mixingOrders = useMemo(
    () => orders
      .map((order) => ({ ...order, mixingStatus: getMixingStatus(order) }))
      .sort((a, b) => priority[a.mixingStatus] - priority[b.mixingStatus] || String(a.id).localeCompare(String(b.id))),
    [orders],
  )
  const firstActionOrder = mixingOrders.find((order) => order.mixingStatus === 'Active' || order.mixingStatus === 'Ready')
  const [selectedOrderId, setSelectedOrderId] = useState(firstActionOrder?.id || mixingOrders[0]?.id || '')
  const selectedOrder = mixingOrders.find((order) => order.id === selectedOrderId) || firstActionOrder || mixingOrders[0]
  const scaledIngredients = selectedOrder ? getScaledIngredients(selectedOrder) : []
  const [startOrder, setStartOrder] = useState(null)
  const [completeOrder, setCompleteOrder] = useState(null)

  useEffect(() => {
    mixingOrders
      .filter((order) => order.mixingStatus === 'Ready' && !order.mixing?.status)
      .forEach((order) => {
        const now = new Date().toISOString().slice(0, 16).replace('T', ' ')
        onUpdateOrder(order.id, (current) => ({
          ...current,
          stage: 'mixing',
          status: 'Ready phối trộn',
          mixing: {
            ...(current.mixing || {}),
            status: 'Ready',
            readyAt: now,
          },
        }))
      })
  }, [mixingOrders, onUpdateOrder])

  const startMixing = (order, form) => {
    const now = new Date().toISOString().slice(0, 16).replace('T', ' ')
    onUpdateOrder(order.id, (current) => ({
      ...current,
      stage: 'mixing',
      status: 'Đang phối trộn',
      mixing: {
        ...(current.mixing || {}),
        status: 'Active',
        readyAt: current.mixing?.readyAt || now,
        startedAt: now,
        operator: form.operator,
        mixerNo: form.mixerNo,
        note: form.note,
      },
    }), `Lệnh ${order.id} bắt đầu phối trộn.`)
    setSelectedOrderId(order.id)
    setStartOrder(null)
  }

  const completeMixing = (order, form) => {
    const now = new Date().toISOString().slice(0, 16).replace('T', ' ')
    onUpdateOrder(order.id, (current) => ({
      ...current,
      stage: 'qc',
      status: 'QC',
      mixing: {
        ...(current.mixing || {}),
        status: 'Completed',
        completedAt: now,
        finalWeightKg: Number(form.finalWeightKg) || 0,
        completionNote: form.completionNote,
        confirmQc: form.confirmQc,
      },
    }), `Lệnh ${order.id} hoàn thành phối trộn và chuyển QC.`)
    setSelectedOrderId(order.id)
    setCompleteOrder(null)
  }

  const ensureReady = (order) => {
    if (order.mixing?.readyAt || order.mixingStatus !== 'Ready') return order
    const now = new Date().toISOString().slice(0, 16).replace('T', ' ')
    onUpdateOrder(order.id, (current) => ({
      ...current,
      stage: 'mixing',
      status: 'Ready phối trộn',
      mixing: {
        ...(current.mixing || {}),
        status: 'Ready',
        readyAt: now,
      },
    }))
    return { ...order, mixing: { ...(order.mixing || {}), status: 'Ready', readyAt: now } }
  }

  const handleStartClick = (order) => {
    setStartOrder(ensureReady(order))
  }

  return (
    <div className="page-content mixing-page">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Tổ phối trộn</h2>
            <p className="panel-text">Chỉ lệnh đã hoàn tất cân hóa, cân rắn và toàn bộ QR/Cân PASS mới được đưa vào trạng thái Ready.</p>
          </div>
        </div>

        <div className="table-wrapper">
          <table className="mixing-order-table">
            <thead>
              <tr>
                <th>STT</th>
                <th>Mã lô</th>
                <th>Ngày sản xuất</th>
                <th>Khách hàng</th>
                <th>Tên sản phẩm</th>
                <th>Mã công thức</th>
                <th>Khối lượng yêu cầu</th>
                <th>Số mẻ</th>
                <th>Trạng thái cân hóa</th>
                <th>Trạng thái cân rắn</th>
                <th>Trạng thái phối trộn</th>
                <th>Thời gian Ready</th>
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {mixingOrders.map((order, index) => {
                const isSelected = selectedOrder?.id === order.id
                const readyAt = order.mixing?.readyAt || (order.mixingStatus === 'Ready' ? 'Sẵn sàng' : '-')
                return (
                  <tr key={order.id} className={isSelected ? 'active-order-row' : ''} onClick={() => setSelectedOrderId(order.id)}>
                    <td>{index + 1}</td>
                    <td>{order.lot || order.id}</td>
                    <td>{order.productionDate || order.createdAt || '-'}</td>
                    <td>{order.customer || '-'}</td>
                    <td>{order.product}</td>
                    <td>{order.formula || '-'}</td>
                    <td>{formatKg(order.quantityKg)}</td>
                    <td>{order.batchCount || '-'}</td>
                    <td><span className={`flow-pill ${statusClass(order.scaleStatus?.chemical || 'Pending')}`}>{order.scaleStatus?.chemical || 'Pending'}</span></td>
                    <td><span className={`flow-pill ${statusClass(order.scaleStatus?.solid || 'Pending')}`}>{order.scaleStatus?.solid || 'Pending'}</span></td>
                    <td><span className={`flow-pill ${statusClass(order.mixingStatus)}`}>{order.mixingStatus}</span></td>
                    <td>{readyAt}</td>
                    <td className="action-row">
                      {order.mixingStatus === 'Ready' && (
                        <button className="primary-button" onClick={(event) => { event.stopPropagation(); handleStartClick(order) }}>Bắt đầu phối trộn</button>
                      )}
                      {order.mixingStatus === 'Active' && (
                        <button className="primary-button" onClick={(event) => { event.stopPropagation(); setCompleteOrder(order) }}>Hoàn thành phối trộn</button>
                      )}
                      {order.mixingStatus !== 'Ready' && order.mixingStatus !== 'Active' && <span className="muted-text">Xem chi tiết</span>}
                    </td>
                  </tr>
                )
              })}
              {mixingOrders.length === 0 && (
                <tr><td colSpan="13" className="empty-row">Không có lệnh sản xuất.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      </section>

      {selectedOrder && (
        <section className="panel mixing-detail-panel">
          <div className="section-heading-row">
            <h3>Chi tiết lệnh phối trộn</h3>
            <span className={`flow-pill ${statusClass(selectedOrder.mixingStatus)}`}>{selectedOrder.mixingStatus}</span>
          </div>

          <div className="mixing-info-grid">
            <div><span>Mã lô</span><strong>{selectedOrder.lot || selectedOrder.id}</strong></div>
            <div><span>Tên sản phẩm</span><strong>{selectedOrder.product}</strong></div>
            <div><span>Khách hàng</span><strong>{selectedOrder.customer || '-'}</strong></div>
            <div><span>Khối lượng yêu cầu</span><strong>{formatKg(selectedOrder.quantityKg)}</strong></div>
            <div><span>Số mẻ</span><strong>{selectedOrder.batchCount || '-'}</strong></div>
            <div><span>Công thức</span><strong>{selectedOrder.formula || '-'}</strong></div>
          </div>

          <h3>Tổng hợp nguyên liệu đã cân</h3>
          <div className="table-wrapper">
            <table className="mixing-material-table">
              <thead>
                <tr>
                  <th>STT</th>
                  <th>Nhóm</th>
                  <th>Mã vật tư</th>
                  <th>Tên nguyên liệu</th>
                  <th>Khối lượng yêu cầu</th>
                  <th>Khối lượng thực tế đã cân</th>
                  <th>Sai lệch</th>
                  <th>Trạng thái PASS/FAIL</th>
                  <th>Thời gian cân</th>
                </tr>
              </thead>
              <tbody>
                {scaledIngredients.map((item, index) => (
                  <tr key={item.id}>
                    <td>{index + 1}</td>
                    <td>{item.groupLabel}</td>
                    <td>{item.materialCode}</td>
                    <td>{item.materialName}</td>
                    <td>{formatKg(item.requiredKg)}</td>
                    <td>{item.actualWeight !== '' ? formatKg(item.actualWeight) : '-'}</td>
                    <td>{item.variance !== '' ? formatKg(item.variance) : '-'}</td>
                    <td><span className={`flow-pill ${statusClass(isScalePass(item) ? 'PASS' : 'FAIL')}`}>{isScalePass(item) ? 'PASS' : 'FAIL'}</span></td>
                    <td>{item.confirmedAt || '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>

          <h3>Nhật ký phối trộn</h3>
          <div className="mixing-log-grid">
            <div><span>Thời gian bắt đầu</span><strong>{selectedOrder.mixing?.startedAt || '-'}</strong></div>
            <div><span>Thời gian hoàn thành</span><strong>{selectedOrder.mixing?.completedAt || '-'}</strong></div>
            <div><span>Người thực hiện</span><strong>{selectedOrder.mixing?.operator || '-'}</strong></div>
            <div><span>Máy trộn số</span><strong>{selectedOrder.mixing?.mixerNo || '-'}</strong></div>
            <div className="wide-detail"><span>Ghi chú phối trộn</span><strong>{selectedOrder.mixing?.note || '-'}</strong></div>
            <div className="wide-detail"><span>Ghi chú hoàn thành</span><strong>{selectedOrder.mixing?.completionNote || '-'}</strong></div>
          </div>
        </section>
      )}

      {startOrder && (
        <StartMixingModal order={startOrder} onClose={() => setStartOrder(null)} onSubmit={(form) => startMixing(startOrder, form)} />
      )}
      {completeOrder && (
        <CompleteMixingModal order={completeOrder} onClose={() => setCompleteOrder(null)} onSubmit={(form) => completeMixing(completeOrder, form)} />
      )}
    </div>
  )
}
