import { useMemo, useState } from 'react'

const QR_WAITING = 'Chờ quét'
const WEIGH_WAITING = 'Chờ cân'

const formatKg = (value) => `${Number(value || 0).toLocaleString('vi-VN', { minimumFractionDigits: 3, maximumFractionDigits: 3 })} kg`
const formatToleranceKg = (value) => value === '' || value === null || value === undefined ? '-' : formatKg(value)
const hasFormulaTolerance = (item = {}) => item.toleranceKg !== '' && item.toleranceKg !== null && item.toleranceKg !== undefined && Number.isFinite(Number(item.toleranceKg))
const toleranceErrorMessage = (requiredKg, actualWeight, toleranceKg) => `Khối lượng ngoài dung sai. Cần cân: ${formatKg(requiredKg)}, Dung sai: ${formatToleranceKg(toleranceKg)}`

function isSolidGroup(group = '') {
  const value = group.toLowerCase()
  return value.includes('nguyên liệu rắn')
    || value.includes('nl rắn')
    || value.includes('nguyen lieu ran')
    || value.includes('nl ran')
    || value.includes('solid')
    || value.includes('rắn')
}

function getSolidIngredients(order) {
  return (order.ingredients || [])
    .filter((item) => isSolidGroup(item.materialGroup))
    .map((item, index) => normalizeSolidItem(item, index))
}

function normalizeSolidItem(item, index) {
  const requiredKg = Number(item.requiredKg ?? item.materialPerLot ?? 0)
  return {
    ...item,
    no: item.no || index + 1,
    requiredKg,
    toleranceKg: item.toleranceKg ?? item.tolerance ?? '',
    qrScanned: item.qrScanned || '',
    qrStatus: item.qrStatus || QR_WAITING,
    actualWeight: item.actualWeight ?? item.actual ?? '',
    weighStatus: item.weighStatus || WEIGH_WAITING,
    confirmedAt: item.confirmedAt || '',
    note: item.note || '',
  }
}

function getSolidStatus(order) {
  const solids = getSolidIngredients(order)
  const stored = order.scaleStatus?.solid
  if (stored === 'Active' || stored === 'Completed') return stored
  if (solids.length === 0) return 'Pending'
  if (solids.every((item) => item.qrStatus === 'PASS' && item.weighStatus === 'PASS')) return 'Completed'
  if (order.stage === 'solid' || order.scaleStatus?.chemical === 'Completed') return 'Ready'
  return stored || 'Pending'
}

function getFirstOpenIndex(items) {
  return items.findIndex((item) => !(item.qrStatus === 'PASS' && item.weighStatus === 'PASS'))
}

function statusClass(status) {
  if (status === 'PASS' || status === 'Completed' || status === 'Ready') return 'pass'
  if (status === 'FAIL') return 'fail'
  if (status === 'Active') return 'active'
  if (status === 'Locked' || status === 'Pending') return 'locked'
  return 'waiting'
}

function updateOrderIngredient(order, itemId, patch) {
  return {
    ...order,
    ingredients: (order.ingredients || []).map((item) =>
      item.id === itemId ? { ...item, ...patch } : item,
    ),
  }
}

export function SolidScalePage({ orders, onUpdateOrder }) {
  const solidOrders = useMemo(
    () => orders
      .filter((order) => getSolidIngredients(order).length > 0 || order.stage === 'solid')
      .map((order) => ({ ...order, solidStatus: getSolidStatus(order) })),
    [orders],
  )
  const storedActiveOrder = solidOrders.find((order) => order.solidStatus === 'Active')
  const [activeOrderId, setActiveOrderId] = useState(storedActiveOrder?.id || solidOrders[0]?.id || '')
  const activeOrder = solidOrders.find((order) => order.id === activeOrderId)
    || storedActiveOrder
    || solidOrders[0]
  const solids = activeOrder ? getSolidIngredients(activeOrder) : []
  const firstOpenIndex = getFirstOpenIndex(solids)
  const currentIndex = firstOpenIndex === -1 ? solids.length : firstOpenIndex
  const currentItem = solids[currentIndex]
  const [qrInput, setQrInput] = useState('')
  const [actualInput, setActualInput] = useState('')
  const [warning, setWarning] = useState('')

  const startOrder = (order) => {
    setActiveOrderId(order.id)
    setQrInput('')
    setActualInput('')
    setWarning('')
    onUpdateOrder(
      order.id,
      (current) => ({
        ...current,
        scaleStatus: { ...(current.scaleStatus || {}), solid: 'Active' },
        status: 'Active',
        solidStartedAt: current.solidStartedAt || new Date().toISOString().slice(0, 16).replace('T', ' '),
      }),
      `Lệnh ${order.id} bắt đầu cân rắn.`,
    )
  }

  const confirmQr = () => {
    if (!activeOrder || !currentItem) return
    const scanned = qrInput.trim()
    const expected = String(currentItem.materialCode || '').trim()
    if (!scanned) {
      setWarning('Vui lòng nhập mã QR/Mã vật tư quét được')
      return
    }
    if (scanned.toUpperCase() !== expected.toUpperCase()) {
      setWarning('Sai mã vật tư, vui lòng kiểm tra lại')
      onUpdateOrder(activeOrder.id, (order) =>
        updateOrderIngredient(order, currentItem.id, {
          qrScanned: scanned,
          qrStatus: 'FAIL',
          weighStatus: WEIGH_WAITING,
          note: 'Sai mã vật tư, vui lòng kiểm tra lại',
        }),
      )
      return
    }
    setWarning('')
    setActualInput('')
    onUpdateOrder(activeOrder.id, (order) =>
      updateOrderIngredient(order, currentItem.id, {
        qrScanned: scanned,
        qrStatus: 'PASS',
        note: '',
      }),
    )
  }

  const retryWeight = () => {
    if (!activeOrder || !currentItem) return
    setWarning('')
    setActualInput('')
    onUpdateOrder(activeOrder.id, (order) =>
      updateOrderIngredient(order, currentItem.id, {
        actualWeight: '',
        weighStatus: WEIGH_WAITING,
        confirmedAt: '',
        note: '',
      }),
    )
  }

  const confirmWeight = () => {
    if (!activeOrder || !currentItem) return
    const actualWeight = Number(actualInput)
    if (!Number.isFinite(actualWeight)) {
      setWarning('Vui lòng nhập khối lượng thực tế')
      return
    }
    if (!hasFormulaTolerance(currentItem)) {
      setWarning('Chưa có dung sai')
      return
    }
    const toleranceKg = Number(currentItem.toleranceKg)
    const min = currentItem.requiredKg - toleranceKg
    const max = currentItem.requiredKg + toleranceKg
    const now = new Date().toISOString().slice(0, 16).replace('T', ' ')

    if (actualWeight < min || actualWeight > max) {
      setWarning(toleranceErrorMessage(currentItem.requiredKg, actualWeight, toleranceKg))
      return
    }

    setWarning('')
    setQrInput('')
    setActualInput('')
    onUpdateOrder(activeOrder.id, (order, updatedAt) => {
      const nextOrder = updateOrderIngredient(order, currentItem.id, {
        actualWeight,
        weighStatus: 'PASS',
        confirmedAt: now,
        note: '',
      })
      const nextSolids = getSolidIngredients(nextOrder)
      const isCompleted = nextSolids.every((item) => item.qrStatus === 'PASS' && item.weighStatus === 'PASS')
      if (!isCompleted) return nextOrder

      const chemicalCompleted = nextOrder.scaleStatus?.chemical === 'Completed'
      return {
        ...nextOrder,
        stage: chemicalCompleted ? 'mixing' : 'chemical',
        status: chemicalCompleted ? 'Ready phối trộn' : 'Chờ cân hóa',
        scaleStatus: {
          ...(nextOrder.scaleStatus || {}),
          solid: 'Completed',
        },
        solidCompletedAt: updatedAt,
      }
    }, `Lệnh ${activeOrder.id} xác nhận cân rắn ${currentItem.materialCode}.`)
  }

  return (
    <div className="page-content chemical-workflow">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Tổ cân rắn</h2>
            <p className="panel-text">Quy trình cân tuần tự nguyên liệu rắn: quét QR, xác nhận mã, cân khối lượng, rồi PASS/FAIL.</p>
          </div>
        </div>

        <div className="chemical-layout">
          <aside className="order-queue">
            <h3>A. Danh sách lệnh sản xuất</h3>
            <div className="table-wrapper">
              <table className="queue-table">
                <thead>
                  <tr>
                    <th>Mã lệnh SX</th>
                    <th>Tên sản phẩm</th>
                    <th>Khối lượng</th>
                    <th>Trạng thái</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {solidOrders.map((order) => (
                    <tr key={order.id} className={activeOrder?.id === order.id ? 'active-order-row' : ''}>
                      <td>{order.id}</td>
                      <td>{order.product}</td>
                      <td>{order.quantityKg} kg</td>
                      <td><span className={`flow-pill ${statusClass(order.solidStatus)}`}>{order.solidStatus}</span></td>
                      <td>
                        <button
                          className="primary-button"
                          disabled={order.solidStatus === 'Completed'}
                          onClick={() => startOrder(order)}
                        >
                          Bắt đầu cân
                        </button>
                      </td>
                    </tr>
                  ))}
                  {solidOrders.length === 0 && (
                    <tr><td colSpan="5" className="empty-row">Không có lệnh cần cân rắn.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </aside>

          <section className="weighing-board">
            <div className="section-heading-row">
              <div>
                <h3>B. Khu vực thao tác cân</h3>
                <p className="panel-text">
                  {activeOrder ? `${activeOrder.id} - ${activeOrder.product}` : 'Chưa chọn lệnh sản xuất'}
                </p>
              </div>
              {activeOrder?.solidCompletedAt && (
                <span className="flow-pill pass">Hoàn thành: {activeOrder.solidCompletedAt}</span>
              )}
            </div>

            <div className="workflow-steps">
              <span className="step active">Bước 1: Quét QR</span>
              <span className="step active">Bước 2: Xác nhận đúng mã vật tư</span>
              <span className={currentItem?.qrStatus === 'PASS' ? 'step active' : 'step locked'}>Bước 3: Cân khối lượng</span>
              <span className="step">Bước 4: PASS/FAIL</span>
            </div>

            {warning && <div className="process-alert">{warning}</div>}

            {currentItem && activeOrder?.solidStatus === 'Active' && (
              <div className="active-operation">
                <div className="operation-field">
                  <label>Mã QR/Mã vật tư quét được</label>
                  <div className="operation-control">
                    <input value={qrInput} onChange={(event) => setQrInput(event.target.value)} placeholder={currentItem.materialCode} />
                    <button className="primary-button" onClick={confirmQr}>Xác nhận QR</button>
                  </div>
                </div>
                <div className="operation-field">
                  <label>Khối lượng thực tế</label>
                  <div className="operation-control">
                    <input
                      type="number"
                      step="0.001"
                      value={actualInput}
                      disabled={currentItem.qrStatus !== 'PASS'}
                      onChange={(event) => setActualInput(event.target.value)}
                      placeholder={`${currentItem.requiredKg} kg`}
                    />
                    <button className="primary-button" disabled={currentItem.qrStatus !== 'PASS'} onClick={confirmWeight}>Xác nhận cân</button>
                    {currentItem.weighStatus === 'FAIL' && <button className="secondary-button" onClick={retryWeight}>Cân lại</button>}
                  </div>
                </div>
              </div>
            )}

            <div className="table-wrapper">
              <table className="chemical-process-table">
                <thead>
                  <tr>
                    <th>STT</th>
                    <th>Mã vật tư</th>
                    <th>Tên nguyên liệu rắn</th>
                    <th>Khối lượng yêu cầu</th>
                    <th>Dung sai cho phép</th>
                    <th>QR đã quét</th>
                    <th>Trạng thái QR</th>
                    <th>Khối lượng thực tế</th>
                    <th>Trạng thái cân</th>
                    <th>Thời gian xác nhận</th>
                    <th>Ghi chú</th>
                  </tr>
                </thead>
                <tbody>
                  {solids.map((item, index) => {
                    const isDone = item.qrStatus === 'PASS' && item.weighStatus === 'PASS'
                    const missingTolerance = item.qrStatus === 'PASS' && !hasFormulaTolerance(item)
                    const isCurrent = activeOrder?.solidStatus === 'Active' && index === currentIndex
                    const isLocked = activeOrder?.solidStatus !== 'Active' || index > currentIndex
                    return (
                      <tr key={item.id} className={item.weighStatus === 'FAIL' || missingTolerance ? 'weighing-weight-fail-row' : isLocked ? 'locked-row' : isCurrent ? 'current-row' : isDone ? 'passed-row' : ''}>
                        <td>{index + 1}</td>
                        <td>{item.materialCode}</td>
                        <td>{item.materialName}</td>
                        <td>{formatKg(item.requiredKg)}</td>
                        <td>{hasFormulaTolerance(item) ? formatToleranceKg(item.toleranceKg) : 'Chưa có dung sai'}</td>
                        <td>{item.qrScanned || '-'}</td>
                        <td><span className={`flow-pill ${statusClass(item.qrStatus)}`}>{item.qrStatus}</span></td>
                        <td>{item.actualWeight !== '' ? formatKg(item.actualWeight) : '-'}</td>
                        <td><span className={`flow-pill ${statusClass(item.weighStatus)}`}>{item.weighStatus === 'PASS' ? 'Đạt' : item.weighStatus === 'FAIL' ? 'Ngoài dung sai' : missingTolerance ? 'Chưa có dung sai' : item.weighStatus}</span></td>
                        <td>{item.confirmedAt || '-'}</td>
                        <td>{isLocked && !isDone ? 'Khóa' : item.note || '-'}</td>
                      </tr>
                    )
                  })}
                  {solids.length === 0 && (
                    <tr><td colSpan="11" className="empty-row">Lệnh này chưa có dòng nguyên liệu rắn.</td></tr>
                  )}
                </tbody>
              </table>
            </div>
          </section>
        </div>
      </section>
    </div>
  )
}
