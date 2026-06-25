import { useMemo, useState } from 'react'

const QR_WAITING = 'Chờ quét'
const WEIGH_WAITING = 'Chờ cân'

const formatKg = (value) => `${Number(value || 0).toLocaleString('vi-VN', { minimumFractionDigits: 3, maximumFractionDigits: 3 })} kg`
const formatToleranceKg = (value) => value === '' || value === null || value === undefined ? '-' : `±${formatKg(value)}`
const hasFormulaTolerance = (item = {}) => item.toleranceKg !== '' && item.toleranceKg !== null && item.toleranceKg !== undefined && Number.isFinite(Number(item.toleranceKg))
const toleranceErrorMessage = (requiredKg, actualWeight, toleranceKg) => `Khối lượng ngoài dung sai. Cần cân: ${formatKg(requiredKg)}, Dung sai: ${formatToleranceKg(toleranceKg)}`

function isChemicalGroup(group = '') {
  const value = group.toLowerCase()
  return value.includes('hóa chế')
    || value.includes('hóa chất')
    || value.includes('hoa che')
    || value.includes('hoa chat')
    || value.includes('chemical')
    || value.includes('hóa')
}

function getChemicalIngredients(order) {
  return (order.ingredients || [])
    .filter((item) => isChemicalGroup(item.materialGroup))
    .map((item, index) => normalizeChemicalItem(item, index))
}

function normalizeChemicalItem(item, index) {
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

function getChemicalStatus(order) {
  const chemicals = getChemicalIngredients(order)
  if (order.scaleStatus?.chemical) return order.scaleStatus.chemical
  if (chemicals.length === 0) return 'Pending'
  if (chemicals.every((item) => item.qrStatus === 'PASS' && item.weighStatus === 'PASS')) return 'Completed'
  return order.stage === 'chemical' ? 'Ready' : 'Pending'
}

function getFirstOpenIndex(chemicals) {
  return chemicals.findIndex((item) => !(item.qrStatus === 'PASS' && item.weighStatus === 'PASS'))
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

export function ChemicalScalePage({ orders, onUpdateOrder }) {
  const chemicalOrders = useMemo(
    () => orders
      .filter((order) => getChemicalIngredients(order).length > 0 || order.stage === 'chemical')
      .map((order) => ({ ...order, chemicalStatus: getChemicalStatus(order) })),
    [orders],
  )
  const storedActiveOrder = chemicalOrders.find((order) => order.chemicalStatus === 'Active')
  const [activeOrderId, setActiveOrderId] = useState(storedActiveOrder?.id || chemicalOrders[0]?.id || '')
  const activeOrder = chemicalOrders.find((order) => order.id === activeOrderId)
    || storedActiveOrder
    || chemicalOrders[0]
  const chemicals = activeOrder ? getChemicalIngredients(activeOrder) : []
  const firstOpenIndex = getFirstOpenIndex(chemicals)
  const currentIndex = firstOpenIndex === -1 ? chemicals.length : firstOpenIndex
  const currentItem = chemicals[currentIndex]
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
        scaleStatus: { ...(current.scaleStatus || {}), chemical: 'Active' },
        status: 'Active',
        chemicalStartedAt: current.chemicalStartedAt || new Date().toISOString().slice(0, 16).replace('T', ' '),
      }),
      `Lệnh ${order.id} bắt đầu cân hóa.`,
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
        `QR hóa FAIL ${activeOrder.id} - ${currentItem.materialCode}.`,
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
      `QR hóa PASS ${activeOrder.id} - ${currentItem.materialCode}.`,
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
      const nextChemicals = getChemicalIngredients(nextOrder)
      const isCompleted = nextChemicals.every((item) => item.qrStatus === 'PASS' && item.weighStatus === 'PASS')
      if (!isCompleted) return nextOrder

      const solidCompleted = nextOrder.scaleStatus?.solid === 'Completed'
      return {
        ...nextOrder,
        stage: solidCompleted ? 'mixing' : 'solid',
        status: solidCompleted ? 'Ready phối trộn' : 'Chờ cân rắn',
        scaleStatus: {
          ...(nextOrder.scaleStatus || {}),
          chemical: 'Completed',
        },
        chemicalCompletedAt: updatedAt,
      }
    }, `Lệnh ${activeOrder.id} xác nhận cân hóa ${currentItem.materialCode}.`)
  }

  return (
    <div className="page-content chemical-workflow">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Tổ cân hóa</h2>
            <p className="panel-text">Quy trình cân tuần tự: quét QR, xác nhận đúng mã vật tư, cân khối lượng, rồi PASS/FAIL.</p>
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
                  {chemicalOrders.map((order) => (
                    <tr key={order.id} className={activeOrder?.id === order.id ? 'active-order-row' : ''}>
                      <td>{order.id}</td>
                      <td>{order.product}</td>
                      <td>{order.quantityKg} kg</td>
                      <td><span className={`flow-pill ${statusClass(order.chemicalStatus)}`}>{order.chemicalStatus}</span></td>
                      <td>
                        <button
                          className="primary-button"
                          disabled={order.chemicalStatus === 'Completed'}
                          onClick={() => startOrder(order)}
                        >
                          Bắt đầu cân
                        </button>
                      </td>
                    </tr>
                  ))}
                  {chemicalOrders.length === 0 && (
                    <tr><td colSpan="5" className="empty-row">Không có lệnh cần cân hóa.</td></tr>
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
              {activeOrder?.chemicalCompletedAt && (
                <span className="flow-pill pass">Hoàn thành: {activeOrder.chemicalCompletedAt}</span>
              )}
            </div>

            <div className="workflow-steps">
              <span className="step active">Bước 1: Quét QR</span>
              <span className="step active">Bước 2: Xác nhận đúng mã vật tư</span>
              <span className={currentItem?.qrStatus === 'PASS' ? 'step active' : 'step locked'}>Bước 3: Cân khối lượng</span>
              <span className="step">Bước 4: PASS/FAIL</span>
            </div>

            {warning && <div className="process-alert">{warning}</div>}

            {currentItem && activeOrder?.chemicalStatus === 'Active' && (
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
                    <th>Tên hóa chất</th>
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
                  {chemicals.map((item, index) => {
                    const isDone = item.qrStatus === 'PASS' && item.weighStatus === 'PASS'
                    const missingTolerance = item.qrStatus === 'PASS' && !hasFormulaTolerance(item)
                    const isCurrent = activeOrder?.chemicalStatus === 'Active' && index === currentIndex
                    const isLocked = activeOrder?.chemicalStatus !== 'Active' || index > currentIndex
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
                  {chemicals.length === 0 && (
                    <tr><td colSpan="11" className="empty-row">Lệnh này chưa có dòng hóa chất.</td></tr>
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
