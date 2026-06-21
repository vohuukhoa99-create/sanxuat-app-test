import { useMemo, useState } from 'react'
import { customerCatalog } from '../data/customerCatalog.js'

const normalizedCustomerCatalog = (customerCatalog || []).map((customer, index) => ({
  id: customer.id || customer.customerCode || `CUS-${index}`,
  customerCode: customer.customerCode || '',
  customerName: customer.customerName || '',
  province: customer.province || '',
  channelCode: customer.channelCode || '',
  status: customer.status || '',
})).filter((customer) => customer.customerCode || customer.customerName)
const formatCustomerOption = (customer = {}) => [customer.customerCode, customer.customerName, customer.province].filter(Boolean).join(' - ')
const customerMatches = (customer = {}, query = '') => {
  const value = query.trim().toLowerCase()
  if (!value) return true
  return [customer.customerCode, customer.customerName, customer.province, customer.channelCode]
    .some((item) => String(item || '').toLowerCase().includes(value))
}

const hssIngredients = [
  { materialCode: 'PASTE 02', materialName: 'PASTE 02', materialGroup: 'Hóa chế', ratioPercent: 4.61 },
  { materialCode: 'IN03', materialName: 'IN03', materialGroup: 'Hóa chế', ratioPercent: 0.02 },
  { materialCode: 'IN02', materialName: 'IN02', materialGroup: 'Hóa chế', ratioPercent: 0.3 },
  { materialCode: 'D01', materialName: 'D01', materialGroup: 'Hóa chế', ratioPercent: 0.07 },
  { materialCode: 'R91', materialName: 'R91', materialGroup: 'Nguyên liệu rắn', ratioPercent: 20 },
  { materialCode: 'SiG0605', materialName: 'SiG0605', materialGroup: 'Nguyên liệu rắn', ratioPercent: 15 },
  { materialCode: 'SiG0703', materialName: 'SiG0703', materialGroup: 'Nguyên liệu rắn', ratioPercent: 10 },
  { materialCode: 'SW34', materialName: 'SW34', materialGroup: 'Nguyên liệu rắn', ratioPercent: 34.5 },
  { materialCode: 'SW92', materialName: 'SW92', materialGroup: 'Nguyên liệu rắn', ratioPercent: 15 },
  { materialCode: 'KT01', materialName: 'KT01', materialGroup: 'Hóa chế', ratioPercent: 0.5 },
]

const defaultOrder = () => ({
  id: `LSX-${new Date().toISOString().slice(0, 10).replaceAll('-', '')}-${String(Date.now()).slice(-4)}`,
  productionDate: new Date().toISOString().slice(0, 10),
  customer: '',
  customerName: '',
  customerCode: '',
  province: '',
  channelCode: '',
  customerSearch: '',
  product: 'HSS 251.023',
  formula: 'HSS 251.023',
  lot: '',
  quantityKg: 1000,
  batchCount: 1,
  machineInTime: '',
  machineOutTime: '',
  status: 'Chờ cân',
  stage: 'chemical',
  ingredients: hssIngredients.map((item, index) => ({
    id: crypto.randomUUID(),
    no: index + 1,
    ...item,
    materialPerBatch: '',
    materialPerLot: Number(((item.ratioPercent * 1000) / 100).toFixed(3)),
    actual: '',
    adjustment: '',
    weighStatus: 'Chờ cân',
    note: '',
  })),
})

const blankIngredient = (quantityKg, no) => ({
  id: crypto.randomUUID(),
  no,
  materialCode: '',
  materialName: '',
  materialGroup: 'Hóa chế',
  ratioPercent: 0,
  materialPerBatch: '',
  materialPerLot: 0,
  actual: '',
  adjustment: '',
  weighStatus: 'Chờ cân',
  note: '',
})

const calcMaterialPerLot = (ratioPercent, quantityKg) =>
  Number((((Number(ratioPercent) || 0) * (Number(quantityKg) || 0)) / 100).toFixed(3))

function CustomerSearchCombobox({ value = {}, inputValue = '', onInputChange, onSelect }) {
  const [open, setOpen] = useState(false)
  const filteredCustomers = useMemo(() => normalizedCustomerCatalog.filter((customer) => customerMatches(customer, inputValue)).slice(0, 30), [inputValue])
  return (
    <div className="customer-combobox">
      <input
        value={inputValue}
        placeholder="Gõ mã, tên, tỉnh/thành hoặc kênh"
        onFocus={() => setOpen(true)}
        onChange={(event) => {
          onInputChange(event.target.value)
          setOpen(true)
        }}
        onBlur={() => window.setTimeout(() => setOpen(false), 120)}
        aria-autocomplete="list"
        aria-expanded={open}
      />
      {value?.customerCode && <span className="customer-combobox-selected">{value.customerCode} / {value.customerName}</span>}
      {open && (
        <div className="customer-combobox-menu">
          {filteredCustomers.length ? filteredCustomers.map((customer) => (
            <button type="button" key={customer.customerCode || customer.id} onMouseDown={(event) => event.preventDefault()} onClick={() => { onSelect(customer); setOpen(false) }}>
              <strong>{customer.customerCode}</strong>
              <span>{customer.customerName}</span>
              <em>{customer.province || '-'}</em>
            </button>
          )) : <div className="customer-combobox-empty">Không có khách hàng phù hợp</div>}
        </div>
      )}
    </div>
  )
}

function OrderCreateModal({ onClose, onSave }) {
  const [form, setForm] = useState(defaultOrder)
  const selectedCustomer = normalizedCustomerCatalog.find((customer) => customer.customerCode === form.customerCode)

  const updateField = (field, value) => {
    setForm((current) => {
      const next = { ...current, [field]: value }
      if (field === 'quantityKg') {
        const quantityKg = Number(value) || 0
        next.ingredients = current.ingredients.map((item) => ({
          ...item,
          materialPerLot: calcMaterialPerLot(item.ratioPercent, quantityKg),
        }))
      }
      return next
    })
  }

  const updateIngredient = (id, field, value) => {
    setForm((current) => ({
      ...current,
      ingredients: current.ingredients.map((item) => {
        if (item.id !== id) return item
        const next = { ...item, [field]: value }
        if (field === 'ratioPercent') {
          next.materialPerLot = calcMaterialPerLot(value, current.quantityKg)
        }
        return next
      }),
    }))
  }

  const addIngredient = () => {
    setForm((current) => ({
      ...current,
      ingredients: [
        ...current.ingredients,
        blankIngredient(current.quantityKg, current.ingredients.length + 1),
      ],
    }))
  }

  const removeIngredient = (id) => {
    setForm((current) => ({
      ...current,
      ingredients: current.ingredients
        .filter((item) => item.id !== id)
        .map((item, index) => ({ ...item, no: index + 1 })),
    }))
  }

  const handleSubmit = (event) => {
    event.preventDefault()
    if (!selectedCustomer) {
      window.alert('Vui lòng chọn khách hàng hợp lệ từ danh mục.')
      return
    }
    const now = new Date().toISOString().slice(0, 16).replace('T', ' ')
    onSave({
      ...form,
      customer: selectedCustomer.customerName,
      customerName: selectedCustomer.customerName,
      customerCode: selectedCustomer.customerCode,
      province: selectedCustomer.province || '',
      channelCode: selectedCustomer.channelCode || '',
      quantityKg: Number(form.quantityKg) || 0,
      batchCount: Number(form.batchCount) || 0,
      createdAt: now,
      updatedAt: now,
      status: form.status || 'Chờ cân',
      stage: form.stage || 'chemical',
      ingredients: form.ingredients.map((item, index) => ({
        ...item,
        no: index + 1,
        ratioPercent: Number(item.ratioPercent) || 0,
        materialPerLot: calcMaterialPerLot(item.ratioPercent, form.quantityKg),
        weighStatus: item.weighStatus || 'Chờ cân',
      })),
    })
  }

  return (
    <div className="modal-backdrop" role="presentation">
      <div className="production-modal" role="dialog" aria-modal="true" aria-labelledby="create-order-title">
        <form onSubmit={handleSubmit}>
          <div className="modal-header">
            <div>
              <span className="section-kicker">Biểu mẫu sản xuất</span>
              <h2 id="create-order-title">Tạo lệnh sản xuất</h2>
            </div>
            <button type="button" className="icon-button" onClick={onClose} aria-label="Đóng">×</button>
          </div>

          <section className="form-section">
            <h3>A. Thông tin chung</h3>
            <div className="production-form-grid">
              <label>Mã lệnh sản xuất<input value={form.id} onChange={(event) => updateField('id', event.target.value)} required /></label>
              <label>Ngày sản xuất<input type="date" value={form.productionDate} onChange={(event) => updateField('productionDate', event.target.value)} /></label>
              <label>Khách hàng / CT-KH
                <CustomerSearchCombobox
                  value={selectedCustomer}
                  inputValue={form.customerSearch}
                  onInputChange={(value) => setForm((current) => ({ ...current, customerSearch: value, customer: '', customerName: '', customerCode: '', province: '', channelCode: '' }))}
                  onSelect={(customer) => setForm((current) => ({ ...current, customerSearch: formatCustomerOption(customer), customer: customer.customerName, customerName: customer.customerName, customerCode: customer.customerCode, province: customer.province || '', channelCode: customer.channelCode || '' }))}
                />
              </label>
              <label>Tên sản phẩm<input value={form.product} onChange={(event) => updateField('product', event.target.value)} required /></label>
              <label>Mã sản phẩm / công thức<input value={form.formula} onChange={(event) => updateField('formula', event.target.value)} required /></label>
              <label>Mã lô / LOT<input value={form.lot} onChange={(event) => updateField('lot', event.target.value)} /></label>
              <label>Khối lượng yêu cầu<input type="number" min="0" step="0.001" value={form.quantityKg} onChange={(event) => updateField('quantityKg', event.target.value)} required /></label>
              <label>Tổng số mẻ sản xuất<input type="number" min="0" step="1" value={form.batchCount} onChange={(event) => updateField('batchCount', event.target.value)} /></label>
              <label>Thời gian vào máy<input type="datetime-local" value={form.machineInTime} onChange={(event) => updateField('machineInTime', event.target.value)} /></label>
              <label>Thời gian ra máy<input type="datetime-local" value={form.machineOutTime} onChange={(event) => updateField('machineOutTime', event.target.value)} /></label>
              <label>Trạng thái lệnh<select value={form.status} onChange={(event) => updateField('status', event.target.value)}><option>Chờ cân</option><option>Đang cân</option><option>Hoàn thành</option><option>Tạm dừng</option></select></label>
            </div>
          </section>

          <section className="form-section">
            <div className="section-heading-row">
              <h3>B. Bảng nguyên liệu / đơn phối trộn</h3>
              <button type="button" className="secondary-button" onClick={addIngredient}>+ Thêm dòng</button>
            </div>
            <div className="table-wrapper recipe-table-wrapper">
              <table className="recipe-table">
                <thead>
                  <tr>
                    <th>STT</th>
                    <th>Mã vật tư</th>
                    <th>Tên nguyên liệu</th>
                    <th>Nhóm nguyên liệu</th>
                    <th>Tỷ lệ %</th>
                    <th>Vật tư / mẻ</th>
                    <th>Vật tư / lô</th>
                    <th>Thực tế</th>
                    <th>Điều chỉnh</th>
                    <th>Trạng thái cân</th>
                    <th>Ghi chú</th>
                    <th></th>
                  </tr>
                </thead>
                <tbody>
                  {form.ingredients.map((item, index) => (
                    <tr key={item.id}>
                      <td>{index + 1}</td>
                      <td><input value={item.materialCode} onChange={(event) => updateIngredient(item.id, 'materialCode', event.target.value)} /></td>
                      <td><input value={item.materialName} onChange={(event) => updateIngredient(item.id, 'materialName', event.target.value)} /></td>
                      <td><select value={item.materialGroup} onChange={(event) => updateIngredient(item.id, 'materialGroup', event.target.value)}><option>Hóa chế</option><option>Nguyên liệu rắn</option></select></td>
                      <td><input type="number" min="0" step="0.001" value={item.ratioPercent} onChange={(event) => updateIngredient(item.id, 'ratioPercent', event.target.value)} /></td>
                      <td><input value={item.materialPerBatch} onChange={(event) => updateIngredient(item.id, 'materialPerBatch', event.target.value)} /></td>
                      <td><input type="number" readOnly value={item.materialPerLot} /></td>
                      <td><input value={item.actual} onChange={(event) => updateIngredient(item.id, 'actual', event.target.value)} /></td>
                      <td><input value={item.adjustment} onChange={(event) => updateIngredient(item.id, 'adjustment', event.target.value)} /></td>
                      <td><select value={item.weighStatus} onChange={(event) => updateIngredient(item.id, 'weighStatus', event.target.value)}><option>Chờ cân</option><option>PASS</option><option>FAIL</option></select></td>
                      <td><input value={item.note} onChange={(event) => updateIngredient(item.id, 'note', event.target.value)} /></td>
                      <td><button type="button" className="danger-button" onClick={() => removeIngredient(item.id)}>Xóa</button></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </section>

          <div className="modal-actions">
            <button type="button" className="secondary-button" onClick={onClose}>Hủy</button>
            <button type="submit" className="primary-button">Lưu lệnh</button>
          </div>
        </form>
      </div>
    </div>
  )
}

export function ProductionOrdersPage({ orders, onCreateOrder }) {
  const [isCreating, setIsCreating] = useState(false)
  const sortedOrders = useMemo(() => [...orders].reverse(), [orders])

  const handleSave = (order) => {
    onCreateOrder(order)
    setIsCreating(false)
  }

  return (
    <div className="page-content">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Danh sách lệnh sản xuất</h2>
            <p className="panel-text">Theo dõi lệnh, công thức, khối lượng và trạng thái cân ban đầu.</p>
          </div>
          <button className="primary-button" onClick={() => setIsCreating(true)}>+ Tạo lệnh sản xuất</button>
        </div>
        <div className="table-wrapper">
          <table>
            <thead>
              <tr>
                <th>Lệnh</th>
                <th>Sản phẩm</th>
                <th>Khách hàng</th>
                <th>Công thức</th>
                <th>LOT</th>
                <th>Trạng thái</th>
                <th>Khối lượng</th>
                <th>Nguyên liệu</th>
                <th>Cập nhật</th>
              </tr>
            </thead>
            <tbody>
              {sortedOrders.map((order) => (
                <tr key={order.id}>
                  <td>{order.id}</td>
                  <td>{order.product}</td>
                  <td>{order.customerName || order.customer || '-'}</td>
                  <td>{order.formula}</td>
                  <td>{order.lot || '-'}</td>
                  <td><span className="status-pill">{order.status || order.stage}</span></td>
                  <td>{order.quantityKg} kg</td>
                  <td>{order.ingredients?.length || '-'}</td>
                  <td>{order.updatedAt}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
      {isCreating && <OrderCreateModal onClose={() => setIsCreating(false)} onSave={handleSave} />}
    </div>
  )
}
