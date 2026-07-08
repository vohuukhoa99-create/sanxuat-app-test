import { useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'

const SETTINGS_KEY = 'sonhoabinh-system-settings-v1'
const PRODUCTION_KEY = 'sonhoabinh-production-v1'
const STORAGE_QUOTA_MESSAGE = 'Bộ nhớ trình duyệt đã đầy. Vui lòng Reset dữ liệu demo hoặc chuyển sang database.'

const tabs = [
  ['materials', 'Nguyên liệu'],
  ['formulas', 'Công thức'],
  ['tolerances', 'Dung sai cân'],
  ['devices', 'Máy móc'],
  ['users', 'Người dùng'],
  ['qr', 'Thiết lập QR'],
  ['production', 'Sản xuất'],
  ['backup', 'Sao lưu'],
  ['factory', 'Thông số nhà máy'],
]

const materialGroups = ['Hóa chế', 'Nguyên liệu rắn', 'Bao bì', 'Thành phẩm']
const userRoles = ['Admin', 'Kế hoạch sản xuất', 'Tổ cân hóa', 'Tổ cân rắn', 'Tổ phối trộn', 'QC', 'Quản đốc', 'Ban giám đốc']

const defaultSettings = {
  materials: [
    { id: 'mat-paste-02', code: 'PASTE 02', name: 'PASTE 02', group: 'Hóa chế', unit: 'kg', standardWeight: 2.31, defaultTolerance: 0.01, supplier: 'NCC Hóa chất A', status: 'Hoạt động', note: 'Công thức HSS 251.023' },
    { id: 'mat-in03', code: 'IN03', name: 'IN03', group: 'Hóa chế', unit: 'kg', standardWeight: 0.01, defaultTolerance: 0.005, supplier: 'NCC Hóa chất A', status: 'Hoạt động', note: '' },
    { id: 'mat-r91', code: 'R91', name: 'R91', group: 'Nguyên liệu rắn', unit: 'kg', standardWeight: 10, defaultTolerance: 0.1, supplier: 'NCC Bột khoáng B', status: 'Hoạt động', note: '' },
    { id: 'mat-sw34', code: 'SW34', name: 'SW34', group: 'Nguyên liệu rắn', unit: 'kg', standardWeight: 17.25, defaultTolerance: 0.1, supplier: 'NCC Bột khoáng B', status: 'Hoạt động', note: '' },
  ],
  formulas: [
    {
      id: 'formula-hss-251023',
      code: 'HSS 251.023',
      productName: 'HSS 251.023',
      productGroup: 'Sơn công nghiệp',
      batchWeight: 50,
      details: [
        { id: 'fd-paste-02', materialCode: 'PASTE 02', materialName: 'PASTE 02', ratio: 4.61, weightPerBatch: 2.31, group: 'Hóa chế' },
        { id: 'fd-in03', materialCode: 'IN03', materialName: 'IN03', ratio: 0.02, weightPerBatch: 0.01, group: 'Hóa chế' },
        { id: 'fd-in02', materialCode: 'IN02', materialName: 'IN02', ratio: 0.3, weightPerBatch: 0.15, group: 'Hóa chế' },
        { id: 'fd-d01', materialCode: 'D01', materialName: 'D01', ratio: 0.07, weightPerBatch: 0.035, group: 'Hóa chế' },
        { id: 'fd-r91', materialCode: 'R91', materialName: 'R91', ratio: 20, weightPerBatch: 10, group: 'Nguyên liệu rắn' },
        { id: 'fd-sig0605', materialCode: 'SiG0605', materialName: 'SiG0605', ratio: 15, weightPerBatch: 7.5, group: 'Nguyên liệu rắn' },
        { id: 'fd-sig0703', materialCode: 'SiG0703', materialName: 'SiG0703', ratio: 10, weightPerBatch: 5, group: 'Nguyên liệu rắn' },
        { id: 'fd-sw34', materialCode: 'SW34', materialName: 'SW34', ratio: 34.5, weightPerBatch: 17.25, group: 'Nguyên liệu rắn' },
        { id: 'fd-sw92', materialCode: 'SW92', materialName: 'SW92', ratio: 15, weightPerBatch: 7.5, group: 'Nguyên liệu rắn' },
        { id: 'fd-kt01', materialCode: 'KT01', materialName: 'KT01', ratio: 0.5, weightPerBatch: 0.25, group: 'Hóa chế' },
      ],
    },
  ],
  tolerances: [
    { id: 'tol-chem-1', group: 'Hóa chất', min: 0, max: 1, tolerance: 0.005 },
    { id: 'tol-chem-2', group: 'Hóa chất', min: 1, max: 10, tolerance: 0.01 },
    { id: 'tol-chem-3', group: 'Hóa chất', min: 10, max: 100, tolerance: 0.05 },
    { id: 'tol-solid-1', group: 'Nguyên liệu rắn', min: 0, max: 10, tolerance: 0.05 },
    { id: 'tol-solid-2', group: 'Nguyên liệu rắn', min: 10, max: 50, tolerance: 0.1 },
    { id: 'tol-solid-3', group: 'Nguyên liệu rắn', min: 50, max: 1000, tolerance: 0.5 },
  ],
  devices: [
    { id: 'dev-ch-01', code: 'CH-01', name: 'Cân hóa số 1', type: 'Cân hóa', status: 'Hoạt động', note: '' },
    { id: 'dev-ch-02', code: 'CH-02', name: 'Cân hóa số 2', type: 'Cân hóa', status: 'Hoạt động', note: '' },
    { id: 'dev-cr-01', code: 'CR-01', name: 'Cân rắn số 1', type: 'Cân rắn', status: 'Hoạt động', note: '' },
    { id: 'dev-cr-02', code: 'CR-02', name: 'Cân rắn số 2', type: 'Cân rắn', status: 'Bảo trì', note: '' },
    { id: 'dev-mt-01', code: 'MT-01', name: 'Máy trộn số 1', type: 'Máy trộn', status: 'Hoạt động', note: '' },
    { id: 'dev-mt-02', code: 'MT-02', name: 'Máy trộn số 2', type: 'Máy trộn', status: 'Hoạt động', note: '' },
  ],
  users: [
    { id: 'user-admin', username: 'admin', fullName: 'Quản trị hệ thống', department: 'CNTT', role: 'Admin', status: 'Hoạt động' },
    { id: 'user-chem', username: 'canhoa01', fullName: 'Nhân viên cân hóa 01', department: 'Tổ cân hóa', role: 'Tổ cân hóa', status: 'Hoạt động' },
    { id: 'user-solid', username: 'canran01', fullName: 'Nhân viên cân rắn 01', department: 'Tổ cân rắn', role: 'Tổ cân rắn', status: 'Hoạt động' },
    { id: 'user-qc', username: 'qc01', fullName: 'QC ca 1', department: 'QC', role: 'QC', status: 'Hoạt động' },
  ],
  qrRules: { materialCode: true, materialName: true, lot: true, importDate: true, preview: '' },
  productionSettings: {
    chemicalScaleCount: 2,
    solidScaleCount: 2,
    warningMinutes: 30,
    rules: [
      'Nếu cân hóa Completed và cân rắn Completed => Ready phối trộn',
      'Nếu phối trộn Completed => Chuyển QC',
      'Nếu QC Passed => Hoàn thành',
    ],
  },
  factory: {
    companyName: 'Công ty Sơn Hòa Bình',
    factoryName: 'Nhà máy Sơn Hòa Bình',
    address: 'Khu công nghiệp Hòa Bình',
    logo: '',
    phone: '028 0000 0000',
    email: 'info@sonhoabinh.vn',
  },
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY)
    if (!raw) return defaultSettings
    return { ...defaultSettings, ...JSON.parse(raw) }
  } catch {
    return defaultSettings
  }
}

function isQuotaExceededError(error) {
  return error?.name === 'QuotaExceededError'
    || error?.name === 'NS_ERROR_DOM_QUOTA_REACHED'
    || error?.code === 22
    || error?.code === 1014
}

function safeSetJsonLocalStorage(key, data) {
  try {
    localStorage.setItem(key, JSON.stringify(data))
    return true
  } catch (error) {
    if (isQuotaExceededError(error)) window.alert(STORAGE_QUOTA_MESSAGE)
    else console.warn(`Không thể lưu localStorage key ${key}.`, error)
    return false
  }
}

function saveSettings(settings) {
  safeSetJsonLocalStorage(SETTINGS_KEY, settings)
}

function newId(prefix) {
  return `${prefix}-${crypto.randomUUID()}`
}

function downloadJson(filename, data) {
  const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  link.remove()
  URL.revokeObjectURL(url)
}

function PseudoQr({ value }) {
  const bits = Array.from({ length: 81 }, (_, index) => {
    const code = value.charCodeAt(index % Math.max(value.length, 1)) || 0
    return (code + index * 7) % 3 === 0
  })
  return <div className="pseudo-qr">{bits.map((bit, index) => <span key={index} className={bit ? 'on' : ''} />)}</div>
}

export function SettingsPage() {
  const [activeTab, setActiveTab] = useState('materials')
  const [settings, setSettings] = useState(loadSettings)
  const materialImportRef = useRef(null)
  const backupImportRef = useRef(null)

  useEffect(() => {
    saveSettings(settings)
  }, [settings])

  const updateSettings = (updater) => {
    setSettings((current) => (typeof updater === 'function' ? updater(current) : updater))
  }

  return (
    <div className="page-content settings-page">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Cài đặt hệ thống</h2>
            <p className="panel-text">Quản trị dữ liệu nền cho toàn bộ quy trình sản xuất Sơn Hòa Bình.</p>
          </div>
        </div>
        <div className="settings-tabs">
          {tabs.map(([id, label]) => (
            <button key={id} className={activeTab === id ? 'active' : ''} onClick={() => setActiveTab(id)}>{label}</button>
          ))}
        </div>
      </section>

      {activeTab === 'materials' && (
        <MaterialsTab
          materials={settings.materials}
          onChange={(materials) => updateSettings((current) => ({ ...current, materials }))}
          importRef={materialImportRef}
        />
      )}
      {activeTab === 'formulas' && (
        <FormulasTab
          formulas={settings.formulas}
          materials={settings.materials}
          onChange={(formulas) => updateSettings((current) => ({ ...current, formulas }))}
        />
      )}
      {activeTab === 'tolerances' && <EditableTableTab title="Dung sai cân" rows={settings.tolerances} fields={['group', 'min', 'max', 'tolerance']} labels={['Nhóm', 'Từ kg', 'Đến kg', 'Dung sai kg']} onChange={(rows) => updateSettings((current) => ({ ...current, tolerances: rows }))} />}
      {activeTab === 'devices' && <EditableTableTab title="Máy móc thiết bị" rows={settings.devices} fields={['code', 'name', 'type', 'status', 'note']} labels={['Mã thiết bị', 'Tên thiết bị', 'Loại', 'Trạng thái', 'Ghi chú']} onChange={(rows) => updateSettings((current) => ({ ...current, devices: rows }))} />}
      {activeTab === 'users' && <EditableTableTab title="Người dùng" rows={settings.users} fields={['username', 'fullName', 'department', 'role', 'status']} labels={['Tài khoản', 'Họ tên', 'Bộ phận', 'Vai trò', 'Trạng thái']} roleOptions={userRoles} onChange={(rows) => updateSettings((current) => ({ ...current, users: rows }))} />}
      {activeTab === 'qr' && <QrTab qrRules={settings.qrRules} materials={settings.materials} onChange={(qrRules) => updateSettings((current) => ({ ...current, qrRules }))} />}
      {activeTab === 'production' && <ProductionSettingsTab settings={settings.productionSettings} onChange={(productionSettings) => updateSettings((current) => ({ ...current, productionSettings }))} />}
      {activeTab === 'backup' && <BackupTab settings={settings} onRestore={setSettings} importRef={backupImportRef} />}
      {activeTab === 'factory' && <FactoryTab factory={settings.factory} onChange={(factory) => updateSettings((current) => ({ ...current, factory }))} />}
    </div>
  )
}

function MaterialsTab({ materials, onChange, importRef }) {
  const [search, setSearch] = useState('')
  const [editingId, setEditingId] = useState('')
  const filtered = useMemo(() => materials.filter((item) =>
    `${item.code} ${item.name} ${item.group} ${item.supplier}`.toLowerCase().includes(search.toLowerCase()),
  ), [materials, search])

  const addRow = () => {
    const row = { id: newId('mat'), code: '', name: '', group: 'Hóa chế', unit: 'kg', standardWeight: 0, defaultTolerance: 0, supplier: '', status: 'Hoạt động', note: '' }
    onChange([row, ...materials])
    setEditingId(row.id)
  }
  const updateRow = (id, field, value) => onChange(materials.map((item) => item.id === id ? { ...item, [field]: value } : item))
  const deleteRow = (id) => onChange(materials.filter((item) => item.id !== id))
  const exportExcel = () => {
    const sheet = XLSX.utils.json_to_sheet(materials)
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, sheet, 'Materials')
    XLSX.writeFile(book, 'danh-muc-nguyen-lieu.xlsx')
  }
  const importExcel = (file) => {
    if (!file) return
    const reader = new FileReader()
    reader.onload = (event) => {
      const workbook = XLSX.read(event.target.result, { type: 'array' })
      const sheet = workbook.Sheets[workbook.SheetNames[0]]
      const rows = XLSX.utils.sheet_to_json(sheet)
      const imported = rows.map((row) => ({
        id: newId('mat'),
        code: row.code || row['Mã vật tư'] || row['Ma vat tu'] || '',
        name: row.name || row['Tên vật tư'] || row['Ten vat tu'] || '',
        group: row.group || row['Nhóm vật tư'] || 'Hóa chế',
        unit: row.unit || row['Đơn vị tính'] || 'kg',
        standardWeight: Number(row.standardWeight || row['Khối lượng chuẩn'] || 0),
        defaultTolerance: Number(row.defaultTolerance || row['Dung sai mặc định'] || 0),
        supplier: row.supplier || row['Nhà cung cấp'] || '',
        status: row.status || row['Trạng thái'] || 'Hoạt động',
        note: row.note || row['Ghi chú'] || '',
      }))
      onChange([...imported, ...materials])
    }
    reader.readAsArrayBuffer(file)
  }

  return (
    <section className="panel settings-section">
      <div className="section-heading-row">
        <h3>Danh mục nguyên liệu</h3>
        <div className="action-row">
          <input className="settings-search" placeholder="Tìm kiếm vật tư" value={search} onChange={(event) => setSearch(event.target.value)} />
          <button className="primary-button" onClick={addRow}>Thêm</button>
          <button className="secondary-button" onClick={() => importRef.current?.click()}>Import Excel</button>
          <button className="secondary-button" onClick={exportExcel}>Export Excel</button>
          <input ref={importRef} type="file" accept=".xlsx,.xls" hidden onChange={(event) => importExcel(event.target.files?.[0])} />
        </div>
      </div>
      <div className="table-wrapper">
        <table className="settings-wide-table">
          <thead><tr><th>Mã vật tư</th><th>Tên vật tư</th><th>Nhóm vật tư</th><th>Đơn vị tính</th><th>Khối lượng chuẩn</th><th>Dung sai mặc định</th><th>Nhà cung cấp</th><th>Trạng thái</th><th>Ghi chú</th><th>Hành động</th></tr></thead>
          <tbody>
            {filtered.map((item) => {
              const editing = editingId === item.id
              return (
                <tr key={item.id}>
                  {editing ? (
                    <>
                      <td><input value={item.code} onChange={(event) => updateRow(item.id, 'code', event.target.value)} /></td>
                      <td><input value={item.name} onChange={(event) => updateRow(item.id, 'name', event.target.value)} /></td>
                      <td><select value={item.group} onChange={(event) => updateRow(item.id, 'group', event.target.value)}>{materialGroups.map((group) => <option key={group}>{group}</option>)}</select></td>
                      <td><input value={item.unit} onChange={(event) => updateRow(item.id, 'unit', event.target.value)} /></td>
                      <td><input type="number" value={item.standardWeight} onChange={(event) => updateRow(item.id, 'standardWeight', Number(event.target.value))} /></td>
                      <td><input type="number" step="0.001" value={item.defaultTolerance} onChange={(event) => updateRow(item.id, 'defaultTolerance', Number(event.target.value))} /></td>
                      <td><input value={item.supplier} onChange={(event) => updateRow(item.id, 'supplier', event.target.value)} /></td>
                      <td><input value={item.status} onChange={(event) => updateRow(item.id, 'status', event.target.value)} /></td>
                      <td><input value={item.note} onChange={(event) => updateRow(item.id, 'note', event.target.value)} /></td>
                    </>
                  ) : (
                    <><td>{item.code}</td><td>{item.name}</td><td>{item.group}</td><td>{item.unit}</td><td>{item.standardWeight}</td><td>{item.defaultTolerance}</td><td>{item.supplier}</td><td>{item.status}</td><td>{item.note}</td></>
                  )}
                  <td className="action-row">
                    <button className="secondary-button" onClick={() => setEditingId(editing ? '' : item.id)}>{editing ? 'Lưu' : 'Sửa'}</button>
                    <button className="danger-button" onClick={() => deleteRow(item.id)}>Xóa</button>
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function FormulasTab({ formulas, materials, onChange }) {
  const [selectedId, setSelectedId] = useState(formulas[0]?.id || '')
  const selected = formulas.find((formula) => formula.id === selectedId) || formulas[0]
  const totalRatio = selected ? selected.details.reduce((sum, row) => sum + Number(row.ratio || 0), 0) : 0
  const addFormula = () => {
    const formula = { id: newId('formula'), code: '', productName: '', productGroup: '', batchWeight: 0, details: [] }
    onChange([formula, ...formulas])
    setSelectedId(formula.id)
  }
  const updateFormula = (field, value) => onChange(formulas.map((formula) => formula.id === selected.id ? { ...formula, [field]: value } : formula))
  const updateDetail = (detailId, field, value) => onChange(formulas.map((formula) => {
    if (formula.id !== selected.id) return formula
    return { ...formula, details: formula.details.map((row) => row.id === detailId ? { ...row, [field]: value } : row) }
  }))
  const addDetail = () => {
    const material = materials[0]
    const detail = { id: newId('fd'), materialCode: material?.code || '', materialName: material?.name || '', ratio: 0, weightPerBatch: 0, group: material?.group || 'Hóa chế' }
    onChange(formulas.map((formula) => formula.id === selected.id ? { ...formula, details: [...formula.details, detail] } : formula))
  }

  if (!selected) return <section className="panel"><button className="primary-button" onClick={addFormula}>Thêm công thức</button></section>

  return (
    <section className="panel settings-section">
      <div className="section-heading-row">
        <h3>Danh mục công thức</h3>
        <div className="action-row">
          <select value={selected.id} onChange={(event) => setSelectedId(event.target.value)}>{formulas.map((formula) => <option key={formula.id} value={formula.id}>{formula.code || formula.productName || 'Công thức mới'}</option>)}</select>
          <button className="primary-button" onClick={addFormula}>Thêm công thức</button>
        </div>
      </div>
      <div className="settings-form-grid">
        <label>Mã công thức<input value={selected.code} onChange={(event) => updateFormula('code', event.target.value)} /></label>
        <label>Tên sản phẩm<input value={selected.productName} onChange={(event) => updateFormula('productName', event.target.value)} /></label>
        <label>Nhóm sản phẩm<input value={selected.productGroup} onChange={(event) => updateFormula('productGroup', event.target.value)} /></label>
        <label>Khối lượng chuẩn/mẻ<input type="number" value={selected.batchWeight} onChange={(event) => updateFormula('batchWeight', Number(event.target.value))} /></label>
      </div>
      <div className={Math.abs(totalRatio - 100) < 0.001 ? 'formula-ratio-ok' : 'formula-ratio-alert'}>Tổng tỷ lệ: {totalRatio.toFixed(3)}%</div>
      <div className="section-heading-row"><h3>Chi tiết công thức</h3><button className="secondary-button" onClick={addDetail}>Thêm dòng</button></div>
      <div className="table-wrapper">
        <table className="settings-wide-table">
          <thead><tr><th>Mã vật tư</th><th>Tên vật tư</th><th>Tỷ lệ %</th><th>Khối lượng/mẻ</th><th>Nhóm vật tư</th><th></th></tr></thead>
          <tbody>
            {selected.details.map((row) => (
              <tr key={row.id}>
                <td><input value={row.materialCode} onChange={(event) => updateDetail(row.id, 'materialCode', event.target.value)} /></td>
                <td><input value={row.materialName} onChange={(event) => updateDetail(row.id, 'materialName', event.target.value)} /></td>
                <td><input type="number" step="0.001" value={row.ratio} onChange={(event) => updateDetail(row.id, 'ratio', Number(event.target.value))} /></td>
                <td><input type="number" step="0.001" value={row.weightPerBatch} onChange={(event) => updateDetail(row.id, 'weightPerBatch', Number(event.target.value))} /></td>
                <td><select value={row.group} onChange={(event) => updateDetail(row.id, 'group', event.target.value)}>{materialGroups.map((group) => <option key={group}>{group}</option>)}</select></td>
                <td><button className="danger-button" onClick={() => onChange(formulas.map((formula) => formula.id === selected.id ? { ...formula, details: formula.details.filter((item) => item.id !== row.id) } : formula))}>Xóa</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function EditableTableTab({ title, rows, fields, labels, roleOptions, onChange }) {
  const updateRow = (id, field, value) => onChange(rows.map((row) => row.id === id ? { ...row, [field]: value } : row))
  const addRow = () => onChange([{ id: newId('row'), ...Object.fromEntries(fields.map((field) => [field, ''])) }, ...rows])
  const deleteRow = (id) => onChange(rows.filter((row) => row.id !== id))
  return (
    <section className="panel settings-section">
      <div className="section-heading-row"><h3>{title}</h3><button className="primary-button" onClick={addRow}>Thêm</button></div>
      <div className="table-wrapper">
        <table className="settings-wide-table">
          <thead><tr>{labels.map((label) => <th key={label}>{label}</th>)}<th>Hành động</th></tr></thead>
          <tbody>
            {rows.map((row) => (
              <tr key={row.id}>
                {fields.map((field) => (
                  <td key={field}>
                    {field === 'role' && roleOptions ? (
                      <select value={row[field]} onChange={(event) => updateRow(row.id, field, event.target.value)}>{roleOptions.map((role) => <option key={role}>{role}</option>)}</select>
                    ) : (
                      <input value={row[field]} type={['min', 'max', 'tolerance'].includes(field) ? 'number' : 'text'} onChange={(event) => updateRow(row.id, field, ['min', 'max', 'tolerance'].includes(field) ? Number(event.target.value) : event.target.value)} />
                    )}
                  </td>
                ))}
                <td><button className="danger-button" onClick={() => deleteRow(row.id)}>Xóa</button></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function QrTab({ qrRules, materials, onChange }) {
  const [selectedCode, setSelectedCode] = useState(materials[0]?.code || '')
  const [lot, setLot] = useState('LOT-001')
  const [importDate, setImportDate] = useState(new Date().toISOString().slice(0, 10))
  const material = materials.find((item) => item.code === selectedCode) || materials[0]
  const qrValue = [
    qrRules.materialCode && `MVT:${material?.code || ''}`,
    qrRules.materialName && `TEN:${material?.name || ''}`,
    qrRules.lot && `LOT:${lot}`,
    qrRules.importDate && `NGAY:${importDate}`,
  ].filter(Boolean).join('|')
  const updateRule = (field, value) => onChange({ ...qrRules, [field]: value, preview: qrValue })
  return (
    <section className="panel settings-section">
      <h3>Thiết lập QR</h3>
      <div className="qr-settings-layout">
        <div className="settings-form-grid">
          <label>Vật tư<select value={selectedCode} onChange={(event) => setSelectedCode(event.target.value)}>{materials.map((item) => <option key={item.id} value={item.code}>{item.code} - {item.name}</option>)}</select></label>
          <label>Lô nhập<input value={lot} onChange={(event) => setLot(event.target.value)} /></label>
          <label>Ngày nhập<input type="date" value={importDate} onChange={(event) => setImportDate(event.target.value)} /></label>
          {['materialCode', 'materialName', 'lot', 'importDate'].map((field) => (
            <label className="checkbox-field" key={field}><input type="checkbox" checked={qrRules[field]} onChange={(event) => updateRule(field, event.target.checked)} /> {field}</label>
          ))}
        </div>
        <div className="qr-preview-panel">
          <PseudoQr value={qrValue} />
          <code>{qrValue}</code>
          <div className="action-row">
            <button className="primary-button" onClick={() => onChange({ ...qrRules, preview: qrValue })}>Sinh QR</button>
            <button className="secondary-button" onClick={() => window.print()}>In QR</button>
            <button className="secondary-button" onClick={() => alert(qrValue)}>Xem QR</button>
          </div>
        </div>
      </div>
    </section>
  )
}

function ProductionSettingsTab({ settings, onChange }) {
  const updateField = (field, value) => onChange({ ...settings, [field]: value })
  const updateRule = (index, value) => onChange({ ...settings, rules: settings.rules.map((rule, currentIndex) => currentIndex === index ? value : rule) })
  return (
    <section className="panel settings-section">
      <h3>Thiết lập sản xuất</h3>
      <div className="settings-form-grid">
        <label>Số cân hóa<input type="number" value={settings.chemicalScaleCount} onChange={(event) => updateField('chemicalScaleCount', Number(event.target.value))} /></label>
        <label>Số cân rắn<input type="number" value={settings.solidScaleCount} onChange={(event) => updateField('solidScaleCount', Number(event.target.value))} /></label>
        <label>Thời gian cảnh báo<input type="number" value={settings.warningMinutes} onChange={(event) => updateField('warningMinutes', Number(event.target.value))} /></label>
      </div>
      <h3>Quy tắc chuyển trạng thái</h3>
      <div className="settings-rule-list">
        {settings.rules.map((rule, index) => <input key={index} value={rule} onChange={(event) => updateRule(index, event.target.value)} />)}
      </div>
    </section>
  )
}

function BackupTab({ settings, onRestore, importRef }) {
  const exportAll = () => {
    const productionData = JSON.parse(localStorage.getItem(PRODUCTION_KEY) || '{}')
    downloadJson('son-hoa-binh-full-backup.json', { ...productionData, systemSettings: settings, exportedAt: new Date().toISOString() })
  }
  const importBackup = (file) => {
    if (!file) return
    const reader = new FileReader()
    reader.onload = (event) => {
      const data = JSON.parse(event.target.result)
      if (data.systemSettings) {
        onRestore({ ...defaultSettings, ...data.systemSettings })
      }
      if (data.productionOrders) {
        safeSetJsonLocalStorage(PRODUCTION_KEY, data)
      }
    }
    reader.readAsText(file)
  }
  return (
    <section className="panel settings-section">
      <h3>Sao lưu dữ liệu</h3>
      <p className="panel-text">Dữ liệu gồm productionOrders, formulas, materials, chemicalWeighing, solidWeighing, mixingLogs, qcLogs, productionLogs, users.</p>
      <div className="backup-actions">
        <button className="primary-button" onClick={exportAll}>Xuất toàn bộ dữ liệu</button>
        <button className="secondary-button" onClick={() => importRef.current?.click()}>Import dữ liệu</button>
        <button className="secondary-button" onClick={exportAll}>Backup hệ thống</button>
        <button className="secondary-button" onClick={() => importRef.current?.click()}>Khôi phục dữ liệu</button>
        <input ref={importRef} type="file" accept=".json" hidden onChange={(event) => importBackup(event.target.files?.[0])} />
      </div>
    </section>
  )
}

function FactoryTab({ factory, onChange }) {
  const updateField = (field, value) => onChange({ ...factory, [field]: value })
  return (
    <section className="panel settings-section">
      <h3>Thông số nhà máy</h3>
      <div className="settings-form-grid">
        <label>Tên công ty<input value={factory.companyName} onChange={(event) => updateField('companyName', event.target.value)} /></label>
        <label>Tên nhà máy<input value={factory.factoryName} onChange={(event) => updateField('factoryName', event.target.value)} /></label>
        <label>Địa chỉ<input value={factory.address} onChange={(event) => updateField('address', event.target.value)} /></label>
        <label>Logo<input value={factory.logo} onChange={(event) => updateField('logo', event.target.value)} placeholder="URL hoặc tên file logo" /></label>
        <label>Điện thoại<input value={factory.phone} onChange={(event) => updateField('phone', event.target.value)} /></label>
        <label>Email<input value={factory.email} onChange={(event) => updateField('email', event.target.value)} /></label>
      </div>
      <div className="factory-preview">
        <strong>{factory.companyName}</strong>
        <span>{factory.factoryName}</span>
        <span>{factory.address}</span>
        <span>{factory.phone} - {factory.email}</span>
      </div>
    </section>
  )
}
