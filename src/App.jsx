import { Children, cloneElement, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import { QRCodeCanvas } from 'qrcode.react'
import { Sidebar } from './components/Sidebar.jsx'
import { TopBar } from './components/TopBar.jsx'
import { defaultNavItems } from './data/navigation.js'
import { USE_SUPABASE } from './utils/supabaseMode.js'
import './App.css'

const DATA_KEY = 'sonhoabinh-v3-data'
const PRODUCTION_ORDERS_KEY = 'productionOrders'
const FORMULAS_KEY = 'formulas'
const PRODUCTION_LOGS_KEY = 'productionLogs'
const QC2_LOGS_KEY = 'qc2Logs'
const QC2_ADJUSTMENTS_KEY = 'qc2Adjustments'
const SUPPLEMENTAL_WEIGHING_KEY = 'supplementalWeighing'
const WEIGHED_CONTAINERS_KEY = 'weighedContainers'
const PACKING_LOGS_KEY = 'packingLogs'
const FINISHED_GOODS_KEY = 'finishedGoods'
const MATERIAL_CATALOG_KEY = 'materialCatalog'
const AUTH_KEY = 'sonhoabinh-v3-auth'
const SESSION_KEY = 'sonhoabinh-v3-session'

const CHEMICAL = 'Hóa'
const SOLID = 'Rắn'
const nonEmptyArray = (...items) => items.find((item) => Array.isArray(item) && item.length > 0) || []
const nowText = () => new Date().toISOString().slice(0, 16).replace('T', ' ')
const todayText = () => new Date().toISOString().slice(0, 10)
const num = (value) => Number(value) || 0
const kg = (value) => `${num(value).toLocaleString('vi-VN', { maximumFractionDigits: 3 })} kg`
const uid = (prefix) => `${prefix}-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 7)}`
const RAW_MATERIAL_QR_TYPE = 'RAW_MATERIAL_LOT'

function parseScaleWeightKg(text = '') {
  const normalized = String(text)
    .replace(/\r/g, '\n')
    .replace(/,/g, '.')
    .toLowerCase()
  const kgMatches = [...normalized.matchAll(/[-+]?\s*\d+(?:\.\d+)?\s*kg/g)]
  const source = kgMatches.at(-1)?.[0] || [...normalized.matchAll(/[-+]?\s*\d+(?:\.\d+)?/g)].at(-1)?.[0] || ''
  if (!source) return null
  const value = Number(source.replace(/kg/g, '').replace(/\s/g, ''))
  return Number.isFinite(value) ? Math.abs(value) : null
}

function rawMaterialLotStatus(lot) {
  const initialQty = num(lot.initialQty)
  const remainingQty = num(lot.remainingQty)
  const minStock = num(lot.minStock)
  if (remainingQty <= 0) return 'Hết tồn'
  if (remainingQty <= minStock || (initialQty > 0 && remainingQty <= initialQty * 0.1)) return 'Sắp hết'
  return 'Còn tồn'
}

function buildRawMaterialLotQr(lot) {
  return JSON.stringify({
    type: RAW_MATERIAL_QR_TYPE,
    materialCode: lot.materialCode,
    materialName: lot.materialName,
    lotCode: lot.lotCode,
    supplier: lot.supplier,
    importDate: lot.importDate,
    remainingQty: num(lot.remainingQty),
    unit: lot.unit || 'kg',
  })
}

function normalizeRawMaterialLot(item = {}) {
  const initialQty = num(item.initialQty ?? item.weight ?? item.quantity ?? item.qty)
  const remainingQty = item.remainingQty == null ? initialQty : num(item.remainingQty)
  const lot = {
    ...item,
    id: item.id || uid('RM'),
    materialCode: item.materialCode || '',
    materialName: item.materialName || item.name || item.materialCode || '',
    materialGroup: item.materialGroup || item.group || CHEMICAL,
    lotCode: item.lotCode || item.lot || '',
    supplier: item.supplier || '',
    importDate: item.importDate || todayText(),
    initialQty,
    remainingQty,
    unit: item.unit || 'kg',
    minStock: num(item.minStock),
  }
  const issuedQty = Math.max(0, num(lot.initialQty) - num(lot.remainingQty))
  const status = item.status || rawMaterialLotStatus(lot)
  const qrCode = buildRawMaterialLotQr({ ...lot, status })
  return {
    ...lot,
    lot: lot.lotCode,
    weight: lot.initialQty,
    issuedQty,
    status,
    qrCode,
  }
}

function normalizeRawMaterialLots(items = []) {
  return items.map(normalizeRawMaterialLot)
}

function normalizeMaterialCatalogItem(item = {}) {
  const materialCode = String(item.materialCode || item.code || item['Mã vật tư'] || item['Ma vat tu'] || item['Mã VT'] || item['Ma VT'] || '').trim()
  if (!materialCode) return null
  return {
    id: item.id || `MAT-${materialCode}`,
    materialCode,
    materialName: String(item.materialName || item.name || item['Tên vật tư'] || item['Ten vat tu'] || item['Tên VT'] || item['Ten VT'] || materialCode).trim(),
    materialGroup: String(item.materialGroup || item.group || item['Nhóm vật tư'] || item['Nhom vat tu'] || item['Nhóm'] || item['Nhom'] || CHEMICAL).trim(),
    unit: String(item.unit || item['Đơn vị tính'] || item['Don vi tinh'] || item['Đơn vị'] || item['Don vi'] || 'kg').trim(),
  }
}

function normalizeMaterialCatalog(items = []) {
  const byCode = new Map()
  ;(items || []).forEach((item) => {
    const normalized = normalizeMaterialCatalogItem(item)
    if (normalized) byCode.set(normalized.materialCode.toUpperCase(), normalized)
  })
  return Array.from(byCode.values()).sort((a, b) => a.materialCode.localeCompare(b.materialCode, 'vi', { numeric: true }))
}

function deriveMaterialCatalog(data = {}) {
  const fromFormulas = (data.formulas || []).flatMap((formula) => formula.items || [])
  const fromRawMaterials = normalizeRawMaterialLots(data.rawMaterials || [])
  return normalizeMaterialCatalog([...fromFormulas, ...fromRawMaterials, ...(data.materialCatalog || [])])
}

function mergeMaterialCatalog(current = [], incoming = []) {
  return normalizeMaterialCatalog([...current, ...incoming])
}

function parseRawMaterialQr(value) {
  try {
    const parsed = JSON.parse(value)
    return parsed && typeof parsed === 'object' ? parsed : null
  } catch {
    return null
  }
}

function validateRawMaterialQr(qrInput, requiredMaterial, lots = []) {
  const qr = parseRawMaterialQr(qrInput)
  if (!qr || qr.type !== RAW_MATERIAL_QR_TYPE) {
    return { ok: false, message: 'QR sai type.', qr }
  }
  if (qr.materialCode !== requiredMaterial.materialCode) {
    return { ok: false, message: `Sai nguyên liệu. Lệnh yêu cầu ${requiredMaterial.materialCode}, QR đang là ${qr.materialCode}.`, qr }
  }
  const lot = lots.find((item) => item.lotCode === qr.lotCode && item.materialCode === qr.materialCode)
  if (!lot) {
    return { ok: false, message: 'Không tìm thấy lô nguyên liệu trong kho.', qr }
  }
  if (num(lot.remainingQty) <= 0) {
    return { ok: false, message: 'Lô nguyên liệu đã hết tồn.', qr, lot }
  }
  return { ok: true, qr, lot }
}

function buildFormulaItems(lines, quantityKg) {
  return lines.map((line, index) => ({
    id: uid('mat'),
    no: index + 1,
    materialCode: line.code,
    materialName: line.name || line.code,
    materialGroup: line.group,
    ratioPercent: line.percent,
    requiredKg: Number(((line.percent * quantityKg) / 100).toFixed(3)),
    toleranceKg: line.group === CHEMICAL ? 0.01 : 0.1,
    qcConfirm: true,
    qcAdjustPercent: '',
    qcAdjustKg: '',
    qrScanned: '',
    qrStatus: 'Chờ quét',
    actualWeight: '',
    weighStatus: 'Chờ cân',
    confirmedAt: '',
    note: '',
  }))
}

const masterFormulaLines = {
  'HNS-252-G1': [
    ['PASTE 02', CHEMICAL, 4.61], ['IN02', CHEMICAL, 0.02], ['IN03', CHEMICAL, 0.30], ['D01', CHEMICAL, 0.07],
    ['R91', SOLID, 20], ['KT01', CHEMICAL, 0.5], ['SiG01', SOLID, 15], ['SiBK02', SOLID, 5], ['SW34', SOLID, 34.5], ['SW92', SOLID, 20],
  ],
  'HNS-252-R2': [
    ['PASTE 02', CHEMICAL, 4.61], ['IN02', CHEMICAL, 0.02], ['IN03', CHEMICAL, 0.30], ['D01', CHEMICAL, 0.07],
    ['R91', SOLID, 20], ['KT01', CHEMICAL, 0.5], ['SiR05', SOLID, 19], ['SiR01', SOLID, 1], ['SW34', SOLID, 34.5], ['SW92', SOLID, 20],
  ],
}

const defaultMixingMachines = [
  { machineCode: 'M01', machineName: 'MÁY 1', capacityKg: 250, motorPower: '7.5HP', department: 'Tổ phối trộn', productionTeam: 'Tổ phối trộn', status: 'READY', note: '' },
  { machineCode: 'M02', machineName: 'MÁY 2', capacityKg: 250, motorPower: '7.5HP', department: 'Tổ phối trộn', productionTeam: 'Tổ phối trộn', status: 'READY', note: '' },
  { machineCode: 'M03', machineName: 'MÁY 3', capacityKg: 550, motorPower: '10HP', department: 'Tổ phối trộn', productionTeam: 'Tổ phối trộn', status: 'READY', note: '' },
  { machineCode: 'M04', machineName: 'MÁY 4', capacityKg: 550, motorPower: '10HP', department: 'Tổ phối trộn', productionTeam: 'Tổ phối trộn', status: 'READY', note: '' },
  { machineCode: 'M05', machineName: 'MÁY 5', capacityKg: 1050, motorPower: '20HP', department: 'Tổ phối trộn', productionTeam: 'Tổ phối trộn', status: 'READY', note: '' },
  { machineCode: 'M06', machineName: 'MÁY 6', capacityKg: 1050, motorPower: '20HP', department: 'Tổ phối trộn', productionTeam: 'Tổ phối trộn', status: 'READY', note: '' },
]
const realMixingMachineCodes = new Set(defaultMixingMachines.map((machine) => machine.machineCode))
const legacyMixingMachineCodes = new Set(['MX01', 'MX02', 'MX03', 'MX04'])
const legacyMixingMachineMap = { MX01: 'M01', MX02: 'M02', MX03: 'M03', MX04: 'M04' }
const normalizeMixingMachineCode = (code) => legacyMixingMachineMap[String(code || '').trim().toUpperCase()] || String(code || '').trim().toUpperCase()

const normalizeMixingMachineStatus = (status) => {
  if (status === 'Hoạt động' || status === 'Rảnh' || status === 'Đang chạy') return 'READY'
  if (status === 'Ngừng sử dụng' || status === 'Dừng' || status === 'Stopped') return 'INACTIVE'
  if (status === 'Bảo trì' || status === 'Lỗi') return 'MAINTENANCE'
  return ['READY', 'INACTIVE', 'MAINTENANCE'].includes(status) ? status : 'READY'
}

const normalizeMixingMachine = (item = {}) => ({
  ...item,
  machineCode: normalizeMixingMachineCode(item.machineCode),
  machineName: item.machineName || '',
  motorPower: item.motorPower || item.machinePower || '',
  capacityKg: num(item.capacityKg),
  department: item.department || item.productionTeam || 'Tổ phối trộn',
  productionTeam: item.productionTeam || item.department || 'Tổ phối trộn',
  status: normalizeMixingMachineStatus(item.status),
  note: item.note || '',
})

const normalizeMixingMachines = (items = []) => {
  const byCode = new Map()
  defaultMixingMachines.forEach((item) => {
    const machine = normalizeMixingMachine(item)
    byCode.set(machine.machineCode, machine)
  })
  ;(items || []).forEach((item) => {
    const rawCode = String(item.machineCode || '').trim().toUpperCase()
    const machine = normalizeMixingMachine(item)
    if (!machine.machineCode || legacyMixingMachineCodes.has(rawCode)) return
    if (realMixingMachineCodes.has(machine.machineCode) && !item.catalogOverride) return
    byCode.set(machine.machineCode, machine)
  })
  return Array.from(byCode.values()).sort((a, b) => a.machineCode.localeCompare(b.machineCode, 'vi', { numeric: true }))
}

const isActiveMixingMachine = (machine = {}) => normalizeMixingMachineStatus(machine.status) === 'READY'
const getActiveMixingMachines = (machines = []) => machines.filter(isActiveMixingMachine)
const kgPerBatch = (value) => `${num(value).toLocaleString('vi-VN', { maximumFractionDigits: 3 })}kg/mẻ`
const formatMachineCapacityKg = (value) => `${num(value).toLocaleString('vi-VN', { maximumFractionDigits: 3 })}kg`
const formatMixingMachineLabel = (machine = {}) => {
  const name = machine.machineName || machine.assignedMachineName || machine.name || ''
  const capacity = machine.capacityKg ?? machine.assignedMachineCapacityKg
  if (name && capacity) return `${name} - ${formatMachineCapacityKg(capacity)}`
  return name || machine.machineCode || '-'
}
const mixingMachineOptionLabel = formatMixingMachineLabel
const getMixingMachineLabelByCode = (code = '', machines = []) => {
  const normalizedCode = normalizeMixingMachineCode(code)
  if (!normalizedCode) return '-'
  const machine = machines.find((item) => item.machineCode === normalizedCode)
  return machine ? formatMixingMachineLabel(machine) : normalizedCode
}
const getOrderAssignedMachineCode = (order = {}) => order.assignedMachineCode || order.mixerMachine || order.assignedMixingMachine || order.mixingMachine || order.mixing?.machineCode || ''
const getOrderAssignedMachineLabel = (order = {}, machines = []) => {
  const code = getOrderAssignedMachineCode(order)
  if (!code) return '-'
  const machine = machines.find((item) => item.machineCode === code)
  const name = order.assignedMachineName || machine?.machineName
  const capacity = order.assignedMachineCapacityKg ?? machine?.capacityKg
  return formatMixingMachineLabel({ machineCode: code, machineName: name, capacityKg: capacity })
}

const productionAssignmentStages = ['Cân hóa', 'Cân rắn', 'Phối trộn', 'QC sản xuất thử', 'QC thành phẩm', 'Đóng gói', 'Kho thành phẩm']
const productionProcessOptions = [
  ['CHEMICAL_WEIGHING', 'Cân hóa'],
  ['SOLID_WEIGHING', 'Cân rắn'],
  ['MIXING', 'Phối trộn'],
  ['QC_TRIAL', 'QC sản xuất thử'],
  ['QC_FINISHED', 'QC thành phẩm'],
  ['PACKAGING', 'Đóng gói'],
  ['FINISHED_GOODS', 'Kho thành phẩm'],
]
const processCodeByName = Object.fromEntries(productionProcessOptions.map(([code, name]) => [name, code]))
const roleAssignmentStageMap = {
  'Cân hóa': ['Cân hóa'],
  'Cân rắn': ['Cân rắn'],
  'Phối trộn': ['Phối trộn'],
}
const getStageAssignmentsForRole = (assignments = [], role = '') => {
  const stages = roleAssignmentStageMap[role]
  return stages ? assignments.filter((item) => stages.includes(item.stage || item.processName)) : assignments
}
const getActiveAssignments = (assignments = [], stage, date = todayText(), shiftCode = '') => (
  assignments.filter((item) => (
    (item.date || item.workDate) === date
    && (item.stage || item.processName) === stage
    && item.status !== 'Hủy'
    && (!shiftCode || item.shiftCode === shiftCode)
  ))
)
const assignmentEmployeeCodes = (assignment = {}) => (
  Array.isArray(assignment.employeeCodes) && assignment.employeeCodes.length
    ? assignment.employeeCodes
    : [assignment.employeeCode].filter(Boolean)
)
const assignmentEmployeeNames = (assignment = {}) => (
  Array.isArray(assignment.employeeNames) && assignment.employeeNames.length
    ? assignment.employeeNames
    : [assignment.employeeName].filter(Boolean)
)
const formatAssignmentEmployees = (assignment = {}) => {
  const names = assignmentEmployeeNames(assignment)
  if (names.length > 1) return `${names.length} người: ${names.join(', ')}`
  const codes = assignmentEmployeeCodes(assignment)
  return names[0] ? `${codes[0] ? `${codes[0]} - ` : ''}${names[0]}` : '-'
}
const assignmentMachineNames = (assignment = {}) => (
  Array.isArray(assignment.machineNames) && assignment.machineNames.length
    ? assignment.machineNames
    : [assignment.machineName || assignment.machineCode].filter(Boolean)
)
const formatAssignmentMachines = (assignment = {}) => {
  const names = assignmentMachineNames(assignment)
  return names.length ? names.join(', ') : 'Không áp dụng'
}
const getAssignmentLogContext = (assignments = []) => {
  const activeAssignments = assignments.filter(Boolean)
  const employeeCodes = [...new Set(activeAssignments.flatMap(assignmentEmployeeCodes))]
  const employeeNames = [...new Set(activeAssignments.flatMap(assignmentEmployeeNames))]
  const assignmentIds = activeAssignments.map((item) => item.assignmentId || item.id).filter(Boolean)
  return {
    assignmentId: assignmentIds.join(', '),
    assignmentIds,
    employeeCodes,
    employeeNames,
    employeeCode: employeeCodes.join(', '),
    employeeName: employeeNames.join(', '),
    employee: employeeNames.join(', '),
  }
}

const productionEmployeeCatalog = [
  ['NV001', 'LÊ PHƯƠNG BÌNH', 'Tổ trộn 1', 'Tổ trưởng', 'Phối trộn'],
  ['NV002', 'NGUYỄN HỮU CHÍ', 'Tổ trộn 1', 'Công nhân', 'Phối trộn'],
  ['NV003', 'LÊ VĂN TRỌNG', 'Tổ trộn 1', 'Công nhân', 'Phối trộn'],
  ['NV004', 'ĐÀM ANH TUẤN', 'Tổ trộn 1', 'Công nhân', 'Phối trộn'],
  ['NV005', 'NGUYỄN VĂN TÂN', 'Tổ trộn 2', 'Tổ trưởng', 'Phối trộn'],
  ['NV006', 'NGÔ ĐẠI ÂN', 'Tổ trộn 2', 'Công nhân', 'Phối trộn'],
  ['NV007', 'NGUYỄN ĐÌNH PHI', 'Tổ trộn 2', 'Công nhân', 'Phối trộn'],
  ['NV008', 'VÕ THANH SỰ', 'Tổ trộn 2', 'Công nhân', 'Phối trộn'],
  ['NV009', 'NGUYỄN ĐẠI PHÚ', 'Tổ Hóa', 'Tổ trưởng', 'Cân hóa'],
  ['NV010', 'NGUYỄN PHÙNG DŨNG', 'Tổ Hóa', 'Công nhân', 'Cân hóa'],
  ['NV011', 'BÙI TÁ DUY', 'Tổ Hóa', 'Công nhân', 'Cân hóa'],
  ['NV012', 'NGUYỄN THÀNH KHIÊN', 'Tổ Hóa', 'Công nhân', 'Cân hóa'],
  ['NV013', 'SƠN HOÀNH', 'Tổ Hóa', 'Công nhân', 'Cân hóa'],
  ['NV014', 'NGUYỄN KHÁNH TƯỜNG', 'Tổ Hóa', 'Công nhân', 'Cân hóa'],
  ['NV015', 'TRẦN MINH KHA', 'Tổ Cát', 'Tổ trưởng', 'Cân rắn'],
  ['NV016', 'TRƯƠNG MINH HOÀNG', 'Tổ Cát', 'Công nhân', 'Cân rắn'],
  ['NV017', 'NGUYỄN VỸ', 'Tổ Cát', 'Công nhân', 'Cân rắn'],
  ['NV018', 'PHẠM THỊ HẠNH', 'QC', 'QC sản xuất', 'QC sản xuất thử'],
  ['NV019', 'TRẦN QUỐC BẢO', 'QC', 'QC thành phẩm', 'QC thành phẩm'],
  ['NV020', 'LÊ MINH TÚ', 'Đóng gói', 'Công nhân', 'Đóng gói'],
  ['NV021', 'VÕ THỊ LAN', 'Kho thành phẩm', 'Thủ kho', 'Kho thành phẩm'],
].map(([code, name, productionTeam, title, operationRole]) => ({
  id: `EMP-${code}`,
  code,
  name,
  productionTeam,
  title,
  operationRole,
  status: 'Hoạt động',
  qrEmployee: `EMP:${code}`,
}))

const initialData = {
  rawMaterials: [
    normalizeRawMaterialLot({ id: 'RM-001', materialCode: 'PASTE 02', materialName: 'Paste nền 02', materialGroup: CHEMICAL, lotCode: 'NVL-260613-01', importDate: '2026-06-13', supplier: 'HB Chemical', initialQty: 320, remainingQty: 320, unit: 'kg' }),
    normalizeRawMaterialLot({ id: 'RM-002', materialCode: 'R91', materialName: 'Bột R91', materialGroup: SOLID, lotCode: 'NVL-260613-02', importDate: '2026-06-13', supplier: 'Silica Việt', initialQty: 800, remainingQty: 800, unit: 'kg' }),
    normalizeRawMaterialLot({ id: 'RM-003', materialCode: 'SW34', materialName: 'SW34', materialGroup: SOLID, lotCode: 'NVL-260612-01', importDate: '2026-06-12', supplier: 'Khoáng sản HB', initialQty: 1000, remainingQty: 1000, unit: 'kg' }),
  ],
  formulas: Object.entries(masterFormulaLines).map(([id, rows], index) => ({
    id,
    code: id.replaceAll('-', ' '),
    product: id === 'HNS-252-G1' ? 'HNS 252.G1' : 'HNS 252.R2',
    version: index === 0 ? 'V3.0' : 'V3.1',
    effectiveDate: '2026-06-13',
    createdBy: 'Phòng kỹ thuật',
    checkedBy: 'QC trưởng',
    approvedBy: 'Giám đốc sản xuất',
    items: rows.map(([code, group, percent], no) => ({ id: uid('fml'), no: no + 1, materialCode: code, materialName: code, materialGroup: group, ratioPercent: percent })),
  })),
  formulaVersions: [],
  orders: [],
  logs: [
    { id: uid('log'), time: '2026-06-13 08:00', entry: 'Khởi tạo quy trình sản xuất V3: Kho NVL, công thức gốc, QC1, QC2, đóng gói và kho thành phẩm.' },
  ],
  productionLogs: [
    { id: uid('log'), time: '2026-06-13 08:00', entry: 'Khởi tạo quy trình sản xuất V3: Kho NVL, công thức gốc, QC1, QC2, đóng gói và kho thành phẩm.' },
  ],
  qc2Logs: [],
  qc2AdjustmentTickets: [],
  supplementalWeighing: [],
  weighedContainers: [],
  stockTransactions: [],
  packingLogs: [],
  finishedGoods: [],
  productCatalog: [
    { id: 'PRD-HNS-252-G1', code: 'HNS 252.G1', name: 'HNS 252.G1', group: 'Sơn', unit: 'kg', status: 'Hoạt động', note: 'Sản phẩm demo' },
    { id: 'PRD-HNS-252-R2', code: 'HNS 252.R2', name: 'HNS 252.R2', group: 'Sơn', unit: 'kg', status: 'Hoạt động', note: 'Sản phẩm demo' },
  ],
  supplierCatalog: [
    { id: 'SUP-HB-CHEM', code: 'HB-CHEM', name: 'HB Chemical', phone: '', address: '', status: 'Hoạt động', note: '' },
    { id: 'SUP-SILICA-VIET', code: 'SILICA-VIET', name: 'Silica Việt', phone: '', address: '', status: 'Hoạt động', note: '' },
  ],
  customerCatalog: [
    { id: 'CUS-DEMO', code: 'CUS-DEMO', name: 'Đơn hàng demo', phone: '', address: '', status: 'Hoạt động', note: '' },
    { id: 'CUS-QC2', code: 'CUS-QC2', name: 'Đơn hàng mẫu QC2', phone: '', address: '', status: 'Hoạt động', note: '' },
  ],
  employeeCatalog: productionEmployeeCatalog,
  teamCatalog: [
    { id: 'TEAM-TP1', code: 'TP1', name: 'Tổ phối trộn 1', leader: 'Lê Phương Bình', note: '', status: 'Hoạt động' },
    { id: 'TEAM-TP2', code: 'TP2', name: 'Tổ phối trộn 2', leader: 'Nguyễn Văn Tân', note: '', status: 'Hoạt động' },
    { id: 'TEAM-TH', code: 'TH', name: 'Tổ Hóa', leader: 'Nguyễn Đại Phú', note: '', status: 'Hoạt động' },
    { id: 'TEAM-TC', code: 'TC', name: 'Tổ Cát', leader: 'Trần Minh Kha', note: '', status: 'Hoạt động' },
    { id: 'TEAM-QC', code: 'QC', name: 'QC', leader: 'Phạm Thị Hạnh', note: '', status: 'Hoạt động' },
    { id: 'TEAM-DG', code: 'DG', name: 'Đóng gói', leader: 'Lê Minh Tú', note: '', status: 'Hoạt động' },
    { id: 'TEAM-KTP', code: 'KTP', name: 'Kho thành phẩm', leader: 'Võ Thị Lan', note: '', status: 'Hoạt động' },
  ],
  shiftCatalog: [
    { id: 'SHIFT-C1', code: 'C1', name: 'Ca ngày', startTime: '07:00', endTime: '17:00', note: '', status: 'Hoạt động' },
    { id: 'SHIFT-C2', code: 'C2', name: 'Ca tăng ca', startTime: '17:00', endTime: '21:00', note: '', status: 'Hoạt động' },
  ],
  productionAssignments: [],
  mixingMachines: defaultMixingMachines,
}

function seedData() {
  const formula = initialData.formulas[0]
  const items = buildFormulaItems(formula.items.map((item) => ({ code: item.materialCode, name: item.materialName, group: item.materialGroup, percent: item.ratioPercent })), 1000)
  return {
    ...initialData,
    orders: [
      {
        id: 'LSX-260613-001',
        product: formula.product,
        lot: 'LOT-HNS-G1-001',
        customer: 'Đơn hàng demo',
        quantityKg: 1000,
        stage: 'qc1',
        status: 'Chờ QC sản xuất thử',
        createdAt: '2026-06-13 08:15',
        updatedAt: '2026-06-13 08:15',
        originalFormulaId: formula.id,
        originalFormulaVersion: formula.version,
        originalFormula: formula.items,
        productionFormulaSnapshot: items,
        qc1AdjustedFormula: null,
        qc2AdjustedFormula: [],
        scaleStatus: { chemical: 'Pending', solid: 'Pending' },
      },
      {
        id: 'LSX-260613-002',
        product: 'HNS 252.R2',
        lot: 'LOT-HNS-R2-001',
        customer: 'Đơn hàng mẫu QC2',
        quantityKg: 750,
        stage: 'finished-qc',
        status: 'Chờ QC thành phẩm',
        createdAt: '2026-06-13 07:40',
        updatedAt: '2026-06-13 11:20',
        originalFormulaId: initialData.formulas[1].id,
        originalFormulaVersion: initialData.formulas[1].version,
        originalFormula: initialData.formulas[1].items,
        productionFormulaSnapshot: buildFormulaItems(initialData.formulas[1].items.map((item) => ({ code: item.materialCode, name: item.materialName, group: item.materialGroup, percent: item.ratioPercent })), 750).map((item) => ({ ...item, qrStatus: 'PASS', weighStatus: 'PASS', actualWeight: item.requiredKg, confirmedAt: '2026-06-13 09:30' })),
        qc1AdjustedFormula: null,
        qc2AdjustedFormula: [],
        scaleStatus: { chemical: 'Completed', solid: 'Completed' },
        mixing: { status: 'Completed', finalWeightKg: 748, completedAt: '2026-06-13 11:20', operator: 'Tổ phối trộn' },
      },
    ],
  }
}

function ensureQcDemoOrders(orders = [], seedOrders = []) {
  const hasQcTrial = orders.some((order) => order.stage === 'qc1')
  const hasFinishedQc = orders.some((order) => order.stage === 'finished-qc' || ['Chờ QC thành phẩm', 'Đang QC thành phẩm', 'Cần điều chỉnh', 'QC thành phẩm không đạt'].includes(order.status))
  const requiredSeeds = [
    !hasQcTrial && seedOrders.find((order) => order.stage === 'qc1'),
    !hasFinishedQc && seedOrders.find((order) => order.stage === 'finished-qc' || order.status === 'Chờ QC thành phẩm'),
  ].filter(Boolean)
  if (!requiredSeeds.length) return orders
  const seedIds = new Set(requiredSeeds.map((order) => order.id))
  return [...requiredSeeds, ...orders.filter((order) => !seedIds.has(order.id))]
}

const CRUD_ACTIONS = ['view', 'create', 'edit', 'delete']
const masterPermissionGroups = [
  ['material', 'Dữ liệu gốc / Danh mục vật tư'],
  ['product', 'Dữ liệu gốc / Danh mục sản phẩm'],
  ['supplier', 'Dữ liệu gốc / Danh mục nhà cung cấp'],
  ['customer', 'Dữ liệu gốc / Danh mục khách hàng'],
  ['employee', 'Dữ liệu gốc / Danh sách nhân viên'],
  ['team', 'Dữ liệu gốc / Danh mục tổ sản xuất'],
  ['shift', 'Dữ liệu gốc / Danh mục ca làm việc'],
  ['machine', 'Dữ liệu gốc / Danh mục máy phối trộn'],
  ['formula', 'Dữ liệu gốc / Công thức gốc'],
]
const masterPermissionIds = masterPermissionGroups.flatMap(([key]) => CRUD_ACTIONS.map((action) => `master.${key}.${action}`))
const productionPermissionGroups = [
  ['assignment', 'Sản xuất / Phân công nhân sự'],
]
const productionPermissionIds = productionPermissionGroups.flatMap(([key]) => CRUD_ACTIONS.map((action) => `production.${key}.${action}`))
const productionExtraPermissionIds = ['production.trace.view', 'production.log.view']
const masterFull = (key) => CRUD_ACTIONS.map((action) => `master.${key}.${action}`)
const masterView = (key) => [`master.${key}.view`]
const productionFull = (key) => CRUD_ACTIONS.map((action) => `production.${key}.${action}`)
const productionView = (key) => [`production.${key}.view`]
const systemPermissionIds = ['admin']
const menuPermissionIds = defaultNavItems.map((item) => item.permission || (item.type !== 'group' ? item.id : '')).filter((id, index, items) => id && items.indexOf(id) === index)
const allSystemPermissionIds = Array.from(new Set([...menuPermissionIds, ...masterPermissionIds, ...productionPermissionIds, ...productionExtraPermissionIds, ...systemPermissionIds, 'formula.secure.view']))
const hasPermission = (permissions = [], permission) => permissions.includes(permission)
const hasAnyPermission = (permissions = [], permissionIds = []) => permissionIds.some((permission) => hasPermission(permissions, permission))
const pagePermission = (item) => item.permission || item.id

const defaultRoles = {
  Admin: allSystemPermissionIds,
  'Kho NL': ['dashboard', 'raw-materials', 'production.log.view', ...masterView('material'), ...masterView('supplier')],
  'Kỹ thuật': ['dashboard', 'production.log.view', ...masterFull('material'), ...masterFull('product'), ...masterFull('formula'), ...masterView('machine'), ...masterView('supplier'), 'formula.secure.view'],
  'Sản xuất': ['dashboard', 'orders', 'logs', 'production.log.view', ...productionFull('assignment'), ...masterView('employee'), ...masterView('team'), ...masterView('shift'), ...masterView('product'), ...masterView('machine'), ...masterView('customer')],
  QC: ['dashboard', 'qc', 'finished-qc', 'production.log.view', 'reports', ...masterView('product'), ...masterView('formula')],
  'Cân hóa': ['dashboard', 'chemical', 'logs', 'production.log.view', ...productionView('assignment')],
  'Cân rắn': ['dashboard', 'solid', 'logs', 'production.log.view', ...productionView('assignment')],
  'Phối trộn': ['dashboard', 'mixing', 'logs', 'production.log.view', ...productionView('assignment'), ...masterView('machine')],
  'Đóng gói': ['dashboard', 'packaging', 'production.log.view'],
  'Kho TP': ['dashboard', 'finished-goods', 'production.log.view', 'reports'],
  'Quản đốc': ['dashboard', 'orders', 'qc', 'chemical', 'solid', 'mixing', 'finished-qc', 'packaging', 'finished-goods', 'logs', 'reports', 'production.trace.view', 'production.log.view', 'production.assignment.view', 'production.assignment.create', 'production.assignment.edit', ...masterView('employee'), ...masterView('team'), ...masterView('shift'), ...masterView('machine'), ...masterView('product'), ...masterView('customer')],
  'Ban giám đốc': ['dashboard', 'logs', 'reports', 'production.trace.view', 'production.log.view', ...productionView('assignment'), ...masterView('material'), ...masterView('product'), ...masterView('supplier'), ...masterView('customer'), ...masterView('employee'), ...masterView('team'), ...masterView('shift'), ...masterView('machine'), ...masterView('formula')],
}
const officialRoleNames = Object.keys(defaultRoles)

const ACTIVE_STATUS = 'Hoạt động'
const LOCKED_STATUS = 'Khóa'
const DEFAULT_PASSWORD = '123456'
const AUTH_VERSION = 3
const removedDefaultUsernames = new Set(['kho-nvl', 'kho-tp', 'hcns', 'kinhdoanh'])

const defaultUsers = [
  { username: 'admin', password: DEFAULT_PASSWORD, role: 'Admin', fullName: 'Quản trị hệ thống' },
  { username: 'kho.nl', password: DEFAULT_PASSWORD, role: 'Kho NL', fullName: 'Kho nguyên liệu' },
  { username: 'kythuat', password: DEFAULT_PASSWORD, role: 'Kỹ thuật', fullName: 'Phòng kỹ thuật' },
  { username: 'sanxuat', password: DEFAULT_PASSWORD, role: 'Sản xuất', fullName: 'Phòng sản xuất' },
  { username: 'qc', password: DEFAULT_PASSWORD, role: 'QC', fullName: 'QC sản xuất' },
  { username: 'canhoa', password: DEFAULT_PASSWORD, role: 'Cân hóa', fullName: 'Tổ cân hóa' },
  { username: 'canran', password: DEFAULT_PASSWORD, role: 'Cân rắn', fullName: 'Tổ cân rắn' },
  { username: 'phoitron', password: DEFAULT_PASSWORD, role: 'Phối trộn', fullName: 'Tổ phối trộn' },
  { username: 'donggoi', password: DEFAULT_PASSWORD, role: 'Đóng gói', fullName: 'Tổ đóng gói' },
  { username: 'kho.tp', password: DEFAULT_PASSWORD, role: 'Kho TP', fullName: 'Kho thành phẩm' },
  { username: 'quandoc', password: DEFAULT_PASSWORD, role: 'Quản đốc', fullName: 'Quản đốc' },
  { username: 'giamdoc', password: DEFAULT_PASSWORD, role: 'Ban giám đốc', fullName: 'Ban giám đốc' },
].map((user) => ({ ...user, department: user.role, status: ACTIVE_STATUS }))

const enforcedDefaultUserRoles = {
  'kho.tp': 'Kho TP',
  quandoc: 'Quản đốc',
  giamdoc: 'Ban giám đốc',
}

const legacyRoleMap = {
  'Kho nguyên liệu': 'Kho NL',
  'Phòng kỹ thuật': 'Kỹ thuật',
  'Phòng sản xuất': 'Sản xuất',
  'Tổ cân hóa': 'Cân hóa',
  'Tổ cân rắn': 'Cân rắn',
  'Tổ phối trộn': 'Phối trộn',
  'Kho thành phẩm': 'Kho TP',
  HCNS: 'Sản xuất',
  'Kinh doanh': 'Admin',
}

const defaultAuth = {
  users: defaultUsers,
  roles: defaultRoles,
  accessLogs: [],
}

function expandLegacyPermissions(permissions = []) {
  const expanded = new Set(permissions)
  const legacyMap = {
    formulas: masterFull('formula'),
    'admin-machines': masterFull('machine'),
    'master-materials': masterFull('material'),
    'master-products': masterFull('product'),
    'master-suppliers': masterFull('supplier'),
    'master-customers': masterFull('customer'),
    'master-employees': masterFull('employee'),
    'master-teams': masterFull('team'),
    'master-shifts': masterFull('shift'),
    'production-assignments': productionFull('assignment'),
    logs: ['production.log.view'],
    'reports-trace': ['production.trace.view'],
    admin: ['admin'],
    'admin-users': ['admin'],
    'admin-roles': ['admin'],
    'admin-permissions': ['admin'],
    'admin-system-logs': ['admin'],
  }
  permissions.forEach((permission) => {
    ;(legacyMap[permission] || []).forEach((item) => expanded.add(item))
  })
  return Array.from(expanded)
}

function normalizeAuthData(saved = {}) {
  const roles = { ...defaultRoles }
  Object.entries(saved.roles || {}).forEach(([role, permissions]) => {
    const nextRole = legacyRoleMap[role] || role
    if (!defaultRoles[nextRole]) return
    roles[nextRole] = Array.from(new Set([...(roles[nextRole] || []), ...expandLegacyPermissions(permissions || [])]))
  })
  roles.Admin = allSystemPermissionIds

  const savedUsers = (saved.users || []).filter((user) => user?.username && !removedDefaultUsernames.has(user.username))
  const savedUsersByUsername = new Map(savedUsers.map((user) => [user.username, user]))
  const seededUsers = defaultUsers.map((user) => {
    const savedUser = savedUsersByUsername.get(user.username)
    if (!savedUser) return user
    const role = enforcedDefaultUserRoles[user.username] || legacyRoleMap[savedUser.role] || savedUser.role || user.role
    return {
      ...user,
      ...savedUser,
      password: DEFAULT_PASSWORD,
      role,
      department: role,
      status: savedUser.status || ACTIVE_STATUS,
    }
  })
  const defaultUsernames = new Set(defaultUsers.map((user) => user.username))
  const migratedSavedUsers = savedUsers
    .filter((user) => user?.username && !defaultUsernames.has(user.username))
    .map((user) => {
      const role = legacyRoleMap[user.role] || user.role || 'Sản xuất'
      const nextRole = officialRoleNames.includes(role) ? role : 'Sản xuất'
      return {
        ...user,
        password: user.password || DEFAULT_PASSWORD,
        role: nextRole,
        department: nextRole,
        status: user.status || ACTIVE_STATUS,
      }
    })

  return {
    ...defaultAuth,
    ...saved,
    authVersion: AUTH_VERSION,
    roles,
    users: [...seededUsers, ...migratedSavedUsers],
    accessLogs: saved.accessLogs || [],
  }
}

function loadStored(key, fallback) {
  try {
    if (!USE_SUPABASE && String(key).toLowerCase().includes('supabase')) return fallback
    const raw = localStorage.getItem(key)
    return raw ? JSON.parse(raw) : fallback
  } catch {
    return fallback
  }
}

function operationLogMeta(user, { employee = '', employeeCode = '', employeeName = '', employeeCodes = [], employeeNames = [], assignmentId = '', assignmentIds = [], assignments = [], stage = '', order, material = {}, machine = {}, actionType = '', result = '', targetQty = '', actualQty = '' } = {}) {
  const orderCode = typeof order === 'string' ? order : (order?.orderCode || order?.id || '')
  const assignmentContext = getAssignmentLogContext(assignments)
  const normalizedEmployeeCodes = employeeCodes.length ? employeeCodes : assignmentContext.employeeCodes
  const normalizedEmployeeNames = employeeNames.length ? employeeNames : assignmentContext.employeeNames
  const inferredEmployeeName = employeeName || normalizedEmployeeNames.join(', ') || employee || user?.fullName || user?.username || 'Chưa xác định'
  const normalizedAssignmentIds = assignmentIds.length ? assignmentIds : assignmentContext.assignmentIds
  return {
    username: user?.username || '',
    user: user?.fullName || user?.username || '',
    userAccount: user?.username || '',
    userRole: user?.role || '',
    assignmentId: assignmentId || normalizedAssignmentIds.join(', '),
    assignmentIds: normalizedAssignmentIds,
    employee: employee || normalizedEmployeeNames.join(', '),
    employeeCodes: normalizedEmployeeCodes,
    employeeNames: normalizedEmployeeNames,
    employeeCode: employeeCode || normalizedEmployeeCodes.join(', '),
    employeeName: inferredEmployeeName,
    role: user?.role || '',
    stage,
    processName: stage || '',
    orderCode,
    productionOrderCode: orderCode,
    lotCode: typeof order === 'object' ? (order?.lot || '') : '',
    materialCode: material.materialCode || '',
    materialName: material.materialName || '',
    targetQty,
    actualQty,
    machineCode: machine.machineCode || '',
    machineName: machine.machineName || machine.name || '',
    actionType: actionType || result || 'Thao tác',
    actionDescription: actionType || result || '',
    actionTime: nowText(),
    resultStatus: result,
    result,
  }
}

function addLogToData(data, entry, meta = {}) {
  const log = { id: uid('log'), time: nowText(), entry, ...meta }
  return {
    ...data,
    logs: [...(data.logs || []), log],
    productionLogs: [...(data.productionLogs || data.logs || []), log],
  }
}

function getEffectiveFormula(order) {
  return order.activeProductionFormula || order.qc1AdjustedFormula || order.productionFormulaSnapshot || []
}

function normalizeProductionOrders(orders = [], formulas = []) {
  const machineCatalog = normalizeMixingMachines()
  return orders.map((order, index) => {
    const formula = formulas.find((item) => item.id === (order.formulaId || order.originalFormulaId))
      || formulas.find((item) => item.code === order.formulaCode)
      || formulas[0]
    const fallbackCode = order.orderCode || order.id || `LSX-${todayText().replaceAll('-', '')}-${String(index + 1).padStart(3, '0')}`
    const originalFormulaSnapshot = order.originalFormulaSnapshot || order.productionFormulaSnapshot || order.ingredients || []
    const activeProductionFormula = order.activeProductionFormula || order.qc1AdjustedFormula || order.productionFormulaSnapshot || originalFormulaSnapshot
    const requestedWeight = num(order.requestedWeight ?? order.quantityKg)
    const assignedMachineCode = normalizeMixingMachineCode(order.assignedMachineCode || order.mixerMachine || order.assignedMixingMachine || order.mixingMachine || order.mixing?.machineCode || '')
    const assignedMachine = machineCatalog.find((machine) => machine.machineCode === assignedMachineCode)
    return {
      ...order,
      id: order.id || fallbackCode,
      orderCode: fallbackCode,
      formulaId: order.formulaId || order.originalFormulaId || formula?.id || '',
      formulaCode: order.formulaCode || formula?.code || order.formula || '',
      formulaVersion: order.formulaVersion || order.originalFormulaVersion || formula?.version || '',
      productName: order.productName || order.product || formula?.product || '',
      product: order.product || order.productName || formula?.product || '',
      customer: order.customer || '',
      lot: order.lot || `LOT-${fallbackCode}`,
      requestedWeight,
      quantityKg: requestedWeight,
      status: order.status || 'Chờ QC sản xuất thử',
      stage: order.stage || 'qc1',
      createdAt: order.createdAt || nowText(),
      updatedAt: order.updatedAt || order.createdAt || nowText(),
      originalFormulaSnapshot,
      activeProductionFormula,
      productionFormulaSnapshot: order.productionFormulaSnapshot || originalFormulaSnapshot,
      qc1Adjustments: order.qc1Adjustments || order.qc1Logs || [],
      qc2Adjustments: order.qc2Adjustments || order.qc2AdjustedFormula || [],
      qc1AdjustedFormula: order.qc1AdjustedFormula || null,
      qc2AdjustedFormula: order.qc2AdjustedFormula || order.qc2Adjustments || [],
      qc2SupplementTickets: order.qc2SupplementTickets || order.qc2AdjustedFormula || [],
      initialOrderSnapshot: order.initialOrderSnapshot || {
        id: order.id || fallbackCode,
        orderCode: fallbackCode,
        productName: order.productName || order.product || formula?.product || '',
        lot: order.lot || `LOT-${fallbackCode}`,
        requestedWeight,
        createdAt: order.createdAt || nowText(),
        originalFormulaSnapshot,
      },
      chemicalStatus: order.chemicalStatus || order.scaleStatus?.chemical || 'Pending',
      solidStatus: order.solidStatus || order.scaleStatus?.solid || 'Pending',
      mixingStatus: order.mixingStatus || order.mixing?.status || 'Pending',
      mixerMachine: assignedMachineCode || order.mixerMachine || '',
      assignedMixingMachine: assignedMachineCode || order.assignedMixingMachine || '',
      assignedMachineCode,
      assignedMachineName: order.assignedMachineName || assignedMachine?.machineName || '',
      assignedMachineCapacityKg: order.assignedMachineCapacityKg ?? assignedMachine?.capacityKg ?? '',
      assignedMachineMotorPower: order.assignedMachineMotorPower || assignedMachine?.motorPower || '',
      assignedMachineDepartment: order.assignedMachineDepartment || assignedMachine?.department || assignedMachine?.productionTeam || '',
      machineChangeHistory: order.machineChangeHistory || [],
      machineAssignmentHistory: order.machineAssignmentHistory || [],
      mixingMachine: normalizeMixingMachineCode(order.mixingMachine || order.mixing?.machineCode || ''),
      mixingStartAt: order.mixingStartAt || order.mixing?.startedAt || '',
      mixingCompletedAt: order.mixingCompletedAt || order.mixing?.completedAt || '',
      mixingQrConfirmation: order.mixingQrConfirmation || {
        chemicalQr: order.chemicalContainerQr || '',
        solidQr: order.solidContainerQr || '',
        status: 'Chưa xác nhận',
        confirmedAt: '',
        note: '',
      },
      qc1Status: order.qc1Status || order.qc1Result || 'Chờ QC sản xuất thử',
      qc2Status: order.qc2Status || order.qc2?.result || 'Pending',
      packagingStatus: order.packagingStatus || order.packingStatus || (order.packaging ? 'Completed' : 'Pending'),
      packingStatus: order.packingStatus || order.packagingStatus || (order.packaging ? 'Completed' : 'Pending'),
      finishedGoodsStatus: order.finishedGoodsStatus || (order.stage === 'completed' ? 'Completed' : 'Pending'),
      scaleStatus: {
        chemical: order.scaleStatus?.chemical || order.chemicalStatus || 'Pending',
        solid: order.scaleStatus?.solid || order.solidStatus || 'Pending',
      },
    }
  })
}

function statusClass(status = '') {
  if (['PASS', 'OK', 'Completed', 'Hoàn thành', 'Đóng gói', 'Kho thành phẩm'].some((word) => status.includes(word))) return 'pass'
  if (['FAIL', 'Không đạt', 'Cần'].some((word) => status.includes(word))) return 'fail'
  if (['Active', 'Đang'].some((word) => status.includes(word))) return 'active'
  return 'locked'
}

function LoginPage({ onLogin, error }) {
  const [username, setUsername] = useState('admin')
  const [password, setPassword] = useState(DEFAULT_PASSWORD)
  return (
    <main className="login-shell">
      <form className="login-card" onSubmit={(event) => { event.preventDefault(); onLogin(username, password) }}>
        <span className="section-kicker">Sơn Hòa Bình V3</span>
        <h1>Đăng nhập hệ thống</h1>
        <label>Tài khoản<input value={username} onChange={(event) => setUsername(event.target.value)} /></label>
        <label>Mật khẩu<input type="password" value={password} onChange={(event) => setPassword(event.target.value)} /></label>
        {error && <div className="process-alert">{error}</div>}
        <button className="primary-button">Đăng nhập</button>
        <p>Tài khoản mặc định: <strong>admin / 123456</strong></p>
      </form>
    </main>
  )
}

function SimpleTable({ headers, rows, empty = 'Không có dữ liệu.', tableClassName = '' }) {
  const labeledRows = rows.map((row) => {
    if (!row?.props?.children) return row
    const cells = Children.toArray(row.props.children).map((cell, index) => {
      if (!cell?.props) return cell
      return cloneElement(cell, { 'data-label': headers[index] || '' })
    })
    return cloneElement(row, {}, cells)
  })
  return (
    <div className="table-wrapper">
      <table className={tableClassName}>
        <thead><tr>{headers.map((header) => <th key={header}>{header}</th>)}</tr></thead>
        <tbody>
          {labeledRows}
          {!rows.length && <tr><td className="empty-row" colSpan={headers.length}>{empty}</td></tr>}
        </tbody>
      </table>
    </div>
  )
}

function getQc2Adjustments(order = {}) {
  return order.qc2Adjustments || []
}

function getQc2SupplementTickets(order = {}) {
  return order.qc2SupplementTickets || order.qc2AdjustedFormula || []
}

function getAdjustmentItems(adjustment = {}) {
  return adjustment.items || adjustment.adjustmentItems || []
}

function getTicketItems(ticket = {}) {
  return ticket.items || []
}

function qc2AdjustmentId(no) {
  return `QC2-ADJ-${String(no).padStart(3, '0')}`
}

function nextQc2AdjustmentNo(data = {}) {
  const numbers = [
    ...(data.qc2AdjustmentTickets || []),
    ...(data.orders || []).flatMap((order) => getQc2Adjustments(order)),
  ].map((ticket) => Number(ticket.adjustmentNo || String(ticket.adjustmentId || ticket.id || '').match(/(\d+)$/)?.[1] || 0))
  return Math.max(0, ...numbers) + 1
}

function qc2AttemptCount(order, data = {}) {
  const resultLogs = (data.qc2Logs || []).filter((log) => (
    log.orderId === order.id
    && (log.result || ['QC2 đạt', 'QC2 cần điều chỉnh', 'QC2 không đạt'].includes(log.action))
  )).length
  return Math.max(resultLogs, getQc2Adjustments(order).length + (order.qc2 ? 1 : 0))
}

function countBy(rows, keyFn, valueFn = () => 1) {
  const totals = rows.reduce((map, row) => {
    const key = keyFn(row) || 'Không xác định'
    map.set(key, (map.get(key) || 0) + valueFn(row))
    return map
  }, new Map())
  return [...totals.entries()].sort((a, b) => b[1] - a[1])
}

function orderCodeText(order) {
  return String(order.orderCode || order.id || '')
}

function sortOldestOrders(a, b) {
  if (a.createdAt && b.createdAt && a.createdAt !== b.createdAt) {
    return String(a.createdAt).localeCompare(String(b.createdAt), 'vi', { numeric: true })
  }
  return orderCodeText(a).localeCompare(orderCodeText(b), 'vi', { numeric: true })
}

function findMaterialByCode(rows = [], code = '') {
  return rows.find((item) => String(item.materialCode || '').toUpperCase() === String(code || '').toUpperCase())
}

function buildQc2Rows(order = {}, currentAdjustments = []) {
  const originalRows = order.originalFormulaSnapshot || order.productionFormulaSnapshot || []
  const qc1Rows = order.activeProductionFormula || order.qc1AdjustedFormula || order.productionFormulaSnapshot || originalRows
  const codes = Array.from(new Set([...originalRows, ...qc1Rows].map((item) => item.materialCode).filter(Boolean)))
  return codes.map((code) => {
    const original = findMaterialByCode(originalRows, code) || {}
    const qc1 = findMaterialByCode(qc1Rows, code) || original
    const draft = findMaterialByCode(currentAdjustments, code) || {}
    const supplementKg = draft.adjustmentKg === '' || draft.adjustmentKg == null ? '' : draft.adjustmentKg
    const supplementNumber = supplementKg === '' ? 0 : num(supplementKg)
    const qc1Kg = num(qc1.requiredKg ?? original.requiredKg)
    return {
      id: qc1.id || original.id || code,
      materialCode: code,
      materialName: qc1.materialName || original.materialName || code,
      materialGroup: qc1.materialGroup || original.materialGroup || CHEMICAL,
      originalKg: num(original.requiredKg ?? original.materialPerLot),
      qc1Kg,
      supplementKg,
      afterQc2Kg: Number((qc1Kg + supplementNumber).toFixed(3)),
      reason: draft.reason || '',
      note: draft.note || '',
    }
  })
}

function mostSupplementedMaterial(items = []) {
  return countBy(items, (item) => item.materialCode, (item) => Math.max(0, num(item.adjustmentKg ?? item.requiredKg)))[0]?.[0] || '-'
}

function displayQc2Status(status = '') {
  const labels = {
    PendingSupplementWeighing: 'Chờ cân bổ sung QC2',
    SupplementWeighed: 'Đã cân bổ sung',
    Completed: 'Hoàn thành',
    Pending: 'Chờ xử lý',
  }
  return labels[status] || status || '-'
}

function displayQcTrialText(text = '') {
  const labels = {
    'QC mẫu': 'QC sản xuất thử',
    'Chờ QC1': 'Chờ QC sản xuất thử',
    'QC1 OK': 'QC sản xuất thử đạt',
    'QC1 đạt': 'QC sản xuất thử đạt',
    'QC1 điều chỉnh': 'Điều chỉnh sau QC sản xuất thử',
    'QC1 đã điều chỉnh & duyệt': 'Điều chỉnh sau QC sản xuất thử',
    'Đang test mẫu SX': 'Đang QC sản xuất thử',
    'Lệnh đang test mẫu': 'Lệnh đang QC sản xuất thử',
  }
  return labels[text] || text || '-'
}

const packagingSpecs = [
  { id: 'box25', label: 'Thùng 25 kg', sizeKg: 25, toleranceKg: 0.2 },
  { id: 'box10', label: 'Thùng 10 kg', sizeKg: 10, toleranceKg: 0.1 },
  { id: 'box5', label: 'Thùng 5 kg', sizeKg: 5, toleranceKg: 0.05 },
]

function qc2FinalWeight(order = {}) {
  return num(order.qc2FinalWeight ?? order.mixing?.finalWeightKg ?? order.mixingFinalWeightKg ?? order.quantityKg)
}

function normalizeWeighedContainers(containers = []) {
  return (containers || []).filter(Boolean).map((item, index) => {
    const qrCode = item.qrCode || item.containerQr || `MIX-${index + 1}`
    return {
      containerId: item.containerId || item.id || `CNT-${qrCode}`,
      qrCode,
      orderId: item.orderId || '',
      orderCode: item.orderCode || item.orderId || '',
      productName: item.productName || item.product || '',
      lot: item.lot || '',
      materialGroup: item.materialGroup || '',
      materials: item.materials || [],
      totalWeight: num(item.totalWeight ?? item.weight),
      weighedBy: item.weighedBy || item.operator || '',
      completedAt: item.completedAt || item.createdAt || '',
      weighingType: item.weighingType || 'Cân chính',
      status: item.status || 'Đã cân xong',
    }
  })
}

function materialGroupCode(group) {
  return group === CHEMICAL ? 'HOA' : 'RAN'
}

function createContainerQrCode(group, containers = []) {
  const date = todayText().replaceAll('-', '')
  const prefix = `MIX-${materialGroupCode(group)}-LSX-${date}`
  const nextNo = normalizeWeighedContainers(containers).filter((item) => String(item.qrCode || '').startsWith(prefix)).length + 1
  return `${prefix}-${String(nextNo).padStart(3, '0')}`
}

function getContainerByQr(containers = [], qr = '') {
  const value = String(qr || '').trim().toUpperCase()
  if (!value) return null
  return normalizeWeighedContainers(containers).find((item) => String(item.qrCode || '').trim().toUpperCase() === value)
}

function getOrderGroupContainers(containers = [], order = {}, weighingType = 'Cân chính') {
  const orderCode = order.orderCode || order.id
  return normalizeWeighedContainers(containers).filter((item) => (
    (item.orderId === order.id || item.orderCode === orderCode)
    && item.weighingType === weighingType
  ))
}

function buildWeighedContainer(order, group, items, containers = [], weighingType = 'Cân chính') {
  const qrCode = createContainerQrCode(group, containers)
  const completedAt = nowText()
  const materials = items.map((item) => ({
    id: item.id,
    materialCode: item.materialCode,
    materialName: item.materialName || item.materialCode,
    materialGroup: item.materialGroup,
    requiredKg: num(item.requiredKg),
    actualWeight: num(item.actualWeight || item.requiredKg),
    qrScanned: item.qrScanned || '',
    qrStatus: item.qrStatus || '',
    qrMatchStatus: item.qrMatchStatus || '',
    confirmedAt: item.confirmedAt || completedAt,
    weighedBy: group === CHEMICAL ? 'Tổ cân hóa' : 'Tổ cân rắn',
  }))
  return {
    containerId: `CNT-${qrCode}`,
    qrCode,
    orderId: order.id,
    orderCode: order.orderCode || order.id,
    productName: order.productName || order.product,
    lot: order.lot,
    materialGroup: group,
    materials,
    totalWeight: Number(materials.reduce((sum, item) => sum + num(item.actualWeight), 0).toFixed(3)),
    weighedBy: group === CHEMICAL ? 'Tổ cân hóa' : 'Tổ cân rắn',
    completedAt,
    weighingType,
    status: 'Đã cân xong',
  }
}

function updateContainerStatuses(containers = [], qrCodes = [], status) {
  const qrSet = new Set(qrCodes.filter(Boolean).map((item) => String(item).trim().toUpperCase()))
  return normalizeWeighedContainers(containers).map((item) => qrSet.has(String(item.qrCode || '').trim().toUpperCase()) ? { ...item, status } : item)
}

function getContainerQrValue(container = {}) {
  return JSON.stringify({
    qrCode: container.qrCode || '',
    orderCode: container.orderCode || '',
    productName: container.productName || '',
    lot: container.lot || '',
    materialGroup: container.materialGroup || '',
    totalWeight: num(container.totalWeight),
    weighingType: container.weighingType || '',
    status: container.status || '',
  })
}

function demoQrMaterials(group, entries, time) {
  return entries.map(([code, weight], index) => ({
    id: `DEMO-QR-${group}-${code}-${index}`,
    materialCode: code,
    materialName: code,
    materialGroup: group,
    requiredKg: weight,
    actualWeight: weight,
    qrScanned: code,
    confirmedAt: time,
  }))
}

function buildDemoQrContainers() {
  return [
    {
      containerId: 'CNT-MIX-HOA-LSX-20260614-001',
      qrCode: 'MIX-HOA-LSX-20260614-001',
      orderId: 'LSX-20260613-003',
      orderCode: 'LSX-20260613-003',
      productName: 'HNS 252.G1',
      lot: 'LOT-LSX-20260613-003',
      materialGroup: CHEMICAL,
      materials: demoQrMaterials(CHEMICAL, [['PASTE02', 46.1], ['IN02', 0.2], ['IN03', 3], ['D01', 0.7], ['KT01', 5]], '2026-06-14 10:00'),
      totalWeight: 55,
      weighedBy: 'Tổ cân hóa',
      completedAt: '2026-06-14 10:02',
      weighingType: 'Cân chính',
      status: 'Đã cân xong',
    },
    {
      containerId: 'CNT-MIX-HOA-LSX-20260614-002',
      qrCode: 'MIX-HOA-LSX-20260614-002',
      orderId: 'LSX-20260613-024',
      orderCode: 'LSX-20260613-024',
      productName: 'HNS 252.G1',
      lot: 'LOT-LSX-20260613-024',
      materialGroup: CHEMICAL,
      materials: demoQrMaterials(CHEMICAL, [['PASTE02', 47.5], ['IN02', 0.2], ['IN03', 3.2], ['D01', 0.8], ['KT01', 5.3]], '2026-06-14 10:40'),
      totalWeight: 57,
      weighedBy: 'Tổ cân hóa',
      completedAt: '2026-06-14 10:42',
      weighingType: 'Cân chính',
      status: 'Đã cân xong',
    },
    {
      containerId: 'CNT-MIX-HOA-LSX-20260614-003',
      qrCode: 'MIX-HOA-LSX-20260614-003',
      orderId: 'LSX-QC2-DEMO-001',
      orderCode: 'LSX-QC2-DEMO-001',
      productName: 'HNS 252.G1',
      lot: 'LOT-QC2-G1-001',
      materialGroup: CHEMICAL,
      materials: demoQrMaterials(CHEMICAL, [['PASTE02', 1.8], ['IN03', 0.5], ['KT01', 0.5]], '2026-06-14 11:00'),
      totalWeight: 2.8,
      weighedBy: 'Tổ cân hóa',
      completedAt: '2026-06-14 11:02',
      weighingType: 'Cân bổ sung QC2',
      status: 'Đã cân xong',
    },
    {
      containerId: 'CNT-MIX-RAN-LSX-20260614-001',
      qrCode: 'MIX-RAN-LSX-20260614-001',
      orderId: 'LSX-20260613-003',
      orderCode: 'LSX-20260613-003',
      productName: 'HNS 252.G1',
      lot: 'LOT-LSX-20260613-003',
      materialGroup: SOLID,
      materials: demoQrMaterials(SOLID, [['R91', 200], ['SIG01', 150], ['SIBK02', 50], ['SW34', 345], ['SW92', 200]], '2026-06-14 10:15'),
      totalWeight: 945,
      weighedBy: 'Tổ cân rắn',
      completedAt: '2026-06-14 10:16',
      weighingType: 'Cân chính',
      status: 'Đã cân xong',
    },
    {
      containerId: 'CNT-MIX-RAN-LSX-20260614-002',
      qrCode: 'MIX-RAN-LSX-20260614-002',
      orderId: 'LSX-20260613-024',
      orderCode: 'LSX-20260613-024',
      productName: 'HNS 252.G1',
      lot: 'LOT-LSX-20260613-024',
      materialGroup: SOLID,
      materials: demoQrMaterials(SOLID, [['R91', 198], ['SIG01', 150], ['SIBK02', 50], ['SW34', 345], ['SW92', 200]], '2026-06-14 10:55'),
      totalWeight: 943,
      weighedBy: 'Tổ cân rắn',
      completedAt: '2026-06-14 10:56',
      weighingType: 'Cân chính',
      status: 'Đã cân xong',
    },
    {
      containerId: 'CNT-MIX-RAN-LSX-20260614-003',
      qrCode: 'MIX-RAN-LSX-20260614-003',
      orderId: 'LSX-QC2-DEMO-001',
      orderCode: 'LSX-QC2-DEMO-001',
      productName: 'HNS 252.G1',
      lot: 'LOT-QC2-G1-001',
      materialGroup: SOLID,
      materials: demoQrMaterials(SOLID, [['SW34', 5], ['SW92', 3]], '2026-06-14 11:15'),
      totalWeight: 8,
      weighedBy: 'Tổ cân rắn',
      completedAt: '2026-06-14 11:16',
      weighingType: 'Cân bổ sung QC2',
      status: 'Đã cân xong',
    },
  ]
}

function buildDemoQrOrder(config) {
  return {
    id: config.id,
    orderCode: config.id,
    product: 'HNS 252.G1',
    productName: 'HNS 252.G1',
    lot: config.lot,
    customer: config.customer || 'Demo QR hỗn hợp',
    quantityKg: config.quantityKg || 1000,
    requestedWeight: config.quantityKg || 1000,
    stage: config.stage || 'mixing',
    status: config.stage === 'mixing-supplement' ? 'Chờ phối trộn bổ sung' : 'Chờ phối trộn',
    createdAt: config.createdAt || '2026-06-13 08:00',
    updatedAt: '2026-06-14 10:35',
    originalFormulaId: 'HNS-252-G1',
    originalFormulaVersion: 'V3.0',
    formulaId: 'HNS-252-G1',
    formulaCode: 'HNS 252 G1',
    formulaVersion: 'V3.0',
    originalFormulaSnapshot: [],
    productionFormulaSnapshot: [],
    activeProductionFormula: [],
    qc1Adjustments: [],
    qc2Adjustments: [],
    qc2SupplementTickets: [],
    ChemicalCompleted: true,
    SolidCompleted: true,
    ReadyMixing: true,
    chemicalStatus: 'Completed',
    solidStatus: 'Completed',
    scaleStatus: { chemical: 'Completed', solid: 'Completed' },
    mixingStatus: 'Pending',
    mixingQrConfirmation: {
      chemicalQr: config.chemicalQr || '',
      solidQr: config.solidQr || '',
      status: 'Chưa xác nhận',
      confirmedAt: '',
      note: config.note || '',
    },
  }
}

function buildDemoQrLogs() {
  return [
    ['DEMO-QR-LOG-001', '2026-06-14 10:00', 'Hoàn thành cân hóa LSX-20260613-003'],
    ['DEMO-QR-LOG-002', '2026-06-14 10:02', 'Tạo QR hỗn hợp Hóa MIX-HOA-LSX-20260614-001'],
    ['DEMO-QR-LOG-003', '2026-06-14 10:15', 'Hoàn thành cân rắn LSX-20260613-003'],
    ['DEMO-QR-LOG-004', '2026-06-14 10:16', 'Tạo QR hỗn hợp Rắn MIX-RAN-LSX-20260614-001'],
    ['DEMO-QR-LOG-005', '2026-06-14 10:20', 'In QR hỗn hợp Hóa'],
    ['DEMO-QR-LOG-006', '2026-06-14 10:21', 'In QR hỗn hợp Rắn'],
    ['DEMO-QR-LOG-007', '2026-06-14 10:30', 'Quét QR hỗn hợp Hóa'],
    ['DEMO-QR-LOG-008', '2026-06-14 10:31', 'Quét QR hỗn hợp Rắn'],
    ['DEMO-QR-LOG-009', '2026-06-14 10:32', 'Xác nhận QR PASS'],
    ['DEMO-QR-LOG-010', '2026-06-14 10:35', 'Bắt đầu phối trộn'],
  ].map(([id, time, entry]) => ({ id, time, entry }))
}

function applyDemoQrData(data = {}) {
  const existingContainers = normalizeWeighedContainers(data.weighedContainers || [])
  const existingQrCodes = new Set(existingContainers.map((item) => item.qrCode))
  const demoContainers = buildDemoQrContainers().filter((item) => !existingQrCodes.has(item.qrCode))
  const existingOrderCodes = new Set((data.orders || []).map((order) => order.orderCode || order.id))
  const demoOrders = [
    buildDemoQrOrder({ id: 'LSX-20260613-003', lot: 'LOT-LSX-20260613-003', chemicalQr: 'MIX-HOA-LSX-20260614-001', solidQr: 'MIX-RAN-LSX-20260614-001', customer: 'Demo QR PASS', createdAt: '2026-06-13 08:03' }),
    buildDemoQrOrder({ id: 'LSX-20260613-024', lot: 'LOT-LSX-20260613-024', chemicalQr: 'MIX-HOA-LSX-20260614-001', solidQr: 'MIX-RAN-LSX-20260614-002', customer: 'Demo QR sai LOT', createdAt: '2026-06-13 08:24' }),
    buildDemoQrOrder({ id: 'LSX-20260613-025', lot: 'LOT-LSX-20260613-003', chemicalQr: 'MIX-HOA-LSX-20260614-001', solidQr: 'MIX-HOA-LSX-20260614-002', customer: 'Demo QR sai nhóm', createdAt: '2026-06-13 08:25' }),
    buildDemoQrOrder({ id: 'LSX-QC2-DEMO-001', lot: 'LOT-QC2-G1-001', chemicalQr: 'MIX-HOA-LSX-20260614-003', solidQr: 'MIX-RAN-LSX-20260614-003', customer: 'Demo QR bổ sung QC2', quantityKg: 10.8, stage: 'mixing-supplement', createdAt: '2026-06-13 09:00' }),
  ].filter((order) => !existingOrderCodes.has(order.orderCode || order.id))
  const existingLogIds = new Set((data.productionLogs || data.logs || []).map((log) => log.id))
  const demoLogs = buildDemoQrLogs().filter((log) => !existingLogIds.has(log.id))
  return {
    ...data,
    weighedContainers: [...existingContainers, ...demoContainers],
    orders: [...(data.orders || []), ...demoOrders],
    productionLogs: [...(data.productionLogs || data.logs || []), ...demoLogs],
    logs: [...(data.logs || data.productionLogs || []), ...demoLogs],
  }
}

function defaultPackingDetails() {
  return packagingSpecs.map((spec) => ({
    ...spec,
    boxes: 0,
    actualWeight: 0,
    note: '',
  }))
}

function getPackingForm(forms, order) {
  const saved = forms[order.id] || order.packaging || {}
  return {
    packer: saved.packer || saved.operator || '',
    startedAt: saved.startedAt || '',
    completedAt: saved.completedAt || '',
    notes: saved.notes || saved.note || '',
    details: packagingSpecs.map((spec) => {
      const row = (saved.details || saved.packingDetails || []).find((item) => item.id === spec.id || item.sizeKg === spec.sizeKg) || {}
      return {
        ...spec,
        boxes: row.boxes || 0,
        actualWeight: row.actualWeight || 0,
        note: row.note || '',
      }
    }),
  }
}

function packingTotals(order, form) {
  const details = form.details || defaultPackingDetails()
  const qcWeight = qc2FinalWeight(order)
  const totalPackedWeight = Number(details.reduce((sum, item) => sum + num(item.actualWeight), 0).toFixed(3))
  const convertedWeight = Number(details.reduce((sum, item) => sum + num(item.sizeKg) * num(item.boxes), 0).toFixed(3))
  const totalTolerance = Number(details.reduce((sum, item) => sum + num(item.toleranceKg) * num(item.boxes), 0).toFixed(3))
  const remainingWeight = Number((qcWeight - totalPackedWeight).toFixed(3))
  const differenceWeight = Number((totalPackedWeight - qcWeight).toFixed(3))
  return { qcWeight, totalPackedWeight, convertedWeight, totalTolerance, remainingWeight, differenceWeight }
}

function getLatestPackingLog(order = {}, packingLogs = []) {
  return packingLogs
    .filter((log) => log.orderId === order.id || log.orderCode === (order.orderCode || order.id))
    .at(-1)
}

function packingSpecSummary(details = []) {
  return details
    .filter((item) => num(item.boxes) > 0)
    .map((item) => `${item.spec || item.label || `${item.sizeKg} kg`}: ${num(item.boxes)} thùng`)
    .join(', ') || '-'
}

function totalPackingBoxes(details = []) {
  return details.reduce((sum, item) => sum + num(item.boxes), 0)
}

function nextFinishedCode(finishedGoods = []) {
  const date = todayText().replaceAll('-', '')
  const prefix = `TP-${date}-`
  const maxNo = finishedGoods
    .map((item) => String(item.finishedCode || item.finishedGoodsCode || item.id || ''))
    .filter((code) => code.startsWith(prefix))
    .map((code) => Number(code.slice(prefix.length)) || 0)
    .reduce((max, value) => Math.max(max, value), 0)
  return `${prefix}${String(maxNo + 1).padStart(3, '0')}`
}

function normalizeFinishedGoodsData(items = []) {
  return (items || []).map((item, index) => ({
    id: item.id || item.finishedGoodsId || item.finishedCode || `FG-${index + 1}`,
    finishedCode: item.finishedCode || item.finishedGoodsCode || item.id || `TP-CU-${String(index + 1).padStart(3, '0')}`,
    orderId: item.orderId || item.orderCode || '',
    orderCode: item.orderCode || item.orderId || '',
    productName: item.productName || item.finishedCode || item.product || '',
    product: item.product || item.productName || item.finishedCode || '',
    lot: item.lot || '',
    spec: item.spec || item.packagingSpec || '-',
    boxes: num(item.boxes ?? item.quantity),
    weight: num(item.weight ?? item.importWeight ?? item.totalPackedWeight),
    importDate: item.importDate || item.date || todayText(),
    location: item.location || item.warehouseLocation || '',
    receiver: item.receiver || item.importer || item.createdBy || '',
    status: item.status || 'Hoàn thành',
    note: item.note || item.notes || '',
  }))
}

function filterFinishedGoods(items, filters) {
  return items.filter((item) => {
    if (filters.fromDate && String(item.importDate || '').slice(0, 10) < filters.fromDate) return false
    if (filters.toDate && String(item.importDate || '').slice(0, 10) > filters.toDate) return false
    if (filters.orderCode && !String(item.orderCode || item.orderId || '').toLowerCase().includes(filters.orderCode.toLowerCase())) return false
    if (filters.product && !String(item.productName || item.product || '').toLowerCase().includes(filters.product.toLowerCase())) return false
    if (filters.lot && !String(item.lot || '').toLowerCase().includes(filters.lot.toLowerCase())) return false
    if (filters.location && !String(item.location || '').toLowerCase().includes(filters.location.toLowerCase())) return false
    return true
  })
}

function buildFinishedGoodsDemoOrder(config) {
  const details = config.details.map((item) => ({
    id: `pkg-${item.sizeKg}`,
    spec: `Thùng ${item.sizeKg} kg`,
    label: `Thùng ${item.sizeKg} kg`,
    sizeKg: item.sizeKg,
    boxes: item.boxes,
    convertedWeight: item.sizeKg * item.boxes,
    toleranceKg: item.sizeKg === 25 ? 0.2 : item.sizeKg === 10 ? 0.1 : 0.05,
    actualWeight: item.weight,
    note: item.note || '',
  }))
  const totalPackedWeight = Number(details.reduce((sum, item) => sum + num(item.actualWeight), 0).toFixed(3))
  const totalBoxes = totalPackingBoxes(details)
  const spec = packingSpecSummary(details)
  const packingLog = {
    packingId: `PKG-${config.id}`,
    orderId: config.id,
    orderCode: config.id,
    productName: config.productName,
    lot: config.lot,
    qc2FinalWeight: config.qc2FinalWeight,
    packingDetails: details,
    totalPackedWeight,
    remainingWeight: Number((config.qc2FinalWeight - totalPackedWeight).toFixed(3)),
    differenceWeight: Number((totalPackedWeight - config.qc2FinalWeight).toFixed(3)),
    packer: config.packer,
    startedAt: config.startedAt,
    completedAt: config.completedAt,
    status: 'completed',
    notes: config.notes || 'Dữ liệu demo kho thành phẩm',
  }
  const order = {
    id: config.id,
    orderCode: config.id,
    customer: config.customer,
    product: config.productName,
    productName: config.productName,
    lot: config.lot,
    requestedWeight: config.qc2FinalWeight,
    quantityKg: config.qc2FinalWeight,
    createdAt: config.createdAt,
    updatedAt: config.completedAt,
    stage: 'finished-goods',
    status: 'Chờ nhập kho thành phẩm',
    orderStatus: 'Chờ nhập kho thành phẩm',
    qc2Status: 'Đạt',
    qc2: { result: 'Đạt', checkedAt: config.qc2At || config.startedAt, orderId: config.id },
    packingStatus: 'completed',
    packagingStatus: 'Completed',
    finishedGoodsStatus: 'pending',
    packaging: {
      details,
      totalPackedWeight,
      remainingWeight: packingLog.remainingWeight,
      differenceWeight: packingLog.differenceWeight,
      packer: config.packer,
      startedAt: config.startedAt,
      completedAt: config.completedAt,
      packingLogId: packingLog.packingId,
      spec,
      boxes: totalBoxes,
    },
    mixing: { status: 'Completed', finalWeightKg: config.qc2FinalWeight, completedAt: config.qc2At || config.startedAt },
    mixingStatus: 'completed',
    chemicalStatus: 'Completed',
    solidStatus: 'Completed',
    scaleStatus: { chemical: 'Completed', solid: 'Completed' },
  }
  return { order, packingLog }
}

function createFinishedGoodsDemoPayload(current = {}) {
  const waitingConfigs = [
    { id: 'LSX-KTP-DEMO-001', customer: 'Công ty Minh Long', productName: 'HNS 252.G1', lot: 'LOT-KTP-001', qc2FinalWeight: 1000, details: [{ sizeKg: 25, boxes: 40, weight: 1000 }], packer: 'Tổ đóng gói A', startedAt: '2026-06-13 14:00', completedAt: '2026-06-13 15:10', createdAt: '2026-06-13 08:00' },
    { id: 'LSX-KTP-DEMO-002', customer: 'Công ty An Phú', productName: 'HNS 252.G1', lot: 'LOT-KTP-002', qc2FinalWeight: 800, details: [{ sizeKg: 10, boxes: 80, weight: 800 }], packer: 'Tổ đóng gói B', startedAt: '2026-06-13 14:20', completedAt: '2026-06-13 15:30', createdAt: '2026-06-13 08:20' },
    { id: 'LSX-KTP-DEMO-003', customer: 'Nhà máy Đông Á', productName: 'HNS 252.G1', lot: 'LOT-KTP-003', qc2FinalWeight: 500, details: [{ sizeKg: 5, boxes: 100, weight: 500 }], packer: 'Tổ đóng gói A', startedAt: '2026-06-13 14:40', completedAt: '2026-06-13 15:45', createdAt: '2026-06-13 08:40' },
    { id: 'LSX-KTP-DEMO-004', customer: 'Công ty Việt Sơn', productName: 'HNS 252.G1', lot: 'LOT-KTP-004', qc2FinalWeight: 1250, details: [{ sizeKg: 25, boxes: 40, weight: 1000 }, { sizeKg: 10, boxes: 25, weight: 250 }], packer: 'Tổ đóng gói C', startedAt: '2026-06-13 15:00', completedAt: '2026-06-13 16:10', createdAt: '2026-06-13 09:00' },
    { id: 'LSX-KTP-DEMO-005', customer: 'Công ty Nam Hải', productName: 'HNS 252.G1', lot: 'LOT-KTP-005', qc2FinalWeight: 650, details: [{ sizeKg: 10, boxes: 50, weight: 500 }, { sizeKg: 5, boxes: 30, weight: 150 }], packer: 'Tổ đóng gói B', startedAt: '2026-06-13 15:20', completedAt: '2026-06-13 16:25', createdAt: '2026-06-13 09:20' },
  ]
  const built = waitingConfigs.map(buildFinishedGoodsDemoOrder)
  const importedGoods = [
    { id: 'FG-KTP-DEMO-001', finishedCode: 'TP-20260613-101', orderId: 'LSX-KTP-IMPORTED-001', orderCode: 'LSX-KTP-IMPORTED-001', productName: 'HNS 252.G1', product: 'HNS 252.G1', lot: 'LOT-TP-001', spec: 'Thùng 25 kg: 40 thùng', boxes: 40, weight: 1000, importDate: '2026-06-13', location: 'KTP-A01', receiver: 'Thủ kho Lan', status: 'Đã nhập kho', note: 'Demo đã nhập kho' },
    { id: 'FG-KTP-DEMO-002', finishedCode: 'TP-20260613-102', orderId: 'LSX-KTP-IMPORTED-002', orderCode: 'LSX-KTP-IMPORTED-002', productName: 'HNS 252.G1', product: 'HNS 252.G1', lot: 'LOT-TP-002', spec: 'Thùng 10 kg: 80 thùng', boxes: 80, weight: 800, importDate: '2026-06-13', location: 'KTP-A02', receiver: 'Thủ kho Minh', status: 'Đã nhập kho', note: 'Demo đã nhập kho' },
    { id: 'FG-KTP-DEMO-003', finishedCode: 'TP-20260613-103', orderId: 'LSX-KTP-IMPORTED-003', orderCode: 'LSX-KTP-IMPORTED-003', productName: 'HNS 252.G1', product: 'HNS 252.G1', lot: 'LOT-TP-003', spec: 'Thùng 5 kg: 100 thùng', boxes: 100, weight: 500, importDate: '2026-06-13', location: 'KTP-B01', receiver: 'Thủ kho Lan', status: 'Đã nhập kho', note: 'Demo đã nhập kho' },
    { id: 'FG-KTP-DEMO-004', finishedCode: 'TP-20260613-104', orderId: 'LSX-KTP-IMPORTED-004', orderCode: 'LSX-KTP-IMPORTED-004', productName: 'HNS 252.G1', product: 'HNS 252.G1', lot: 'LOT-TP-004', spec: 'Thùng 25 kg: 20 thùng, Thùng 10 kg: 30 thùng', boxes: 50, weight: 800, importDate: '2026-06-13', location: 'KTP-B02', receiver: 'Thủ kho Huy', status: 'Đã nhập kho', note: 'Demo đã nhập kho' },
    { id: 'FG-KTP-DEMO-005', finishedCode: 'TP-20260613-105', orderId: 'LSX-KTP-IMPORTED-005', orderCode: 'LSX-KTP-IMPORTED-005', productName: 'HNS 252.G1', product: 'HNS 252.G1', lot: 'LOT-TP-005', spec: 'Thùng 10 kg: 40 thùng, Thùng 5 kg: 50 thùng', boxes: 90, weight: 650, importDate: '2026-06-13', location: 'KTP-C01', receiver: 'Thủ kho Minh', status: 'Đã nhập kho', note: 'Demo đã nhập kho' },
  ]
  const existingOrderIds = new Set((current.orders || []).map((order) => order.id))
  const existingPackingIds = new Set((current.packingLogs || []).map((log) => log.packingId))
  const existingFinishedIds = new Set(normalizeFinishedGoodsData(current.finishedGoods || []).map((item) => item.id))
  const orders = built.map((item) => item.order).filter((order) => !existingOrderIds.has(order.id))
  const packingLogs = built.map((item) => item.packingLog).filter((log) => !existingPackingIds.has(log.packingId))
  const finishedGoods = importedGoods.filter((item) => !existingFinishedIds.has(item.id))
  const productionLog = { id: 'PROD-DEMO-KTP-001', time: nowText(), entry: 'Đã tạo dữ liệu demo kho thành phẩm.' }
  const hasProductionLog = (current.productionLogs || current.logs || []).some((log) => log.id === productionLog.id)
  return { orders, packingLogs, finishedGoods, productionLogs: hasProductionLog ? [] : [productionLog] }
}

function buildQc2DemoFormula(quantityKg, qc1Changes = {}) {
  const source = masterFormulaLines['HNS-252-G1'].map(([code, group, percent]) => ({ code, name: code, group, percent }))
  const original = buildFormulaItems(source, quantityKg).map((item) => ({
    ...item,
    id: `DEMO-${quantityKg}-${item.materialCode}`.replaceAll(' ', '-'),
    qrStatus: 'PASS',
    weighStatus: 'PASS',
    actualWeight: item.requiredKg,
    confirmedAt: '2026-06-13 09:30',
  }))
  const active = original.map((item) => {
    const delta = qc1Changes[item.materialCode] || 0
    const requiredKg = Number((item.requiredKg + delta).toFixed(3))
    return {
      ...item,
      id: `${item.id}-QC1`,
      requiredKg,
      actualWeight: requiredKg,
      note: delta ? 'QC sản xuất thử điều chỉnh' : item.note,
    }
  })
  const changes = active
    .map((item) => {
      const originalItem = findMaterialByCode(original, item.materialCode)
      const diff = Number((item.requiredKg - num(originalItem?.requiredKg)).toFixed(3))
      return diff ? { ...item, originalKg: originalItem?.requiredKg, diff } : null
    })
    .filter(Boolean)
  return { original, active, changes }
}

function buildQc2DemoOrder(config) {
  const { original, active, changes } = buildQc2DemoFormula(config.quantityKg, config.qc1Changes)
  const baseTime = config.createdAt || '2026-06-13 08:00'
  const mixingCompletedAt = config.mixingCompletedAt || '2026-06-13 12:00'
  return {
    id: config.id,
    orderCode: config.id,
    customer: config.customer,
    product: 'HNS 252.G1',
    productName: 'HNS 252.G1',
    lot: config.lot,
    formulaId: 'HNS-252-G1',
    formulaCode: 'HNS 252 G1',
    formulaVersion: 'V3.0',
    originalFormulaId: 'HNS-252-G1',
    originalFormulaVersion: 'V3.0',
    requestedWeight: config.quantityKg,
    quantityKg: config.quantityKg,
    createdAt: baseTime,
    updatedAt: config.updatedAt || mixingCompletedAt,
    stage: config.stage || 'finished-qc',
    status: config.status || 'Chờ QC thành phẩm',
    qc1Result: changes.length ? 'QC1 đã điều chỉnh & duyệt' : 'QC1 đạt',
    qc1Status: changes.length ? 'QC1 đã điều chỉnh & duyệt' : 'QC1 đạt',
    qc1Adjustments: [{ id: `QC1-${config.id}`, time: '2026-06-13 09:15', result: changes.length ? 'Điều chỉnh' : 'Đạt', changes }],
    qc1Logs: [{ id: `QC1-${config.id}`, time: '2026-06-13 09:15', result: changes.length ? 'Điều chỉnh' : 'Đạt', changes }],
    originalFormulaSnapshot: original,
    productionFormulaSnapshot: original,
    activeProductionFormula: active,
    qc1AdjustedFormula: active,
    qc2Adjustments: config.qc2Adjustments || [],
    qc2SupplementTickets: config.qc2SupplementTickets || [],
    qc2AdjustedFormula: config.qc2SupplementTickets || [],
    qc2: config.qc2 || null,
    qc2Status: config.qc2Status || 'pending',
    chemicalStatus: 'Completed',
    solidStatus: 'Completed',
    scaleStatus: { chemical: 'Completed', solid: 'Completed' },
    ChemicalCompleted: true,
    SolidCompleted: true,
    ReadyMixing: true,
    mixingStatus: 'completed',
    mixingMachine: config.machineCode,
    mixingStartAt: config.mixingStartAt || '2026-06-13 11:00',
    mixingCompletedAt,
    mixing: {
      status: 'Completed',
      machineCode: config.machineCode,
      startedAt: config.mixingStartAt || '2026-06-13 11:00',
      completedAt: mixingCompletedAt,
      finalWeightKg: config.finalWeightKg,
      operator: 'Tổ phối trộn',
    },
    packagingStatus: config.stage === 'packaging' ? 'Pending' : 'Pending',
    finishedGoodsStatus: 'Pending',
    initialOrderSnapshot: {
      id: config.id,
      orderCode: config.id,
      productName: 'HNS 252.G1',
      lot: config.lot,
      requestedWeight: config.quantityKg,
      createdAt: baseTime,
      originalFormulaSnapshot: original,
    },
  }
}

function createQc2DemoPayload(current = {}) {
  const adjustmentId = 'QC2-ADJ-001'
  const adjustedAt = '2026-06-13 13:10'
  const adjustedItems = [
    { id: 'DEMO-QC2-ADJ-001-PASTE02', adjustmentId, orderId: 'LSX-QC2-DEMO-003', changeType: 'existing', materialCode: 'PASTE 02', materialName: 'PASTE 02', materialGroup: CHEMICAL, adjustmentKg: 1.5, requiredKg: 1.5, reason: 'Tăng độ phủ màu', note: 'Bổ sung sau QC2 lần 1' },
    { id: 'DEMO-QC2-ADJ-001-IN03', adjustmentId, orderId: 'LSX-QC2-DEMO-003', changeType: 'existing', materialCode: 'IN03', materialName: 'IN03', materialGroup: CHEMICAL, adjustmentKg: 0.2, requiredKg: 0.2, reason: 'Cân chỉnh sắc độ', note: 'Bổ sung sau QC2 lần 1' },
  ]
  const adjustmentTicket = {
    id: adjustmentId,
    adjustmentId,
    AdjustmentID: adjustmentId,
    orderId: 'LSX-QC2-DEMO-003',
    OrderID: 'LSX-QC2-DEMO-003',
    adjustmentNo: 1,
    AdjustmentNo: 1,
    qc2No: 1,
    createdBy: 'QC Demo',
    CreatedBy: 'QC Demo',
    createdAt: adjustedAt,
    CreatedAt: adjustedAt,
    reason: 'Màu nhạt sau phối trộn',
    Reason: 'Màu nhạt sau phối trộn',
    status: 'SupplementWeighed',
    Status: 'SupplementWeighed',
    qc2ResultAfterAdjustment: 'Chờ QC lại',
    qc2Record: { result: 'Cần điều chỉnh', color: 'Nhạt', ph: '7.2', viscosity: '92 KU', density: '1.31', coverage: 'Chưa đạt', fineness: 'Đạt', note: 'Cần bổ sung màu', checkedAt: adjustedAt, orderId: 'LSX-QC2-DEMO-003' },
    items: adjustedItems,
    totalSupplementKg: 1.7,
  }
  const supplementTicket = {
    id: 'BS-QC2-DEMO-003-001',
    adjustmentId,
    orderId: 'LSX-QC2-DEMO-003',
    type: 'QC2SupplementWeighing',
    label: 'Cân bổ sung QC2',
    status: 'Completed',
    createdAt: adjustedAt,
    createdBy: 'QC Demo',
    items: adjustedItems.map((item) => ({ ...item, id: `${item.id}-CAN`, toleranceKg: 0.01, qrScanned: item.materialCode, qrStatus: 'PASS', actualWeight: item.requiredKg, weighStatus: 'PASS', confirmedAt: '2026-06-13 13:35' })),
  }
  const demoOrders = [
    buildQc2DemoOrder({ id: 'LSX-QC2-DEMO-001', customer: 'Công ty Minh Long', lot: 'LOT-QC2-G1-001', quantityKg: 1000, finalWeightKg: 998.6, machineCode: 'M01', createdAt: '2026-06-13 07:30', mixingCompletedAt: '2026-06-13 11:35', qc2Status: 'waiting' }),
    buildQc2DemoOrder({ id: 'LSX-QC2-DEMO-002', customer: 'Công ty An Phú', lot: 'LOT-QC2-G1-002', quantityKg: 750, finalWeightKg: 749.2, machineCode: 'M02', createdAt: '2026-06-13 07:50', mixingCompletedAt: '2026-06-13 11:55', qc1Changes: { 'SW34': 0.5 }, qc2Status: 'pending' }),
    buildQc2DemoOrder({ id: 'LSX-QC2-DEMO-003', customer: 'Nhà máy Đông Á', lot: 'LOT-QC2-G1-003', quantityKg: 1200, finalWeightKg: 1201.1, machineCode: 'M03', createdAt: '2026-06-13 08:10', mixingCompletedAt: '2026-06-13 12:15', qc1Changes: { 'PASTE 02': 0.8, IN03: 0.05 }, qc2Status: 'waiting', qc2: adjustmentTicket.qc2Record, qc2Adjustments: [adjustmentTicket], qc2SupplementTickets: [supplementTicket] }),
    buildQc2DemoOrder({ id: 'LSX-QC2-DEMO-004', customer: 'Công ty Việt Sơn', lot: 'LOT-QC2-G1-004', quantityKg: 500, finalWeightKg: 499.7, machineCode: 'M04', createdAt: '2026-06-13 08:35', mixingCompletedAt: '2026-06-13 12:40', qc2Status: 'pending' }),
    buildQc2DemoOrder({ id: 'LSX-QC2-DEMO-005', customer: 'Công ty Nam Hải', lot: 'LOT-QC2-G1-005', quantityKg: 900, finalWeightKg: 900.4, machineCode: 'M05', createdAt: '2026-06-13 08:55', mixingCompletedAt: '2026-06-13 13:00', qc1Changes: { KT01: 0.1 }, qc2Status: 'waiting' }),
    buildQc2DemoOrder({ id: 'LSX-QC2-DEMO-006', customer: 'Công ty Bình Minh', lot: 'LOT-QC2-G1-006', quantityKg: 650, finalWeightKg: 649.8, machineCode: 'M06', createdAt: '2026-06-13 09:15', mixingCompletedAt: '2026-06-13 13:20', stage: 'packaging', status: 'Hoàn thành', qc2Status: 'Đạt', qc2: { result: 'Đạt', color: 'Đạt', ph: '7.1', viscosity: '95 KU', density: '1.30', coverage: 'Đạt', fineness: 'Đạt', note: 'Đạt QC2 lần đầu', checkedAt: '2026-06-13 13:45', orderId: 'LSX-QC2-DEMO-006' } }),
  ]
  const existingIds = new Set((current.orders || []).map((order) => order.id))
  const newOrders = demoOrders.filter((order) => !existingIds.has(order.id))
  const hasAdjustment = (current.qc2AdjustmentTickets || current.qc2Adjustments || []).some((ticket) => (ticket.adjustmentId || ticket.id) === adjustmentId)
  const hasSupplement = (current.supplementalWeighing || []).some((ticket) => ticket.id === supplementTicket.id)
  const demoLogs = [
    { id: 'QC2-DEMO-LOG-003-001', orderId: 'LSX-QC2-DEMO-003', time: adjustedAt, action: 'QC2 cần điều chỉnh', result: 'Cần điều chỉnh', adjustmentId },
    { id: 'QC2-DEMO-LOG-006-001', orderId: 'LSX-QC2-DEMO-006', time: '2026-06-13 13:45', action: 'QC2 đạt', result: 'Đạt' },
  ]
  const existingLogIds = new Set((current.qc2Logs || []).map((log) => log.id))
  const productionLogRows = [
    { id: 'PROD-DEMO-QC2-001', time: nowText(), entry: 'Đã tạo dữ liệu demo QC thành phẩm.' },
  ]
  const existingProductionLogIds = new Set((current.productionLogs || current.logs || []).map((log) => log.id))
  return {
    orders: newOrders,
    qc2AdjustmentTickets: hasAdjustment ? [] : [adjustmentTicket],
    supplementalWeighing: hasSupplement ? [] : [supplementTicket],
    qc2Logs: demoLogs.filter((log) => !existingLogIds.has(log.id)),
    productionLogs: productionLogRows.filter((log) => !existingProductionLogIds.has(log.id)),
  }
}

function PseudoQr({ value }) {
  const bits = Array.from({ length: 81 }, (_, index) => (value.charCodeAt(index % Math.max(value.length, 1)) + index * 7) % 3 === 0)
  return <div className="pseudo-qr">{bits.map((on, index) => <span className={on ? 'on' : ''} key={index} />)}</div>
}

function RawMaterialsPage({ data, setData }) {
  const importMaterialCatalogRef = useRef(null)
  const [form, setForm] = useState({ materialCode: '', materialName: '', materialGroup: CHEMICAL, lot: '', importDate: todayText(), supplier: '', weight: 0, unit: 'kg' })
  const materialCatalog = deriveMaterialCatalog(data)
  const selectedCatalogMaterial = materialCatalog.find((item) => item.materialCode === form.materialCode)
  const previewLot = normalizeRawMaterialLot({
    materialCode: form.materialCode,
    materialName: form.materialName,
    materialGroup: form.materialGroup,
    lotCode: form.lot,
    supplier: form.supplier,
    importDate: form.importDate,
    initialQty: form.weight,
    remainingQty: form.weight,
    unit: form.unit,
  })
  const qrValue = buildRawMaterialLotQr(previewLot)
  const selectMaterial = (materialCode) => {
    const material = materialCatalog.find((item) => item.materialCode === materialCode)
    setForm((current) => ({
      ...current,
      materialCode,
      materialName: material?.materialName || '',
      materialGroup: material?.materialGroup || CHEMICAL,
      unit: material?.unit || 'kg',
    }))
  }
  const save = () => {
    if (!selectedCatalogMaterial || !form.lot || !form.supplier || !form.weight || !form.importDate) return
    const item = normalizeRawMaterialLot({ ...form, id: uid('RM'), lotCode: form.lot, initialQty: form.weight, remainingQty: form.weight, qrCode: qrValue })
    setData((current) => addLogToData({
      ...current,
      materialCatalog: mergeMaterialCatalog(current.materialCatalog || [], [selectedCatalogMaterial]),
      rawMaterials: [item, ...normalizeRawMaterialLots(current.rawMaterials || [])],
    }, `Tạo QR nguyên liệu khi nhập kho: ${item.materialCode} - lô ${item.lotCode}.`))
    setForm({ materialCode: '', materialName: '', materialGroup: CHEMICAL, lot: '', importDate: todayText(), supplier: '', weight: 0, unit: 'kg' })
  }
  const printQr = () => {
    const html = `<html><head><title>QR NVL</title></head><body><h2>${form.materialCode || 'QR demo'}</h2><pre>${qrValue}</pre><script>window.print()</script></body></html>`
    const win = window.open('', '_blank')
    win.document.write(html)
    win.document.close()
  }
  const downloadMaterialCatalogTemplate = () => {
    const rows = [
      { 'Mã vật tư': 'PASTE 02', 'Tên vật tư': 'Paste nền 02', 'Nhóm vật tư': CHEMICAL, 'Đơn vị tính': 'kg' },
      { 'Mã vật tư': 'R91', 'Tên vật tư': 'Bột R91', 'Nhóm vật tư': SOLID, 'Đơn vị tính': 'kg' },
    ]
    const sheet = XLSX.utils.json_to_sheet(rows)
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, sheet, 'Danh muc vat tu')
    XLSX.writeFile(book, 'mau-danh-muc-vat-tu.xlsx')
  }
  const importMaterialCatalogExcel = (file) => {
    if (!file) return
    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: '' })
        const imported = normalizeMaterialCatalog(rows.map((row) => ({
          materialCode: readExcelCell(row, ['Mã vật tư', 'Ma vat tu', 'Mã VT', 'Ma VT']),
          materialName: readExcelCell(row, ['Tên vật tư', 'Ten vat tu', 'Tên VT', 'Ten VT']),
          materialGroup: readExcelCell(row, ['Nhóm vật tư', 'Nhom vat tu', 'Nhóm', 'Nhom']),
          unit: readExcelCell(row, ['Đơn vị tính', 'Don vi tinh', 'Đơn vị', 'Don vi']),
        })))
        if (!imported.length) return
        setData((current) => addLogToData({
          ...current,
          materialCatalog: mergeMaterialCatalog(deriveMaterialCatalog(current), imported),
        }, `Tải danh mục vật tư Excel: cập nhật ${imported.length} mã vật tư.`))
      } catch (error) {
        console.error(error)
      } finally {
        if (importMaterialCatalogRef.current) importMaterialCatalogRef.current.value = ''
      }
    }
    reader.readAsArrayBuffer(file)
  }
  return (
    <div className="page-content">
      <section className="panel">
        <div className="panel-header-row">
          <div><h2>Kho nguyên liệu</h2><p className="panel-text">Nhập NVL, sinh QR/Barcode demo, in QR và lưu localStorage.</p></div>
          <div className="action-row">
            <button className="secondary-button" onClick={() => importMaterialCatalogRef.current?.click()}>Tải danh mục vật tư Excel</button>
            <button className="secondary-button" onClick={downloadMaterialCatalogTemplate}>Tải file mẫu</button>
          </div>
        </div>
        <input ref={importMaterialCatalogRef} type="file" accept=".xlsx,.xls" hidden onChange={(event) => importMaterialCatalogExcel(event.target.files?.[0])} />
        <div className="material-entry-layout">
          <div className="material-form-area">
            <div className="material-form-grid">
              <div className="material-form-column">
                <label className="material-form-field">Mã vật tư *
                  <input list="raw-material-catalog" value={form.materialCode} onChange={(event) => selectMaterial(event.target.value)} placeholder="Tìm hoặc chọn mã vật tư" />
                  <datalist id="raw-material-catalog">{materialCatalog.map((item) => <option key={item.materialCode} value={item.materialCode}>{item.materialName}</option>)}</datalist>
                </label>
                <label className="material-form-field">Lô nhập<input value={form.lot} onChange={(event) => setForm({ ...form, lot: event.target.value })} /></label>
                <label className="material-form-field">Nhà cung cấp *<input value={form.supplier} onChange={(event) => setForm({ ...form, supplier: event.target.value })} /></label>
                <label className="material-form-field">Nhóm vật tư *<input value={form.materialGroup} readOnly /></label>
              </div>
              <div className="material-form-column">
                <label className="material-form-field">Tên vật tư *<input value={form.materialName} readOnly /></label>
                <label className="material-form-field">Khối lượng *<input type="number" value={form.weight} onChange={(event) => setForm({ ...form, weight: event.target.value })} /></label>
                <label className="material-form-field">Đơn vị tính *<input value={form.unit} readOnly /></label>
                <label className="material-form-field">Ngày nhập *<input type="date" value={form.importDate} onChange={(event) => setForm({ ...form, importDate: event.target.value })} /></label>
              </div>
            </div>
            <div className="material-form-actions">
              <button className="primary-button" onClick={save}>Lưu nhập kho</button>
              <button className="secondary-button" onClick={printQr}>In QR demo</button>
            </div>
          </div>
          <div className="qr-preview-panel material-qr-panel"><PseudoQr value={qrValue} /><code>{qrValue}</code></div>
        </div>
      </section>
      <section className="panel">
        <h3>Tồn kho nguyên liệu theo lô</h3>
        <SimpleTable headers={['Mã vật tư', 'Tên vật tư', 'Nhóm', 'Lô nhập', 'Nhà cung cấp', 'Ngày nhập', 'Tồn ban đầu', 'Đã xuất cho SX', 'Tồn còn lại', 'Đơn vị', 'QR', 'Trạng thái']} rows={normalizeRawMaterialLots(data.rawMaterials || []).map((item) => (
          <tr key={item.id}>
            <td>{item.materialCode}</td>
            <td>{item.materialName}</td>
            <td>{item.materialGroup}</td>
            <td>{item.lotCode}</td>
            <td>{item.supplier}</td>
            <td>{item.importDate}</td>
            <td>{num(item.initialQty).toLocaleString('vi-VN')}</td>
            <td>{num(item.issuedQty).toLocaleString('vi-VN')}</td>
            <td>{num(item.remainingQty).toLocaleString('vi-VN')}</td>
            <td>{item.unit}</td>
            <td><code>{item.qrCode}</code></td>
            <td><span className={`dispatch-badge ${item.status === 'Hết tồn' ? 'fail' : item.status === 'Sắp hết' ? 'waiting' : 'ready'}`}>{item.status}</span></td>
          </tr>
        ))} />
      </section>
    </div>
  )
}

const emptyFormulaLine = () => ({
  id: uid('fml'),
  materialCode: '',
  materialName: '',
  materialGroup: CHEMICAL,
  ratioPercent: 0,
})

const emptyMasterFormulaDraft = () => ({
  code: '',
  product: '',
  productGroup: '',
  version: 'V1.0',
  effectiveDate: todayText(),
  note: '',
  items: [emptyFormulaLine()],
})

const formulaTotalPercent = (items = []) => Number(items.reduce((sum, item) => sum + num(item.ratioPercent), 0).toFixed(3))
const formulaHasDuplicateMaterials = (items = []) => {
  const codes = items.map((item) => String(item.materialCode || '').trim()).filter(Boolean)
  return new Set(codes).size !== codes.length
}

const readExcelCell = (row, keys) => {
  for (const key of keys) {
    const value = row[key]
    if (value !== undefined && value !== null && String(value).trim() !== '') return value
  }
  return ''
}

function FormulasPage({ data, setData, permissions = [] }) {
  const [selectedId, setSelectedId] = useState(data.formulas[0]?.id || '')
  const [createOpen, setCreateOpen] = useState(false)
  const [formulaDraft, setFormulaDraft] = useState(emptyMasterFormulaDraft)
  const [formulaMessage, setFormulaMessage] = useState('')
  const importFormulaRef = useRef(null)
  const selected = data.formulas.find((item) => item.id === selectedId) || data.formulas[0]
  const versions = (data.formulaVersions || []).filter((version) => version.formulaId === selectedId)
  const approvedVersions = versions.filter((version) => version.status === 'Đã duyệt')
  const latestApproved = approvedVersions.at(-1)
  const [selectedVersionId, setSelectedVersionId] = useState('')
  const selectedVersion = versions.find((version) => version.id === selectedVersionId) || latestApproved || null
  const [draft, setDraft] = useState(null)
  const activeVersion = draft || selectedVersion
  const comparisonItems = (selected?.items || []).map((baseItem) => {
    const adjusted = activeVersion?.items?.find((item) => item.baseItemId === baseItem.id || item.materialCode === baseItem.materialCode)
    const adjustedPercent = adjusted?.adjustedPercent ?? baseItem.ratioPercent
    const diff = Number((num(adjustedPercent) - num(baseItem.ratioPercent)).toFixed(3))
    return { ...baseItem, adjustedPercent, diff, adjustmentNote: adjusted?.adjustmentNote || '' }
  })
  const adjustedTotal = Number(comparisonItems.reduce((sum, item) => sum + num(item.adjustedPercent), 0).toFixed(3))
  const canApprove = adjustedTotal === 100
  const draftTotal = formulaTotalPercent(formulaDraft.items)
  const draftDuplicateMaterials = formulaHasDuplicateMaterials(formulaDraft.items)
  const draftCodeExists = (data.formulas || []).some((formula) => formula.code?.trim().toLowerCase() === formulaDraft.code.trim().toLowerCase() || formula.id?.trim().toLowerCase() === formulaDraft.code.trim().toLowerCase())
  const canSaveFormula = formulaDraft.code.trim()
    && formulaDraft.product.trim()
    && formulaDraft.items.length > 0
    && formulaDraft.items.every((item) => item.materialCode.trim() && item.materialName.trim() && item.materialGroup.trim())
    && draftTotal === 100
    && !draftDuplicateMaterials
    && !draftCodeExists
  const canCreateFormula = hasPermission(permissions, 'master.formula.create')
  const canEditFormula = hasPermission(permissions, 'master.formula.edit')
  const canSecureViewFormula = hasPermission(permissions, 'formula.secure.view')

  const openCreateFormula = () => {
    if (!canCreateFormula) return
    setFormulaDraft(emptyMasterFormulaDraft())
    setFormulaMessage('')
    setCreateOpen(true)
  }

  const updateFormulaDraft = (field, value) => setFormulaDraft((current) => ({ ...current, [field]: value }))
  const updateFormulaLine = (lineId, field, value) => setFormulaDraft((current) => ({
    ...current,
    items: current.items.map((item) => item.id === lineId ? { ...item, [field]: field === 'ratioPercent' ? num(value) : value } : item),
  }))
  const addFormulaLine = () => setFormulaDraft((current) => ({ ...current, items: [...current.items, emptyFormulaLine()] }))
  const removeFormulaLine = (lineId) => setFormulaDraft((current) => ({
    ...current,
    items: current.items.length > 1 ? current.items.filter((item) => item.id !== lineId) : current.items,
  }))

  const saveMasterFormula = () => {
    if (!canCreateFormula || !canSaveFormula) return
    const formula = {
      id: formulaDraft.code.trim(),
      code: formulaDraft.code.trim(),
      product: formulaDraft.product.trim(),
      productGroup: formulaDraft.productGroup.trim(),
      version: formulaDraft.version.trim() || 'V1.0',
      effectiveDate: formulaDraft.effectiveDate || todayText(),
      note: formulaDraft.note.trim(),
      createdBy: 'Phòng kỹ thuật',
      checkedBy: '',
      approvedBy: '',
      createdAt: nowText(),
      items: formulaDraft.items.map((item, index) => ({
        id: uid('fml'),
        no: index + 1,
        materialCode: item.materialCode.trim(),
        materialName: item.materialName.trim(),
        materialGroup: item.materialGroup.trim(),
        ratioPercent: num(item.ratioPercent),
      })),
    }
    setData((current) => addLogToData({ ...current, formulas: [formula, ...(current.formulas || [])] }, `Tạo công thức gốc ${formula.code}.`))
    setSelectedId(formula.id)
    setSelectedVersionId('')
    setDraft(null)
    setCreateOpen(false)
    setFormulaMessage(`Đã tạo công thức ${formula.code}.`)
  }

  const downloadFormulaTemplate = () => {
    if (!canCreateFormula) return
    const rows = [
      { 'Mã CT': 'HNS-NEW-001', 'Tên SP': 'HNS Demo', 'Nhóm sản phẩm': 'Sơn', Version: 'V1.0', 'Ngày hiệu lực': todayText(), 'Ghi chú': '', 'Mã VT': 'PASTE 02', 'Tên VT': 'Paste nền 02', 'Nhóm': CHEMICAL, 'Tỷ lệ %': 4.61 },
      { 'Mã CT': 'HNS-NEW-001', 'Tên SP': 'HNS Demo', 'Nhóm sản phẩm': 'Sơn', Version: 'V1.0', 'Ngày hiệu lực': todayText(), 'Ghi chú': '', 'Mã VT': 'R91', 'Tên VT': 'Bột R91', 'Nhóm': SOLID, 'Tỷ lệ %': 95.39 },
    ]
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, XLSX.utils.json_to_sheet(rows), 'Cong thuc goc')
    XLSX.writeFile(book, 'mau-cong-thuc-goc.xlsx')
  }

  const importFormulaExcel = (file) => {
    if (!canCreateFormula || !file) return
    const reader = new FileReader()
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array' })
        const sheet = workbook.Sheets[workbook.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(sheet)
        const grouped = rows.reduce((acc, row) => {
          const code = String(readExcelCell(row, ['Mã CT', 'Ma CT', 'Mã công thức', 'Ma cong thuc'])).trim()
          if (!code) return acc
          if (!acc[code]) {
            acc[code] = {
              code,
              product: String(readExcelCell(row, ['Tên SP', 'Ten SP', 'Tên sản phẩm', 'Ten san pham'])).trim(),
              productGroup: String(readExcelCell(row, ['Nhóm sản phẩm', 'Nhom san pham'])).trim(),
              version: String(readExcelCell(row, ['Version', 'Phiên bản', 'Phien ban'])).trim() || 'V1.0',
              effectiveDate: String(readExcelCell(row, ['Ngày hiệu lực', 'Ngay hieu luc'])).trim() || todayText(),
              note: String(readExcelCell(row, ['Ghi chú', 'Ghi chu'])).trim(),
              items: [],
            }
          }
          acc[code].items.push({
            id: uid('fml'),
            materialCode: String(readExcelCell(row, ['Mã VT', 'Ma VT', 'Mã vật tư', 'Ma vat tu'])).trim(),
            materialName: String(readExcelCell(row, ['Tên VT', 'Ten VT', 'Tên vật tư', 'Ten vat tu'])).trim(),
            materialGroup: String(readExcelCell(row, ['Nhóm', 'Nhom'])).trim(),
            ratioPercent: num(readExcelCell(row, ['Tỷ lệ %', 'Ty le %', 'Tỉ lệ %', 'Ti le %'])),
          })
          return acc
        }, {})
        const existingCodes = new Set((data.formulas || []).flatMap((formula) => [formula.code, formula.id].filter(Boolean).map((value) => value.trim().toLowerCase())))
        const imported = Object.values(grouped)
        const errors = []
        imported.forEach((formula) => {
          if (existingCodes.has(formula.code.toLowerCase())) errors.push(`Trùng mã công thức ${formula.code}`)
          if (!formula.product) errors.push(`Thiếu Tên SP cho ${formula.code}`)
          if (formulaTotalPercent(formula.items) !== 100) errors.push(`Tổng tỷ lệ của ${formula.code} chưa bằng 100% (${formulaTotalPercent(formula.items)}%)`)
          if (formulaHasDuplicateMaterials(formula.items)) errors.push(`Trùng vật tư trong ${formula.code}`)
          if (formula.items.some((item) => !item.materialCode || !item.materialName || !item.materialGroup)) errors.push(`Thiếu thông tin vật tư trong ${formula.code}`)
        })
        if (!imported.length) errors.push('File không có dòng công thức hợp lệ.')
        if (errors.length) {
          setFormulaMessage(errors.join(' | '))
          return
        }
        const formulas = imported.map((formula) => ({
          id: formula.code,
          code: formula.code,
          product: formula.product,
          productGroup: formula.productGroup,
          version: formula.version,
          effectiveDate: formula.effectiveDate,
          note: formula.note,
          createdBy: 'Phòng kỹ thuật',
          checkedBy: '',
          approvedBy: '',
          createdAt: nowText(),
          items: formula.items.map((item, index) => ({ ...item, id: uid('fml'), no: index + 1 })),
        }))
        setData((current) => addLogToData({ ...current, formulas: [...formulas, ...(current.formulas || [])] }, `Tải ${formulas.length} công thức gốc từ Excel.`))
        setSelectedId(formulas[0].id)
        setSelectedVersionId('')
        setDraft(null)
        setFormulaMessage(`Đã tải ${formulas.length} công thức từ Excel.`)
      } catch (error) {
        setFormulaMessage(`Không đọc được file Excel: ${error.message}`)
      } finally {
        if (importFormulaRef.current) importFormulaRef.current.value = ''
      }
    }
    reader.readAsArrayBuffer(file)
  }

  const startAdjustment = () => {
    if (!canEditFormula || !selected) return
    const versionNo = versions.length + 1
    const newDraft = {
      id: uid('FVER'),
      formulaId: selected.id,
      baseVersion: selected.version,
      adjustedVersion: `${selected.version}-DC${versionNo}`,
      effectiveDate: selected.effectiveDate,
      createdBy: selected.createdBy,
      checkedBy: selected.checkedBy,
      approvedBy: selected.approvedBy,
      adjustmentReason: '',
      status: 'Nháp',
      createdAt: nowText(),
      items: selected.items.map((item) => ({
        id: uid('fver-item'),
        baseItemId: item.id,
        materialCode: item.materialCode,
        materialName: item.materialName,
        materialGroup: item.materialGroup,
        originalPercent: item.ratioPercent,
        adjustedPercent: item.ratioPercent,
        adjustmentNote: '',
      })),
    }
    setDraft(newDraft)
    setSelectedVersionId('')
    setData((current) => addLogToData(current, `Tạo phiên bản điều chỉnh ${newDraft.adjustedVersion} từ công thức gốc ${selected.code}.`))
  }

  const updateDraftField = (field, value) => setDraft((current) => ({ ...current, [field]: value }))
  const updateDraftItem = (baseItemId, field, value) => setDraft((current) => ({
    ...current,
    items: current.items.map((item) => item.baseItemId === baseItemId ? { ...item, [field]: field === 'adjustedPercent' ? num(value) : value } : item),
  }))

  const saveDraft = () => {
    if (!canEditFormula || !draft || !canApprove) return
    const approved = { ...draft, status: 'Đã duyệt', approvedAt: nowText() }
    setData((current) => {
      const exists = (current.formulaVersions || []).some((version) => version.id === approved.id)
      const formulaVersions = exists
        ? current.formulaVersions.map((version) => version.id === approved.id ? approved : version)
        : [...(current.formulaVersions || []), approved]
      return addLogToData({ ...current, formulaVersions }, `Duyệt phiên bản điều chỉnh ${approved.adjustedVersion} của công thức ${selected.code}.`)
    })
    setSelectedVersionId(approved.id)
    setDraft(null)
  }

  const restoreBase = () => {
    setDraft(null)
    setSelectedVersionId('')
    setData((current) => addLogToData(current, `Khôi phục xem công thức gốc ${selected.code}.`))
  }

  const diffClass = (diff) => (diff > 0 ? 'diff-up' : diff < 0 ? 'diff-down' : 'diff-same')
  if (!selected) return null
  return (
    <div className="page-content">
      <section className="panel">
        <div className="panel-header-row">
          <div><h2>Công thức gốc</h2><p className="panel-text">Thư viện công thức chuẩn do Phòng Kỹ thuật ban hành. Tỷ lệ gốc không bị ghi đè; mọi thay đổi được lưu thành phiên bản điều chỉnh riêng để so sánh.</p></div>
          <div className="action-row">
            <button className="primary-button" disabled={!canCreateFormula} onClick={openCreateFormula}>Tạo công thức gốc</button>
            <button className="secondary-button" disabled={!canCreateFormula} onClick={() => importFormulaRef.current?.click()}>Tải công thức Excel</button>
            <button className="secondary-button" disabled={!canCreateFormula} onClick={downloadFormulaTemplate}>Tải file mẫu</button>
            <button className="primary-button" disabled={!canEditFormula} onClick={startAdjustment}>Tạo phiên bản điều chỉnh</button>
            <button className="secondary-button" onClick={() => setSelectedVersionId(latestApproved?.id || '')}>So sánh với công thức gốc</button>
            <button className="secondary-button" onClick={restoreBase}>Khôi phục về công thức gốc</button>
          </div>
        </div>
        <input ref={importFormulaRef} type="file" accept=".xlsx,.xls" hidden onChange={(event) => importFormulaExcel(event.target.files?.[0])} />
        {formulaMessage && <div className={formulaMessage.startsWith('Đã') ? 'formula-ratio-ok' : 'formula-ratio-alert'}>{formulaMessage}</div>}
        <div className="log-tabs">{data.formulas.map((formula) => <button className={formula.id === selectedId ? 'active' : ''} key={formula.id} onClick={() => { setSelectedId(formula.id); setSelectedVersionId(''); setDraft(null) }}>{formula.code}</button>)}</div>
      </section>
      <section className="panel">
        <div className="production-form-grid">
          <label>Mã công thức<input readOnly value={selected.code} /></label>
          <label>Sản phẩm<input readOnly value={selected.product} /></label>
          <label>Version gốc<input readOnly value={selected.version} /></label>
          <label>Version điều chỉnh<select value={draft ? draft.id : selectedVersion?.id || ''} onChange={(event) => { setSelectedVersionId(event.target.value); setDraft(null) }}><option value="">Công thức gốc</option>{versions.map((version) => <option key={version.id} value={version.id}>{version.adjustedVersion} - {version.status}</option>)}</select></label>
          <label>Ngày hiệu lực<input type="date" readOnly={!draft} value={activeVersion?.effectiveDate || selected.effectiveDate} onChange={(event) => updateDraftField('effectiveDate', event.target.value)} /></label>
          <label>Người lập<input readOnly={!draft} value={activeVersion?.createdBy || selected.createdBy} onChange={(event) => updateDraftField('createdBy', event.target.value)} /></label>
          <label>Người kiểm tra<input readOnly={!draft} value={activeVersion?.checkedBy || selected.checkedBy} onChange={(event) => updateDraftField('checkedBy', event.target.value)} /></label>
          <label>Người duyệt<input readOnly={!draft} value={activeVersion?.approvedBy || selected.approvedBy} onChange={(event) => updateDraftField('approvedBy', event.target.value)} /></label>
          <label className="wide-field">Lý do điều chỉnh<input readOnly={!draft} value={activeVersion?.adjustmentReason || ''} onChange={(event) => updateDraftField('adjustmentReason', event.target.value)} /></label>
        </div>
        {canSecureViewFormula && adjustedTotal !== 100 && <div className="formula-ratio-alert">Tổng tỷ lệ điều chỉnh chưa bằng 100% ({adjustedTotal}%)</div>}
        {canSecureViewFormula && adjustedTotal === 100 && <div className="formula-ratio-ok">Tổng tỷ lệ điều chỉnh = 100%</div>}
        {!canSecureViewFormula && <div className="process-alert">Bạn có quyền xem công thức, nhưng chưa có quyền formula.secure.view nên tỷ lệ phần trăm được ẩn.</div>}
        <SimpleTable headers={canSecureViewFormula ? ['STT', 'Mã vật tư', 'Tên vật tư', 'Nhóm', 'Tỷ lệ gốc %', 'Tỷ lệ điều chỉnh %', 'Chênh lệch %', 'Ghi chú điều chỉnh'] : ['STT', 'Mã vật tư', 'Tên vật tư', 'Nhóm', 'Ghi chú điều chỉnh']} rows={comparisonItems.map((item, index) => (
          <tr key={item.id}>
            <td>{index + 1}</td>
            <td>{item.materialCode}</td>
            <td>{item.materialName}</td>
            <td>{item.materialGroup}</td>
            {canSecureViewFormula && <td>{item.ratioPercent}%</td>}
            {canSecureViewFormula && <td>{draft ? <input type="number" value={item.adjustedPercent} onChange={(event) => updateDraftItem(item.id, 'adjustedPercent', event.target.value)} /> : `${item.adjustedPercent}%`}</td>}
            {canSecureViewFormula && <td><span className={diffClass(item.diff)}>{item.diff > 0 ? '+' : ''}{item.diff}%</span></td>}
            <td>{draft ? <input value={item.adjustmentNote} onChange={(event) => updateDraftItem(item.id, 'adjustmentNote', event.target.value)} /> : item.adjustmentNote || '-'}</td>
          </tr>
        ))} />
        {draft && <div className="modal-actions"><button className="secondary-button" onClick={() => setDraft(null)}>Hủy</button><button className="primary-button" disabled={!canApprove} onClick={saveDraft}>Lưu phiên bản điều chỉnh</button></div>}
      </section>
      {createOpen && (
        <div className="modal-backdrop" role="presentation">
          <div className="production-modal formula-create-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><h2>Tạo công thức gốc</h2><p className="panel-text">Nhập thông tin công thức và tỷ lệ vật tư. Chỉ lưu khi tổng tỷ lệ bằng 100%.</p></div>
              <button className="secondary-button" onClick={() => setCreateOpen(false)}>Đóng</button>
            </div>
            <div className="production-form-grid">
              <label>Mã công thức<input value={formulaDraft.code} onChange={(event) => updateFormulaDraft('code', event.target.value)} /></label>
              <label>Tên sản phẩm<input value={formulaDraft.product} onChange={(event) => updateFormulaDraft('product', event.target.value)} /></label>
              <label>Nhóm sản phẩm<input value={formulaDraft.productGroup} onChange={(event) => updateFormulaDraft('productGroup', event.target.value)} /></label>
              <label>Version<input value={formulaDraft.version} onChange={(event) => updateFormulaDraft('version', event.target.value)} /></label>
              <label>Ngày hiệu lực<input type="date" value={formulaDraft.effectiveDate} onChange={(event) => updateFormulaDraft('effectiveDate', event.target.value)} /></label>
              <label className="wide-field">Ghi chú<input value={formulaDraft.note} onChange={(event) => updateFormulaDraft('note', event.target.value)} /></label>
            </div>
            <div className={draftTotal === 100 && !draftDuplicateMaterials && !draftCodeExists ? 'formula-ratio-ok' : 'formula-ratio-alert'}>
              Tổng tỷ lệ: {draftTotal}%{draftCodeExists ? ' - Trùng mã công thức' : ''}{draftDuplicateMaterials ? ' - Trùng vật tư' : ''}
            </div>
            <SimpleTable headers={['Mã vật tư', 'Tên vật tư', 'Nhóm', 'Tỷ lệ %', '']} rows={formulaDraft.items.map((item) => (
              <tr key={item.id}>
                <td><input value={item.materialCode} onChange={(event) => updateFormulaLine(item.id, 'materialCode', event.target.value)} /></td>
                <td><input value={item.materialName} onChange={(event) => updateFormulaLine(item.id, 'materialName', event.target.value)} /></td>
                <td><input value={item.materialGroup} onChange={(event) => updateFormulaLine(item.id, 'materialGroup', event.target.value)} /></td>
                <td><input type="number" value={item.ratioPercent} onChange={(event) => updateFormulaLine(item.id, 'ratioPercent', event.target.value)} /></td>
                <td><button className="danger-button" onClick={() => removeFormulaLine(item.id)} disabled={formulaDraft.items.length === 1}>Xóa</button></td>
              </tr>
            ))} />
            <div className="modal-actions">
              <button className="secondary-button" onClick={addFormulaLine}>Thêm dòng</button>
              <button className="secondary-button" onClick={() => setCreateOpen(false)}>Hủy</button>
              <button className="primary-button" disabled={!canSaveFormula} onClick={saveMasterFormula}>Lưu công thức</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function OrdersPage({ data, setData, permissions = [] }) {
  const [form, setForm] = useState({ formulaId: data.formulas[0]?.id || '', quantityKg: 1000, lot: '', customer: '', mixerMachine: '', productionRequestNo: '', note: '' })
  const [message, setMessage] = useState('')
  const [warning, setWarning] = useState('')
  const [detailOrderId, setDetailOrderId] = useState('')
  const [detailTab, setDetailTab] = useState('info')
  const machines = getActiveMixingMachines(normalizeMixingMachines(data.mixingMachines))
  const formula = data.formulas.find((item) => item.id === form.formulaId)
  const libraryItems = formula ? formula.items.map((item) => ({
    code: item.materialCode,
    name: item.materialName,
    group: item.materialGroup,
    percent: item.ratioPercent,
  })) : []
  const preview = formula ? buildFormulaItems(libraryItems, num(form.quantityKg)) : []
  const nextOrderCode = () => {
    const date = todayText().replaceAll('-', '')
    const countToday = data.orders.filter((order) => String(order.orderCode || order.id || '').includes(`LSX-${date}`)).length + 1
    return `LSX-${date}-${String(countToday).padStart(3, '0')}`
  }
  const create = () => {
    setMessage('')
    setWarning('')
    if (!formula) {
      setWarning('Chưa chọn công thức')
      return
    }
    if (!Number.isFinite(Number(form.quantityKg)) || Number(form.quantityKg) <= 0) {
      setWarning('Khối lượng đơn hàng không hợp lệ')
      return
    }
    const assignedMachine = machines.find((machine) => machine.machineCode === form.mixerMachine)
    if (!assignedMachine) {
      setWarning('Vui lòng chọn máy phối trộn.')
      return
    }
    const sourceLabel = `công thức gốc ${formula.version}`
    const id = nextOrderCode()
    const createdAt = nowText()
    const originalFormulaSnapshot = preview.map((item) => ({ ...item }))
    const order = {
      id,
      orderCode: id,
      formulaId: formula.id,
      formulaCode: formula.code,
      formulaVersion: formula.version,
      productName: formula.product,
      product: formula.product,
      lot: form.lot || `LOT-${id}`,
      customer: form.customer,
      mixerMachine: form.mixerMachine,
      assignedMachineCode: assignedMachine.machineCode,
      assignedMachineName: assignedMachine.machineName,
      assignedMachineCapacityKg: assignedMachine.capacityKg,
      assignedMachineMotorPower: assignedMachine.motorPower,
      assignedMachineDepartment: assignedMachine.department || assignedMachine.productionTeam,
      machineChangeHistory: [],
      machineAssignmentHistory: [{
        id: uid('MCH'),
        orderId: id,
        orderCode: id,
        lot: form.lot || `LOT-${id}`,
        assignedMachine: assignedMachine.machineCode,
        performedMachine: '',
        changedBy: 'Phòng SX',
        changedAt: createdAt,
        reason: 'Chỉ định máy khi tạo lệnh sản xuất',
      }],
      productionRequestNo: form.productionRequestNo,
      note: form.note,
      requestedWeight: num(form.quantityKg),
      quantityKg: num(form.quantityKg),
      stage: 'qc1',
      status: 'Chờ QC sản xuất thử',
      createdAt,
      updatedAt: createdAt,
      originalFormulaId: formula.id,
      originalFormulaVersion: formula.version,
      formulaSource: 'base',
      formulaVersionId: '',
      originalFormula: formula.items,
      originalFormulaSnapshot,
      activeProductionFormula: preview.map((item) => ({ ...item })),
      productionFormulaSnapshot: originalFormulaSnapshot,
      qc1Adjustments: [],
      qc2Adjustments: [],
      qc1AdjustedFormula: null,
      qc2AdjustedFormula: [],
      chemicalStatus: 'Pending',
      solidStatus: 'Pending',
      mixingStatus: 'Pending',
      mixingMachine: '',
      mixingStartAt: '',
      mixingCompletedAt: '',
      qc1Status: 'Chờ QC sản xuất thử',
      qc2Status: 'Pending',
      packagingStatus: 'Pending',
      finishedGoodsStatus: 'Pending',
      scaleStatus: { chemical: 'Pending', solid: 'Pending' },
    }
    setData((current) => addLogToData({ ...current, orders: [order, ...current.orders] }, `Sử dụng ${sourceLabel} của ${formula.code} để tạo lệnh SX ${id}, chỉ định máy ${formatMixingMachineLabel(assignedMachine)}.`))
    setMessage('Tạo lệnh sản xuất thành công')
    setForm((current) => ({ ...current, lot: '', customer: '', mixerMachine: '', productionRequestNo: '', note: '' }))
  }
  const detailOrder = data.orders.find((order) => order.id === detailOrderId)
  return (
    <div className="page-content">
      <section className="panel">
        <div className="panel-header-row"><div><h2>Lệnh sản xuất</h2><p className="panel-text">Phòng SX tạo lệnh từ công thức gốc. Hệ thống tự quy đổi khối lượng từng nguyên liệu theo khối lượng yêu cầu.</p></div></div>
        {warning && <div className="process-alert">{warning}</div>}
        {message && <div className="formula-ratio-ok">{message}</div>}
        <div className="production-form-grid order-create-form">
          <label>Công thức gốc<select value={form.formulaId} onChange={(event) => setForm({ ...form, formulaId: event.target.value })}>{data.formulas.map((item) => <option value={item.id} key={item.id}>{item.code} - {item.version}</option>)}</select></label>
          <label>Khối lượng đơn hàng<input type="number" value={form.quantityKg} onChange={(event) => setForm({ ...form, quantityKg: event.target.value })} /></label>
          <label>Mã LOT<input value={form.lot} onChange={(event) => setForm({ ...form, lot: event.target.value })} /></label>
          <label>Khách hàng<input value={form.customer} onChange={(event) => setForm({ ...form, customer: event.target.value })} /></label>
          <label>Máy phối trộn *<select value={form.mixerMachine} onChange={(event) => setForm({ ...form, mixerMachine: event.target.value })}><option value="">Chọn máy</option>{machines.map((machine) => <option key={machine.machineCode} value={machine.machineCode}>{mixingMachineOptionLabel(machine)}</option>)}</select></label>
          <label>Phiếu yêu cầu SX<input value={form.productionRequestNo} placeholder="Nhập số phiếu yêu cầu sản xuất" onChange={(event) => setForm({ ...form, productionRequestNo: event.target.value })} /></label>
          <label className="wide-field">Ghi chú<textarea value={form.note} placeholder="Ví dụ: Giống mẫu đã duyệt ngày .../..." onChange={(event) => setForm({ ...form, note: event.target.value })} /></label>
        </div>
        <button className="primary-button" onClick={create}>Tạo lệnh sản xuất</button>
      </section>
      <section className="panel">
        <h3>Danh sách lệnh</h3>
        <SimpleTable tableClassName="orders-table" headers={['Mã lệnh SX', 'Sản phẩm', 'Công thức gốc', 'Version', 'LOT', 'Khối lượng', 'Máy phối trộn', 'Phiếu yêu cầu SX', 'Ghi chú', 'Trạng thái', 'Ngày tạo', 'Hành động']} rows={data.orders.map((order) => (
          <tr key={order.id}><td>{order.orderCode || order.id}</td><td>{order.productName || order.product}</td><td>{order.formulaCode || order.originalFormulaId}</td><td>{order.formulaVersion || order.originalFormulaVersion}</td><td>{order.lot}</td><td>{kg(order.requestedWeight ?? order.quantityKg)}</td><td>{getOrderAssignedMachineLabel(order, machines)}</td><td>{order.productionRequestNo || '-'}</td><td className="ellipsis-cell" title={order.note || ''}>{order.note || '-'}</td><td><span className={`flow-pill ${statusClass(order.status)}`}>{displayQcTrialText(order.status)}</span></td><td>{order.createdAt}</td><td><button className="secondary-button" onClick={() => { setDetailOrderId(order.id); setDetailTab('info') }}>Chi tiết</button></td></tr>
        ))} />
      </section>
      {detailOrder && (
        <div className="modal-backdrop" role="presentation">
          <div className="production-modal order-detail-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><span className="section-kicker">Chi tiết lệnh sản xuất</span><h2>{detailOrder.orderCode || detailOrder.id}</h2></div>
              <button type="button" className="icon-button" onClick={() => setDetailOrderId('')} aria-label="Đóng">×</button>
            </div>
            <div className="log-tabs">
              {[
                ['info', 'Thông tin lệnh'],
                ['materials', 'Định mức nguyên liệu'],
                ['qc1', 'Điều chỉnh QC sản xuất thử'],
                ['qc2', 'Điều chỉnh QC2'],
                ['chemical', 'Nhật ký cân hóa'],
                ['solid', 'Nhật ký cân rắn'],
                ['mixing', 'Nhật ký phối trộn'],
                ['finishedQc', 'Nhật ký QC thành phẩm'],
              ].map(([id, label]) => <button key={id} className={detailTab === id ? 'active' : ''} onClick={() => setDetailTab(id)}>{label}</button>)}
            </div>
            <OrderDetailTabs order={detailOrder} tab={detailTab} productionLogs={data.productionLogs || []} qc2Logs={data.qc2Logs || []} permissions={permissions} />
          </div>
        </div>
      )}
    </div>
  )
}

function FormulaTable({ items, secure = false }) {
  const headers = secure ? ['STT', 'Mã VT', 'Tên vật tư', 'Nhóm', 'Tỷ lệ %', 'Khối lượng/lệnh'] : ['STT', 'Mã VT', 'Tên vật tư', 'Nhóm', 'Khối lượng/lệnh']
  return <SimpleTable headers={headers} rows={items.map((item, index) => (
    <tr key={item.id || index}>
      <td>{index + 1}</td>
      <td>{item.materialCode}</td>
      <td>{item.materialName}</td>
      <td>{item.materialGroup}</td>
      {secure && <td>{item.ratioPercent}%</td>}
      <td>{kg(item.requiredKg)}</td>
    </tr>
  ))} />
}

function OrderDetailTabs({ order, tab, productionLogs, qc2Logs, permissions = [] }) {
  const materials = order.activeProductionFormula || order.productionFormulaSnapshot || []
  const originalMaterials = order.originalFormulaSnapshot || order.productionFormulaSnapshot || []
  const canViewFormula = hasAnyPermission(permissions, ['master.formula.view', 'formula.secure.view'])
  const canSecureViewFormula = hasPermission(permissions, 'formula.secure.view')
  const qc1Rows = order.qc1Adjustments || order.qc1Logs || []
  const qc2Rows = order.qc2Adjustments || order.qc2AdjustedFormula || []
  const relatedLogs = productionLogs.filter((log) => String(log.entry || '').includes(order.orderCode || order.id))
  const groupLogs = (group) => materials.filter((item) => item.materialGroup === group)

  if (tab === 'info') {
    return (
      <div className="production-log-grid">
        {[
          ['Mã lệnh SX', order.orderCode || order.id],
          ['Sản phẩm', order.productName || order.product],
          ['Khách hàng', order.customer || '-'],
          ['Công thức gốc', order.formulaCode || order.originalFormulaId],
          ['Version', order.formulaVersion || order.originalFormulaVersion],
          ['LOT', order.lot],
          ['Khối lượng yêu cầu', kg(order.requestedWeight ?? order.quantityKg)],
          ['Trạng thái', displayQcTrialText(order.status)],
          ['Ngày tạo', order.createdAt],
          ['Cân hóa', order.chemicalStatus || order.scaleStatus?.chemical || 'Pending'],
          ['Cân rắn', order.solidStatus || order.scaleStatus?.solid || 'Pending'],
          ['Phối trộn', order.mixingStatus || order.mixing?.status || 'Pending'],
        ].map(([label, value]) => <div key={label}><span>{label}</span><strong>{value}</strong></div>)}
      </div>
    )
  }

  if (tab === 'materials') {
    if (!canViewFormula) return <p className="empty-alert">Bạn chưa có quyền xem dữ liệu công thức của lệnh sản xuất.</p>
    return <FormulaTable secure={canSecureViewFormula} items={originalMaterials.map((item) => ({ ...item, requiredKg: item.requiredKg ?? item.materialPerLot }))} />
  }
  if (tab === 'qc1') {
    return qc1Rows.length ? <SimpleTable headers={['Thời gian', 'Kết quả', 'Số dòng chỉnh']} rows={qc1Rows.map((row) => <tr key={row.id || row.time}><td>{row.time}</td><td>{displayQcTrialText(row.result)}</td><td>{row.changes?.length || 0}</td></tr>)} /> : <p className="empty-alert">Chưa có điều chỉnh QC sản xuất thử.</p>
  }
  if (tab === 'qc2') {
    return qc2Rows.length ? <SimpleTable headers={['Phiếu', 'Trạng thái', 'Người điều chỉnh', 'Nội dung điều chỉnh']} rows={qc2Rows.map((ticket) => <tr key={ticket.id}><td>{ticket.adjustmentId || ticket.id}</td><td>{ticket.status}</td><td>{ticket.createdBy || '-'}</td><td>{getAdjustmentItems(ticket).map((item) => `${item.materialCode}: ${num(item.adjustmentKg ?? item.requiredKg) > 0 ? '+' : ''}${kg(item.adjustmentKg ?? item.requiredKg)}`).join(', ')}</td></tr>)} /> : <p className="empty-alert">Chưa có điều chỉnh QC2.</p>
  }
  if (tab === 'chemical' || tab === 'solid') {
    const rows = groupLogs(tab === 'chemical' ? CHEMICAL : SOLID)
    return <SimpleTable headers={['Mã vật tư', 'Cần cân', 'QR', 'Thực cân', 'Trạng thái', 'Thời gian']} rows={rows.map((item) => <tr key={item.id}><td>{item.materialCode}</td><td>{kg(item.requiredKg)}</td><td>{item.qrStatus || '-'}</td><td>{item.actualWeight === '' ? '-' : kg(item.actualWeight)}</td><td>{item.weighStatus || '-'}</td><td>{item.confirmedAt || '-'}</td></tr>)} />
  }
  if (tab === 'mixing') {
    return <SimpleTable headers={['Thời gian bắt đầu', 'Thời gian hoàn thành', 'Trạng thái', 'Người thực hiện', 'Khối lượng sau phối trộn']} rows={[<tr key="mixing"><td>{order.mixing?.startedAt || '-'}</td><td>{order.mixing?.completedAt || '-'}</td><td>{order.mixing?.status || '-'}</td><td>{order.mixing?.operator || '-'}</td><td>{order.mixing?.finalWeightKg ? kg(order.mixing.finalWeightKg) : '-'}</td></tr>]} />
  }
  if (tab === 'finishedQc') {
    const rows = qc2Logs.filter((log) => log.orderId === order.id)
    return rows.length ? <SimpleTable headers={['Thời gian', 'Hành động', 'Kết quả']} rows={rows.map((log) => <tr key={log.id}><td>{log.time}</td><td>{log.action}</td><td>{log.result || '-'}</td></tr>)} /> : <p className="empty-alert">Chưa có nhật ký QC thành phẩm.</p>
  }
  return <SimpleTable headers={['Thời gian', 'Nội dung']} rows={relatedLogs.map((log) => <tr key={log.id}><td>{log.time}</td><td>{log.entry}</td></tr>)} />
}

function QCPage({ data, setData, user }) {
  return (
    <QC1 data={data} setData={setData} user={user} />
  )
}

function QC1({ data, setData, user }) {
  const waitingOrders = data.orders.filter((order) => order.stage === 'qc1')
  const [activeOrderId, setActiveOrderId] = useState(waitingOrders[0]?.id || '')
  const [showAddMaterial, setShowAddMaterial] = useState(false)
  const [newMaterial, setNewMaterial] = useState({ materialCode: '', materialName: '', materialGroup: CHEMICAL, requiredKg: 0, reason: '', note: '' })
  const activeOrder = data.orders.find((order) => order.id === activeOrderId && order.stage === 'qc1') || waitingOrders[0]
  const completedOrders = data.orders.filter((order) => order.qc1Result && order.stage !== 'qc1')

  const currentAssignments = getActiveAssignments(data.productionAssignments || [], 'QC')
  const assignmentEmployeeText = getAssignmentLogContext(currentAssignments).employee
  const patchOrder = (orderId, updater, log, result = 'Cập nhật QC') => setData((current) => addLogToData(
    { ...current, orders: current.orders.map((order) => order.id === orderId ? { ...updater(order), updatedAt: nowText() } : order) },
    log,
    operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'QC', order: orderId, result }),
  ))
  const updateItem = (orderId, itemId, field, value) => {
    patchOrder(orderId, (order) => ({
      ...order,
      activeProductionFormula: getEffectiveFormula(order).map((item) => item.id === itemId ? { ...item, [field]: value } : item),
    }), `QC1 ghi nhận thao tác trên ${orderId} - ${itemId}.`, 'Cập nhật chỉ tiêu QC')
  }
  const saveNewMaterial = () => {
    if (!activeOrder || !newMaterial.materialCode) return
    const addedAt = nowText()
    const material = {
      id: uid('qc1-new'),
      no: getEffectiveFormula(activeOrder).length + 1,
      materialCode: newMaterial.materialCode,
      materialName: newMaterial.materialName || newMaterial.materialCode,
      materialGroup: newMaterial.materialGroup,
      ratioPercent: Number(((num(newMaterial.requiredKg) / num(activeOrder.quantityKg || activeOrder.requestedWeight || 1)) * 100).toFixed(4)),
      requiredKg: 0,
      qcAdjustKg: num(newMaterial.requiredKg),
      qcAdjustPercent: '',
      qcAdjustReason: newMaterial.reason,
      toleranceKg: newMaterial.materialGroup === CHEMICAL ? 0.01 : 0.1,
      qcConfirm: true,
      qrScanned: '',
      qrStatus: 'Chờ quét',
      actualWeight: '',
      weighStatus: 'Chờ cân',
      confirmedAt: '',
      note: newMaterial.note,
      isQc1Added: true,
      addedAt,
    }
    patchOrder(activeOrder.id, (order) => ({
      ...order,
      activeProductionFormula: [...getEffectiveFormula(order), material],
    }), `QC sản xuất thử bổ sung NVL mới ${material.materialCode} - ${material.materialName}, khối lượng ${kg(material.qcAdjustKg)}, thời gian ${addedAt}.`, 'Bổ sung NVL')
    setNewMaterial({ materialCode: '', materialName: '', materialGroup: CHEMICAL, requiredKg: 0, reason: '', note: '' })
    setShowAddMaterial(false)
  }
  const getAdjustedRows = (order) => getEffectiveFormula(order || {}).map((item) => {
    const adjustedValue = item.qcAdjustKg === '' || item.qcAdjustKg == null ? item.requiredKg : num(item.qcAdjustKg)
    const diff = Number((adjustedValue - num(item.requiredKg)).toFixed(3))
    return { ...item, adjustedValue, diff }
  }).filter((item) => item.diff !== 0)

  const moveNext = (currentOrderId) => {
    const nextOrder = data.orders.find((order) => order.stage === 'qc1' && order.id !== currentOrderId)
    setActiveOrderId(nextOrder?.id || '')
  }

  const approve = (order, adjusted) => {
    patchOrder(order.id, (current) => {
      const base = getEffectiveFormula(current)
      const nextFormula = adjusted ? base.map((item) => {
        const ratio = item.qcAdjustPercent === '' ? item.ratioPercent : num(item.qcAdjustPercent)
        const requiredKg = item.qcAdjustKg === '' ? Number(((ratio * current.quantityKg) / 100).toFixed(3)) : num(item.qcAdjustKg)
        return { ...item, ratioPercent: item.isQc1Added ? Number(((requiredKg / current.quantityKg) * 100).toFixed(4)) : ratio, requiredKg, note: item.note || 'QC sản xuất thử điều chỉnh' }
      }) : base
      const history = {
        id: uid('qc1'),
        time: nowText(),
        result: adjusted ? 'Điều chỉnh' : 'Đạt',
        changes: adjusted ? getAdjustedRows(current) : [],
        addedMaterials: adjusted ? getAdjustedRows(current).filter((item) => item.isQc1Added) : [],
        createdBy: 'QC',
      }
      return {
        ...current,
        stage: 'weighing',
        status: 'Chờ cân',
        qc1Result: adjusted ? 'QC1 đã điều chỉnh & duyệt' : 'QC1 đạt',
        qc1Status: adjusted ? 'QC1 đã điều chỉnh & duyệt' : 'QC1 đạt',
        activeProductionFormula: nextFormula,
        qc1AdjustedFormula: adjusted ? nextFormula : current.qc1AdjustedFormula,
        qc1Adjustments: [...(current.qc1Adjustments || current.qc1Logs || []), history],
        qc1Logs: [...(current.qc1Logs || []), history],
        chemicalStatus: 'Ready',
        solidStatus: 'Ready',
        scaleStatus: { chemical: 'Ready', solid: 'Ready' },
      }
    }, adjusted ? `Điều chỉnh sau QC sản xuất thử lệnh SX ${order.id}, tự động có hiệu lực sau duyệt và chuyển sang Chờ cân.` : `QC sản xuất thử đạt lệnh ${order.id}, tách nguyên liệu sang tổ cân hóa và tổ cân rắn.`, adjusted ? 'Duyệt có điều chỉnh' : 'Đạt')
    moveNext(order.id)
  }

  const adjustedRows = getAdjustedRows(activeOrder)
  const hasChanges = adjustedRows.length > 0

  return (
    <div className="page-content qc-sample-page qc-trial-page">
      <section className="panel qc-sample-layout qc-trial-panel qc-trial-layout">
        <aside className="qc-queue">
          <div>
            <span className="section-kicker">Danh sách chờ</span>
            <h2>Danh sách chờ QC sản xuất thử</h2>
          </div>
          <div className="qc-queue-group">
            <h3>Lệnh đang QC sản xuất thử</h3>
            {activeOrder ? <QueueButton order={activeOrder} active onClick={() => setActiveOrderId(activeOrder.id)} /> : <p className="muted-text">Không có lệnh đang QC sản xuất thử.</p>}
          </div>
          <div className="qc-queue-group">
            <h3>Danh sách lệnh chờ kiểm tra</h3>
            {waitingOrders.filter((order) => order.id !== activeOrder?.id).map((order) => <QueueButton key={order.id} order={order} onClick={() => setActiveOrderId(order.id)} />)}
            {waitingOrders.filter((order) => order.id !== activeOrder?.id).length === 0 && <p className="muted-text">Không có lệnh chờ kiểm tra.</p>}
          </div>
          <div className="qc-queue-group">
            <h3>Danh sách đã kiểm tra</h3>
            {completedOrders.slice(-5).reverse().map((order) => <QueueButton key={order.id} order={order} completed onClick={() => {}} />)}
            {completedOrders.length === 0 && <p className="muted-text">Chưa có lệnh đã kiểm tra.</p>}
          </div>
        </aside>

        <main className="qc-active-panel">
          {!activeOrder && <div className="empty-alert">Không có lệnh QC sản xuất thử đang chờ xử lý.</div>}
          {activeOrder && (
            <>
              <div className="qc-trial-header">
                <div className="qc-trial-info">
                  <span className="section-kicker">Đang QC sản xuất thử</span>
                  <h2>Đang QC sản xuất thử</h2>
                  <div className="qc-trial-summary-grid">
                    <div><span>Mã lệnh SX</span><strong>{activeOrder.orderCode || activeOrder.id}</strong></div>
                    <div><span>Sản phẩm</span><strong>{activeOrder.productName || activeOrder.product}</strong></div>
                    <div><span>Công thức gốc</span><strong>{activeOrder.formulaCode || activeOrder.originalFormulaId} / {activeOrder.formulaVersion || activeOrder.originalFormulaVersion}</strong></div>
                    <div><span>Khối lượng</span><strong>{kg(activeOrder.requestedWeight ?? activeOrder.quantityKg)}</strong></div>
                    <div><span>LOT</span><strong>{activeOrder.lot}</strong></div>
                    <div><span>Khách hàng</span><strong>{activeOrder.customer || '-'}</strong></div>
                  </div>
                </div>
                <div className="qc-trial-action-bar">
                  {!hasChanges && <button className="primary-button touch-button" onClick={() => approve(activeOrder, false)}>Xác nhận sản xuất thử đạt</button>}
                  {hasChanges && <button className="primary-button touch-button" onClick={() => approve(activeOrder, true)}>Lưu điều chỉnh và duyệt sản xuất</button>}
                  <button className="secondary-button touch-button" onClick={() => setShowAddMaterial(true)}>Thêm NVL</button>
                </div>
              </div>

              <div className="qc-trial-table-wrapper">
                <table className="qc-trial-table compact-qc-table">
                  <colgroup>
                    <col style={{ width: '80px' }} />
                    <col style={{ width: '90px' }} />
                    <col style={{ width: '60px' }} />
                    <col style={{ width: '80px' }} />
                    <col style={{ width: '90px' }} />
                    <col style={{ width: '90px' }} />
                    <col style={{ width: '150px' }} />
                    <col style={{ width: '150px' }} />
                  </colgroup>
                  <thead>
                    <tr>
                      <th className="qc-trial-th">Mã VT</th>
                      <th className="qc-trial-th">Tên VT</th>
                      <th className="qc-trial-th">Nhóm</th>
                      <th className="qc-trial-th">Theo lệnh</th>
                      <th className="qc-trial-th">Sau QC</th>
                      <th className="qc-trial-th">Chênh lệch</th>
                      <th className="qc-trial-th">Lý do</th>
                      <th className="qc-trial-th">Ghi chú</th>
                    </tr>
                  </thead>
                  <tbody>
                    {getEffectiveFormula(activeOrder).map((item) => {
                  const adjustedValue = item.qcAdjustKg === '' || item.qcAdjustKg == null ? item.requiredKg : num(item.qcAdjustKg)
                  const diff = Number((adjustedValue - num(item.requiredKg)).toFixed(3))
                  return (
                    <tr key={item.id} className="qc-trial-main-row">
                      <td>{item.materialCode} {item.isQc1Added && <span className="qc-added-badge">NVL bổ sung</span>}</td>
                      <td>{item.materialName || item.materialCode}</td>
                      <td>{item.materialGroup}</td>
                      <td>{kg(item.requiredKg)}</td>
                      <td><input className="qc-value-input compact-input" type="number" value={item.qcAdjustKg === '' || item.qcAdjustKg == null ? item.requiredKg : item.qcAdjustKg} onChange={(event) => updateItem(activeOrder.id, item.id, 'qcAdjustKg', event.target.value)} /></td>
                      <td><span className={`qc-diff-badge ${diff > 0 ? 'diff-up' : diff < 0 ? 'diff-down' : 'diff-same'}`}>{diff > 0 ? '+' : ''}{kg(diff)}</span></td>
                      <td><input className="qc-reason-input compact-input" value={item.qcAdjustReason || ''} onChange={(event) => updateItem(activeOrder.id, item.id, 'qcAdjustReason', event.target.value)} /></td>
                      <td><input className="qc-note-input compact-input" value={item.note} onChange={(event) => updateItem(activeOrder.id, item.id, 'note', event.target.value)} /></td>
                    </tr>
                  )
                    })}
                  </tbody>
                </table>
              </div>

              <section className="v3-card qc-adjustment-list">
                <h3>Các thay đổi so với lệnh SX</h3>
                {hasChanges ? (
                  <div className="qc-change-table">
                    <SimpleTable headers={['Mã vật tư', 'Theo lệnh', 'Giá trị sau QC', 'Chênh lệch', 'Lý do', 'Ghi chú']} rows={adjustedRows.map((item) => (
                      <tr key={item.id}>
                        <td>{item.materialCode}</td>
                        <td>{kg(item.requiredKg)}</td>
                        <td>{kg(item.adjustedValue)}</td>
                        <td><span className={`qc-diff-badge ${item.diff > 0 ? 'diff-up' : item.diff < 0 ? 'diff-down' : 'diff-same'}`}>{item.diff > 0 ? '+' : ''}{kg(item.diff)}</span></td>
                        <td>{item.qcAdjustReason || '-'}</td>
                        <td>{item.note || '-'}</td>
                      </tr>
                    ))} />
                  </div>
                ) : <p className="muted-text">Không có thay đổi so với lệnh sản xuất</p>}
              </section>
            </>
          )}
        </main>
      </section>
      {showAddMaterial && (
        <div className="modal-backdrop" role="presentation">
          <div className="mixing-modal qc-add-material-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><span className="section-kicker">QC sản xuất thử</span><h2>Thêm NVL</h2></div>
              <button type="button" className="icon-button" onClick={() => setShowAddMaterial(false)} aria-label="Đóng">×</button>
            </div>
            <div className="production-form-grid">
              <label>Mã vật tư<input value={newMaterial.materialCode} onChange={(event) => setNewMaterial({ ...newMaterial, materialCode: event.target.value })} /></label>
              <label>Tên vật tư<input value={newMaterial.materialName} onChange={(event) => setNewMaterial({ ...newMaterial, materialName: event.target.value })} /></label>
              <label>Nhóm vật tư<select value={newMaterial.materialGroup} onChange={(event) => setNewMaterial({ ...newMaterial, materialGroup: event.target.value })}><option>{CHEMICAL}</option><option>{SOLID}</option></select></label>
              <label>Khối lượng bổ sung (kg)<input type="number" value={newMaterial.requiredKg} onChange={(event) => setNewMaterial({ ...newMaterial, requiredKg: event.target.value })} /></label>
              <label className="wide-field">Lý do bổ sung<input value={newMaterial.reason} onChange={(event) => setNewMaterial({ ...newMaterial, reason: event.target.value })} /></label>
              <label className="wide-field">Ghi chú<input value={newMaterial.note} onChange={(event) => setNewMaterial({ ...newMaterial, note: event.target.value })} /></label>
            </div>
            <div className="modal-actions">
              <button type="button" className="secondary-button" onClick={() => setShowAddMaterial(false)}>Hủy</button>
              <button type="button" className="primary-button" onClick={saveNewMaterial}>Lưu NVL</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function QueueButton({ order, active = false, completed = false, onClick }) {
  return (
    <button className={`qc-queue-item ${active ? 'active' : ''} ${completed ? 'completed' : ''}`} onClick={onClick}>
      <strong>{order.orderCode || order.id}</strong>
      <span>{order.productName || order.product}</span>
      <span>{kg(order.requestedWeight ?? order.quantityKg)}</span>
    </button>
  )
}

// eslint-disable-next-line no-unused-vars
function QC2({ data, setData }) {
  if (data?.useQc2V3 !== false) return <FinishedProductQcPage data={data} setData={setData} user={{ role: 'Admin' }} />
  const orders = data.orders.filter((order) => order.stage === 'qc2')
  // eslint-disable-next-line react-hooks/rules-of-hooks
  const [forms, setForms] = useState({})
  const getForm = (id) => forms[id] || { result: 'OK', color: '', ph: '', viscosity: '', density: '', coverage: '', note: '', materialCode: '', materialName: '', materialGroup: CHEMICAL, addKg: 0, reason: '' }
  const setForm = (id, patch) => setForms((current) => ({ ...current, [id]: { ...getForm(id), ...patch } }))
  const save = (order) => {
    const form = getForm(order.id)
    setData((current) => {
      const orders = current.orders.map((item) => {
        if (item.id !== order.id) return item
        if (form.result === 'OK') return { ...item, stage: 'packaging', status: 'Hoàn thành', qc2: form, updatedAt: nowText() }
        if (form.result === 'Không đạt') return { ...item, stage: 'qc2', status: 'QC2 không đạt', qc2: form, updatedAt: nowText() }
        const ticket = { id: uid('BS'), status: 'Pending', items: [{ id: uid('add'), materialCode: form.materialCode, materialName: form.materialName, materialGroup: form.materialGroup, requiredKg: num(form.addKg), toleranceKg: form.materialGroup === CHEMICAL ? 0.01 : 0.1, qrScanned: '', qrStatus: 'Chờ quét', actualWeight: '', weighStatus: 'Chờ cân', reason: form.reason, note: form.note }] }
        return { ...item, stage: 'supplement-weighing', status: 'Cân bổ sung', qc2: form, qc2AdjustedFormula: [...(item.qc2AdjustedFormula || []), ticket], updatedAt: nowText() }
      })
      const log = form.result === 'OK' ? `QC2 OK lệnh ${order.id}, chuyển đóng gói.` : form.result === 'Cần chỉnh màu' ? `QC2 điều chỉnh màu ${order.id}, tạo phiếu cân bổ sung.` : `QC2 không đạt lệnh ${order.id}.`
      return addLogToData({ ...current, orders }, log)
    })
  }
  return <section className="panel"><h2>QC2 - Test màu thành phẩm</h2>{orders.map((order) => {
    const form = getForm(order.id)
    return <article className="v3-card" key={order.id}>
      <div className="section-heading-row"><div><h3>{order.id} - {order.product}</h3><p className="panel-text">LOT {order.lot}, khối lượng sau phối trộn {kg(order.mixing?.finalWeightKg || order.quantityKg)}</p></div><button className="primary-button" onClick={() => save(order)}>Lưu kết quả QC2</button></div>
      <div className="production-form-grid">
        <label>Kết quả<select value={form.result} onChange={(event) => setForm(order.id, { result: event.target.value })}><option>OK</option><option>Cần chỉnh màu</option><option>Không đạt</option></select></label>
        {['color', 'ph', 'viscosity', 'density', 'coverage', 'note'].map((field) => <label key={field}>{({ color: 'Màu sắc', ph: 'pH', viscosity: 'Độ nhớt', density: 'Tỷ trọng', coverage: 'Độ phủ', note: 'Ghi chú' })[field]}<input value={form[field]} onChange={(event) => setForm(order.id, { [field]: event.target.value })} /></label>)}
      </div>
      {form.result === 'Cần chỉnh màu' && <div className="production-form-grid"><label>Mã vật tư<input value={form.materialCode} onChange={(event) => setForm(order.id, { materialCode: event.target.value })} /></label><label>Tên vật tư<input value={form.materialName} onChange={(event) => setForm(order.id, { materialName: event.target.value })} /></label><label>Nhóm<select value={form.materialGroup} onChange={(event) => setForm(order.id, { materialGroup: event.target.value })}><option>{CHEMICAL}</option><option>{SOLID}</option></select></label><label>Khối lượng bổ sung<input type="number" value={form.addKg} onChange={(event) => setForm(order.id, { addKg: event.target.value })} /></label><label>Lý do điều chỉnh<input value={form.reason} onChange={(event) => setForm(order.id, { reason: event.target.value })} /></label></div>}
    </article>
  })}{orders.length === 0 && <p className="empty-alert">Không có lệnh chờ QC2.</p>}</section>
}

function FinishedProductQcPage({ data, setData, user }) {
  const canEdit = ['QC', 'Quản đốc', 'Admin'].includes(user?.role)
  const qcStatuses = ['Chờ QC thành phẩm', 'Đang QC thành phẩm', 'Cần điều chỉnh', 'QC thành phẩm không đạt']
  const orders = data.orders
    .filter((order) => order.stage === 'finished-qc' || qcStatuses.includes(order.status))
    .slice()
    .sort(sortOldestOrders)
  const [activeOrderId, setActiveOrderId] = useState(orders[0]?.id || '')
  const [qc2Tab, setQc2Tab] = useState('current')
  const [notice, setNotice] = useState('')
  const [form, setForm] = useState({
    color: '',
    ph: '',
    viscosity: '',
    density: '',
    coverage: '',
    fineness: '',
    note: '',
    result: 'Đạt',
    reason: '',
    adjustments: [],
    newMaterials: [],
  })
  const activeOrder = orders.find((order) => order.id === activeOrderId) || orders[0]
  const activeAdjustments = activeOrder ? getQc2Adjustments(activeOrder) : []
  const activeItems = activeAdjustments.flatMap((ticket) => getAdjustmentItems(ticket).map((item) => ({ ...item, ticket })))
  const qc2Rows = activeOrder ? buildQc2Rows(activeOrder, form.adjustments) : []
  const currentAssignments = getActiveAssignments(data.productionAssignments || [], 'QC')
  const assignmentEmployeeText = getAssignmentLogContext(currentAssignments).employee

  const createDemoData = () => {
    const payload = createQc2DemoPayload(data)
    setData((current) => ({
      ...current,
      orders: [...(current.orders || []), ...payload.orders],
      qc2Logs: [...(current.qc2Logs || []), ...payload.qc2Logs],
      qc2AdjustmentTickets: [...(current.qc2AdjustmentTickets || current.qc2Adjustments || []), ...payload.qc2AdjustmentTickets],
      supplementalWeighing: [...(current.supplementalWeighing || []), ...payload.supplementalWeighing],
      productionLogs: [...(current.productionLogs || current.logs || []), ...payload.productionLogs],
      logs: [...(current.logs || []), ...payload.productionLogs],
    }))
    if (payload.orders[0]) setActiveOrderId(payload.orders[0].id)
    setNotice('Đã tạo dữ liệu demo QC thành phẩm')
  }

  const selectOrder = (order) => {
    setActiveOrderId(order.id)
    setQc2Tab('current')
    setForm({
      color: order.qc2?.color || '',
      ph: order.qc2?.ph || '',
      viscosity: order.qc2?.viscosity || '',
      density: order.qc2?.density || '',
      coverage: order.qc2?.coverage || '',
      fineness: order.qc2?.fineness || '',
      note: order.qc2?.note || '',
      result: 'Đạt',
      reason: '',
      adjustments: [],
      newMaterials: [],
    })
    setData((current) => addLogToData({
      ...current,
      orders: current.orders.map((item) => item.id === order.id ? { ...item, status: 'Đang QC thành phẩm', qc2Status: 'Đang QC', updatedAt: nowText() } : item),
      qc2Logs: [...(current.qc2Logs || []), { id: uid('qc2'), orderId: order.id, time: nowText(), action: 'Bắt đầu QC thành phẩm' }],
    }, `Bắt đầu QC thành phẩm lệnh ${order.id}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'QC', order, result: 'Bắt đầu QC thành phẩm' })))
  }

  const updateExistingAdjustment = (row, field, value) => {
    setForm((current) => ({
      ...current,
      adjustments: current.adjustments.some((item) => item.materialCode === row.materialCode)
        ? current.adjustments.map((item) => item.materialCode === row.materialCode ? { ...item, [field]: field === 'adjustmentKg' ? value : value } : item)
        : [...current.adjustments, {
          id: uid('adj'),
          materialCode: row.materialCode,
          materialName: row.materialName,
          materialGroup: row.materialGroup,
          changeType: 'existing',
          adjustmentKg: field === 'adjustmentKg' ? value : '',
          reason: field === 'reason' ? value : '',
          note: field === 'note' ? value : '',
        }],
    }))
  }

  const addNewMaterial = () => {
    setForm((current) => ({
      ...current,
      newMaterials: [...current.newMaterials, { id: uid('newqc2'), materialCode: '', materialName: '', materialGroup: CHEMICAL, adjustmentKg: 0, reason: '', note: '' }],
    }))
  }

  const updateNewMaterial = (id, field, value) => {
    setForm((current) => ({
      ...current,
      newMaterials: current.newMaterials.map((item) => item.id === id ? { ...item, [field]: field === 'adjustmentKg' ? Number(value) || 0 : value } : item),
    }))
  }

  const saveQc2 = () => {
    if (!activeOrder) return
    const checkedAt = nowText()
    const qc2Record = { ...form, checkedAt, orderId: activeOrder.id }
    setData((current) => {
      if (form.result === 'Đạt') {
        return addLogToData({
          ...current,
          orders: current.orders.map((item) => item.id === activeOrder.id ? {
            ...item,
            stage: 'packaging',
            status: 'Hoàn thành',
            qc2: qc2Record,
            qc2Status: 'Đạt',
            updatedAt: checkedAt,
          } : item),
          qc2Logs: [...(current.qc2Logs || []), { id: uid('qc2'), orderId: activeOrder.id, time: checkedAt, action: 'QC2 đạt', result: 'Đạt' }],
        }, `QC2 đạt lệnh ${activeOrder.id}. Chuyển đóng gói.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'QC', order: activeOrder, result: 'QC2 đạt' }))
      }

      if (form.result === 'Không đạt') {
        return addLogToData({
          ...current,
          orders: current.orders.map((item) => item.id === activeOrder.id ? {
            ...item,
            stage: 'finished-qc',
            status: 'QC thành phẩm không đạt',
            qc2: qc2Record,
            qc2Status: 'Không đạt',
            updatedAt: checkedAt,
          } : item),
          qc2Logs: [...(current.qc2Logs || []), { id: uid('qc2'), orderId: activeOrder.id, time: checkedAt, action: 'QC2 không đạt', result: 'Không đạt' }],
        }, `QC thành phẩm không đạt lệnh ${activeOrder.id}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'QC', order: activeOrder, result: 'QC2 không đạt' }))
      }

      const adjustmentNo = nextQc2AdjustmentNo(current)
      const adjustmentId = qc2AdjustmentId(adjustmentNo)
      const existingItems = form.adjustments
        .filter((item) => item.materialCode && num(item.adjustmentKg) !== 0)
        .map((item) => ({
          ...item,
          id: uid('adjline'),
          adjustmentId,
          orderId: activeOrder.id,
          adjustmentKg: Number(num(item.adjustmentKg).toFixed(3)),
          requiredKg: Math.max(0, Number(num(item.adjustmentKg).toFixed(3))),
          materialName: item.materialName || item.materialCode,
        }))
      const newItems = form.newMaterials
        .filter((item) => item.materialCode && num(item.adjustmentKg) > 0)
        .map((item) => ({
          ...item,
          id: uid('adjline'),
          adjustmentId,
          orderId: activeOrder.id,
          changeType: 'new',
          adjustmentKg: Number(num(item.adjustmentKg).toFixed(3)),
          requiredKg: Number(num(item.adjustmentKg).toFixed(3)),
          materialName: item.materialName || item.materialCode,
        }))
      const adjustmentItems = [...existingItems, ...newItems]
      const weighItems = adjustmentItems
        .filter((item) => item.adjustmentKg > 0)
        .map((item) => ({
          ...item,
          id: uid('add'),
          toleranceKg: item.materialGroup === CHEMICAL ? 0.01 : 0.1,
          qrScanned: '',
          qrStatus: 'Chờ quét',
          actualWeight: '',
          weighStatus: 'Chờ cân',
          confirmedAt: '',
        }))
      const adjustmentTicket = {
        id: adjustmentId,
        adjustmentId,
        AdjustmentID: adjustmentId,
        orderId: activeOrder.id,
        OrderID: activeOrder.id,
        adjustmentNo,
        AdjustmentNo: adjustmentNo,
        createdBy: user?.name || user?.username || user?.role || 'QC',
        CreatedBy: user?.name || user?.username || user?.role || 'QC',
        createdAt: checkedAt,
        CreatedAt: checkedAt,
        reason: form.reason || adjustmentItems.map((item) => item.reason).filter(Boolean).join('; ') || 'QC2 cần điều chỉnh',
        Reason: form.reason || adjustmentItems.map((item) => item.reason).filter(Boolean).join('; ') || 'QC2 cần điều chỉnh',
        status: weighItems.length ? 'PendingSupplementWeighing' : 'Completed',
        Status: weighItems.length ? 'PendingSupplementWeighing' : 'Completed',
        qc2No: activeAdjustments.length + 1,
        qc2ResultAfterAdjustment: weighItems.length ? 'Chờ cân bổ sung QC2' : 'Chờ QC lại',
        qc2Record,
        items: adjustmentItems,
        totalSupplementKg: weighItems.reduce((sum, item) => sum + num(item.requiredKg), 0),
        orderSnapshotBeforeAdjustment: activeOrder.initialOrderSnapshot || {
          id: activeOrder.id,
          orderCode: activeOrder.orderCode || activeOrder.id,
          productName: activeOrder.productName || activeOrder.product,
          lot: activeOrder.lot,
          requestedWeight: activeOrder.requestedWeight ?? activeOrder.quantityKg,
          createdAt: activeOrder.createdAt,
        },
        originalFormulaSnapshot: activeOrder.originalFormulaSnapshot || activeOrder.productionFormulaSnapshot || [],
        qc1AdjustedFormulaSnapshot: activeOrder.qc1AdjustedFormula || null,
      }
      const ticket = {
        id: uid('BS'),
        adjustmentId,
        orderId: activeOrder.id,
        type: 'QC2SupplementWeighing',
        label: 'Cân bổ sung QC2',
        status: weighItems.length ? 'Pending' : 'Completed',
        createdAt: checkedAt,
        createdBy: adjustmentTicket.createdBy,
        items: weighItems,
      }
      return addLogToData({
        ...current,
        orders: current.orders.map((item) => item.id === activeOrder.id ? {
          ...item,
          stage: weighItems.length ? 'supplement-weighing' : 'finished-qc',
          status: weighItems.length ? 'Chờ cân bổ sung QC2' : 'Chờ QC thành phẩm',
          qc2: qc2Record,
          qc2Status: 'Cần điều chỉnh',
          qc2Adjustments: [...getQc2Adjustments(item), adjustmentTicket],
          qc2SupplementTickets: [...getQc2SupplementTickets(item), ticket],
          qc2AdjustedFormula: [...getQc2SupplementTickets(item), ticket],
          updatedAt: checkedAt,
        } : item),
        qc2AdjustmentTickets: [...(current.qc2AdjustmentTickets || []), adjustmentTicket],
        supplementalWeighing: [...(current.supplementalWeighing || []), { ...ticket, orderId: activeOrder.id }],
        qc2Logs: [...(current.qc2Logs || []), { id: uid('qc2'), orderId: activeOrder.id, time: checkedAt, action: 'QC2 cần điều chỉnh', result: 'Cần điều chỉnh', adjustmentId }],
      }, `QC2 cần điều chỉnh lệnh ${activeOrder.id}. Tạo phiếu điều chỉnh ${adjustmentId} và phiếu cân bổ sung ${ticket.id}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'QC', order: activeOrder, result: 'QC2 cần điều chỉnh' }))
    })
  }

  const totalSupplementKg = activeItems.reduce((sum, item) => sum + Math.max(0, num(item.adjustmentKg ?? item.requiredKg)), 0)
  const topMaterial = mostSupplementedMaterial(activeItems)
  const reasons = activeAdjustments.map((ticket) => ticket.reason).filter(Boolean).join('; ') || '-'
  const firstPassText = activeOrder?.qc2?.result === 'Đạt' && activeAdjustments.length === 0 ? 'Đạt lần đầu' : activeAdjustments.length > 0 ? 'Phải điều chỉnh' : '-'

  return (
    <div className="page-content qc-sample-page qc2-page">
      <section className="panel qc-sample-layout qc2-layout finished-qc-layout qc-finished-layout">
        <aside className="qc-queue finished-qc-left qc-finished-left qc-waiting-panel">
          <div>
            <span className="section-kicker">Danh sách chờ</span>
            <h2>Danh sách lệnh chờ QC thành phẩm</h2>
          </div>
          <button className="primary-button touch-button" onClick={createDemoData}>Tạo dữ liệu demo QC2</button>
          {notice && <div className="process-alert success-alert">{notice}</div>}
          <div className="qc-queue-group">
            {orders.map((order) => (
              <button key={order.id} className={`qc-queue-item ${order.id === activeOrder?.id ? 'active' : ''}`} onClick={() => selectOrder(order)}>
                <strong>{order.orderCode || order.id}</strong>
                <span>{order.productName || order.product}</span>
                <span>LOT {order.lot}</span>
                <span>{kg(order.mixing?.finalWeightKg || order.quantityKg)}</span>
                <span>Số lần QC2: {qc2AttemptCount(order, data)}</span>
              </button>
            ))}
            {orders.length === 0 && <p className="muted-text">Không có lệnh chờ QC thành phẩm.</p>}
          </div>
        </aside>

        <main className="qc-active-panel qc2-active-panel">
          {!activeOrder && <div className="empty-alert">Không có lệnh thành phẩm đang chờ kiểm tra.</div>}
          {activeOrder && (
            <>
              <div>
                <span className="section-kicker">Đang kiểm tra thành phẩm</span>
                <div className="section-heading-row">
                  <h2>Đang kiểm tra thành phẩm</h2>
                  <div className="log-tabs">
                    <button className={qc2Tab === 'current' ? 'active' : ''} onClick={() => setQc2Tab('current')}>QC2 hiện tại</button>
                    <button className={qc2Tab === 'history' ? 'active' : ''} onClick={() => setQc2Tab('history')}>Lịch sử điều chỉnh</button>
                    <button className={qc2Tab === 'analysis' ? 'active' : ''} onClick={() => setQc2Tab('analysis')}>Phân tích chất lượng</button>
                  </div>
                </div>
                <div className="qc-order-summary finished-qc-info-grid qc-finished-info-grid">
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Mã lệnh SX</span><strong className="value">{activeOrder.orderCode || activeOrder.id}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Khách hàng</span><strong className="value">{activeOrder.customer || '-'}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Sản phẩm</span><strong className="value">{activeOrder.productName || activeOrder.product}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">LOT</span><strong className="value">{activeOrder.lot}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Công thức gốc</span><strong className="value">{activeOrder.formulaCode || activeOrder.originalFormulaId}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Khối lượng yêu cầu</span><strong className="value">{kg(activeOrder.requestedWeight ?? activeOrder.quantityKg)}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Khối lượng sau phối trộn</span><strong className="value">{kg(activeOrder.mixing?.finalWeightKg || activeOrder.quantityKg)}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Máy phối trộn</span><strong className="value">{getOrderAssignedMachineLabel(activeOrder, data.mixingMachines)}</strong></div>
                  <div className="finished-qc-info-card qc-finished-info-card"><span className="label">Hoàn thành phối trộn</span><strong className="value">{activeOrder.mixingCompletedAt || activeOrder.mixing?.completedAt || '-'}</strong></div>
                </div>
              </div>

              {qc2Tab === 'current' && (
                <>
                  <section className="v3-card">
                    <div className="section-heading-row">
                      <h3>Bảng QC thành phẩm</h3>
                      <button className="secondary-button" onClick={addNewMaterial}>Thêm NVL bổ sung</button>
                    </div>
                    <div className="qc2-material-table finished-qc-table-wrapper qc-finished-table-wrapper">
                      <table className="finished-qc-table qc-finished-table">
                        <thead>
                          <tr>
                            <th>Mã VT</th>
                            <th>Tên vật tư</th>
                            <th>Nhóm</th>
                            <th><span>Khối lượng</span><span>gốc</span></th>
                            <th><span>Khối lượng</span><span>sau QC1</span></th>
                            <th>QC2 bổ sung</th>
                            <th><span>Khối lượng</span><span>sau QC2</span></th>
                            <th><span>Lý do</span><span>điều chỉnh</span></th>
                            <th>Ghi chú</th>
                          </tr>
                        </thead>
                        <tbody>
                          {qc2Rows.map((row) => (
                            <tr key={row.id}>
                              <td>{row.materialCode}</td>
                              <td>{row.materialName}</td>
                              <td>{row.materialGroup}</td>
                              <td>{kg(row.originalKg)}</td>
                              <td>{kg(row.qc1Kg)}</td>
                              <td><input type="number" step="0.001" value={row.supplementKg} onChange={(event) => updateExistingAdjustment(row, 'adjustmentKg', event.target.value)} /></td>
                              <td>{kg(row.afterQc2Kg)}</td>
                              <td><input value={row.reason} onChange={(event) => updateExistingAdjustment(row, 'reason', event.target.value)} /></td>
                              <td><input value={row.note} onChange={(event) => updateExistingAdjustment(row, 'note', event.target.value)} /></td>
                            </tr>
                          ))}
                          {!qc2Rows.length && <tr><td className="empty-row" colSpan="9">Không có dữ liệu QC thành phẩm.</td></tr>}
                        </tbody>
                      </table>
                    </div>
                  </section>

                  {form.newMaterials.length > 0 && (
                    <section className="v3-card">
                      <h3>NVL bổ sung mới</h3>
                      <SimpleTable headers={['Mã VT', 'Tên vật tư', 'Nhóm', 'Khối lượng bổ sung', 'Lý do bổ sung', 'Ghi chú']} rows={form.newMaterials.map((item) => (
                        <tr key={item.id}>
                          <td><input value={item.materialCode} onChange={(event) => updateNewMaterial(item.id, 'materialCode', event.target.value)} /></td>
                          <td><input value={item.materialName} onChange={(event) => updateNewMaterial(item.id, 'materialName', event.target.value)} /></td>
                          <td><select value={item.materialGroup} onChange={(event) => updateNewMaterial(item.id, 'materialGroup', event.target.value)}><option>{CHEMICAL}</option><option>{SOLID}</option></select></td>
                          <td><input type="number" step="0.001" value={item.adjustmentKg} onChange={(event) => updateNewMaterial(item.id, 'adjustmentKg', event.target.value)} /></td>
                          <td><input value={item.reason} onChange={(event) => updateNewMaterial(item.id, 'reason', event.target.value)} /></td>
                          <td><input value={item.note} onChange={(event) => updateNewMaterial(item.id, 'note', event.target.value)} /></td>
                        </tr>
                      ))} />
                    </section>
                  )}

                  <section className="v3-card">
                    <h3>Chỉ tiêu QC thành phẩm</h3>
                    <div className="production-form-grid">
                      <label>Màu sắc<input value={form.color} onChange={(event) => setForm({ ...form, color: event.target.value })} /></label>
                      <label>pH<input value={form.ph} onChange={(event) => setForm({ ...form, ph: event.target.value })} /></label>
                      <label>Độ nhớt<input value={form.viscosity} onChange={(event) => setForm({ ...form, viscosity: event.target.value })} /></label>
                      <label>Tỷ trọng<input value={form.density} onChange={(event) => setForm({ ...form, density: event.target.value })} /></label>
                      <label>Độ phủ<input value={form.coverage} onChange={(event) => setForm({ ...form, coverage: event.target.value })} /></label>
                      <label>Độ mịn<input value={form.fineness} onChange={(event) => setForm({ ...form, fineness: event.target.value })} /></label>
                      <label>Ghi chú kiểm tra<input value={form.note} onChange={(event) => setForm({ ...form, note: event.target.value })} /></label>
                      <label>Kết quả QC2<select value={form.result} onChange={(event) => setForm({ ...form, result: event.target.value })}><option>Đạt</option><option>Cần điều chỉnh</option><option>Không đạt</option></select></label>
                      <label className="wide-field">Lý do điều chỉnh<input value={form.reason} onChange={(event) => setForm({ ...form, reason: event.target.value })} /></label>
                    </div>
                    <div className="modal-actions">
                      {canEdit ? <button type="button" className="primary-button touch-button" onClick={saveQc2}>Lưu kết quả QC2</button> : <span className="muted-text">Chỉ xem</span>}
                    </div>
                  </section>
                </>
              )}

              {qc2Tab === 'history' && (
                <section className="v3-card">
                  <h3>Lịch sử điều chỉnh</h3>
                  <SimpleTable headers={['Lần QC2', 'Ngày giờ', 'Người QC', 'Lý do', 'NVL bổ sung', 'Tổng kg bổ sung', 'Kết quả sau điều chỉnh']} rows={activeAdjustments.map((ticket, index) => (
                    <tr key={ticket.id}>
                      <td>{ticket.qc2No || index + 1}</td>
                      <td>{ticket.createdAt || '-'}</td>
                      <td>{ticket.createdBy || '-'}</td>
                      <td>{ticket.reason || '-'}</td>
                      <td>{getAdjustmentItems(ticket).map((item) => `${item.materialCode} ${num(item.adjustmentKg ?? item.requiredKg) > 0 ? '+' : ''}${kg(item.adjustmentKg ?? item.requiredKg)}`).join(', ')}</td>
                      <td>{kg(ticket.totalSupplementKg || getAdjustmentItems(ticket).reduce((sum, item) => sum + Math.max(0, num(item.adjustmentKg ?? item.requiredKg)), 0))}</td>
                      <td>{ticket.qc2ResultAfterAdjustment || displayQc2Status(ticket.status)}</td>
                    </tr>
                  ))} empty="Chưa có lịch sử điều chỉnh QC2." />
                </section>
              )}

              {qc2Tab === 'analysis' && (
                <div className="qc2-analysis-grid qc2-order-analysis">
                  <article className="executive-kpi"><span>Số lần QC2</span><strong>{qc2AttemptCount(activeOrder, data)}</strong></article>
                  <article className="executive-kpi"><span>Tổng kg bổ sung</span><strong>{kg(totalSupplementKg)}</strong></article>
                  <article className="executive-kpi"><span>NVL bổ sung nhiều nhất</span><strong>{topMaterial}</strong></article>
                  <article className="executive-kpi"><span>Đánh giá</span><strong>{firstPassText}</strong></article>
                  <section className="v3-card">
                    <h3>Lý do điều chỉnh</h3>
                    <p className="panel-text">{reasons}</p>
                  </section>
                </div>
              )}
            </>
          )}
        </main>
      </section>
    </div>
  )
}

function WeighingPage({ data, setData, group, user }) {
  const label = group === CHEMICAL ? 'Cân hóa chất' : 'Cân nguyên liệu rắn'
  const activeTitle = group === CHEMICAL ? 'Lệnh đang cân hóa' : 'Lệnh đang cân rắn'
  const completionKey = group === CHEMICAL ? 'ChemicalCompleted' : 'SolidCompleted'
  const statusKey = group === CHEMICAL ? 'chemicalStatus' : 'solidStatus'
  const scaleKey = group === CHEMICAL ? 'chemical' : 'solid'
  const isSupplementOrder = (order) => order.stage === 'supplement-weighing'
  const getSupplementTickets = (order) => getQc2SupplementTickets(order).filter((ticket) => ticket.status !== 'Completed')
  const getItems = (order) => order.stage === 'supplement-weighing'
    ? getSupplementTickets(order).flatMap((ticket) => getTicketItems(ticket).map((item) => ({ ...item, ticketId: ticket.id, adjustmentId: ticket.adjustmentId, weighingType: ticket.label || 'Cân bổ sung QC2' }))).filter((item) => item.materialGroup === group)
    : getEffectiveFormula(order).filter((item) => item.materialGroup === group)
  const isDone = (item) => item.qrStatus === 'PASS' && item.weighStatus === 'PASS'
  const groupCompleted = (order) => {
    if (isSupplementOrder(order)) {
      const items = getItems(order)
      return items.length > 0 && items.every(isDone)
    }
    return Boolean(order[completionKey]) || order[statusKey] === 'Completed' || order.scaleStatus?.[scaleKey] === 'Completed'
  }
  const relevantOrders = data.orders.filter((order) => getItems(order).length > 0 || groupCompleted(order))
  const pendingOrders = relevantOrders.filter((order) => ['weighing', 'supplement-weighing'].includes(order.stage) && !groupCompleted(order) && getItems(order).length > 0)
  const completedOrders = relevantOrders.filter((order) => groupCompleted(order))
  const [activeOrderId, setActiveOrderId] = useState('')
  const [warning, setWarning] = useState('')
  const [printQrModal, setPrintQrModal] = useState(null)
  const [weighingDetailModal, setWeighingDetailModal] = useState(null)
  const [scaleStatus, setScaleStatus] = useState('Chưa kết nối cân')
  const [scaleRawText, setScaleRawText] = useState('')
  const [scaleWeightKg, setScaleWeightKg] = useState(null)
  const [scaleSupported, setScaleSupported] = useState(true)
  const [scaleBaudRate, setScaleBaudRate] = useState(9600)
  const [scalePortLabel, setScalePortLabel] = useState('Chưa kết nối')
  const scalePortRef = useRef(null)
  const scaleReaderRef = useRef(null)
  const scaleReadingRef = useRef(false)
  const activeOrder = pendingOrders.find((order) => order.id === activeOrderId)
  const waitingOrders = pendingOrders.filter((order) => order.id !== activeOrder?.id && order.status === 'Chờ cân')
  const activeItems = activeOrder ? getItems(activeOrder) : []
  const doneCount = activeItems.filter(isDone).length
  const progress = activeItems.length ? Math.round((doneCount / activeItems.length) * 100) : 0
  const activeItem = activeItems.find((item) => !isDone(item))
  const canFinish = Boolean(activeOrder && activeItems.length && doneCount === activeItems.length)
  const activeWeighingType = activeOrder?.stage === 'supplement-weighing' ? 'Cân bổ sung QC thành phẩm' : 'Cân chính'
  const activeContainers = activeOrder ? getOrderGroupContainers(data.weighedContainers || [], activeOrder, activeWeighingType).filter((item) => item.materialGroup === group) : []
  const completedWeighingContainers = normalizeWeighedContainers(data.weighedContainers || []).filter((item) => item.materialGroup === group).reverse()
  const latestContainer = completedWeighingContainers[0]
  const assignmentStage = group === CHEMICAL ? 'Cân hóa' : 'Cân rắn'
  const currentAssignments = getActiveAssignments(data.productionAssignments || [], assignmentStage)
  const assignmentEmployeeText = getAssignmentLogContext(currentAssignments).employee

  useEffect(() => {
    setScaleSupported(typeof navigator !== 'undefined' && Boolean(navigator.serial))
    return () => {
      scaleReadingRef.current = false
      scaleReaderRef.current?.cancel().catch(() => {})
      scalePortRef.current?.close().catch(() => {})
    }
  }, [])

  const connectScale = async () => {
    if (typeof navigator === 'undefined' || !navigator.serial) {
      setScaleSupported(false)
      setWarning('Trình duyệt không hỗ trợ kết nối cân. Vui lòng dùng Chrome hoặc Edge.')
      return
    }
    try {
      scaleReadingRef.current = false
      await scaleReaderRef.current?.cancel().catch(() => {})
      await scalePortRef.current?.close().catch(() => {})
      const port = await navigator.serial.requestPort()
      await port.open({
        baudRate: Number(scaleBaudRate),
        dataBits: 8,
        stopBits: 1,
        parity: 'none',
        flowControl: 'none',
      })
      const portInfo = port.getInfo?.() || {}
      const portLabel = portInfo.usbVendorId
        ? `USB-RS232 VID:${portInfo.usbVendorId} PID:${portInfo.usbProductId || '-'}`
        : 'Cổng serial đã chọn'
      scalePortRef.current = port
      scaleReadingRef.current = true
      setScaleStatus(`Đã kết nối cân (${scaleBaudRate} baud)`)
      setScalePortLabel(portLabel)
      setScaleRawText('')
      setScaleWeightKg(null)
      setWarning('')
      const decoder = new TextDecoder()
      let buffer = ''
      while (scaleReadingRef.current && port.readable) {
        const reader = port.readable.getReader()
        scaleReaderRef.current = reader
        try {
          while (scaleReadingRef.current) {
            const { value, done } = await reader.read()
            if (done) break
            const chunk = decoder.decode(value, { stream: true })
            if (!chunk) continue
            buffer = `${buffer}${chunk}`.slice(-240)
            setScaleRawText(buffer.trim())
            const parsed = parseScaleWeightKg(buffer)
            if (parsed != null) setScaleWeightKg(parsed)
          }
        } finally {
          reader.releaseLock()
        }
      }
    } catch (error) {
      if (error?.name === 'NotFoundError') {
        setScaleStatus('Chưa chọn cổng cân')
        setScalePortLabel('Chưa kết nối')
        return
      }
      setScaleStatus('Mất kết nối cân')
      setWarning(`Không đọc được dữ liệu cân: ${error.message}`)
    }
  }

  const handleViewWeighingDetail = (container) => {
    setWeighingDetailModal(container)
    setData((current) => addLogToData(current, `Xem chi tiết cân ${container.materialGroup} ${container.qrCode} cho lệnh ${container.orderCode}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: assignmentStage, order: container.orderCode, result: 'Xem chi tiết cân' })))
  }
  const handlePrintQr = (container) => {
    setPrintQrModal(container)
    setData((current) => addLogToData(current, `In QR hỗn hợp ${container.materialGroup} ${container.qrCode} cho lệnh ${container.orderCode}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: assignmentStage, order: container.orderCode, result: 'In QR hỗn hợp' })))
    setTimeout(() => window.print(), 80)
  }
  const printSelectedContainer = (container) => {
    setData((current) => addLogToData(current, `In QR hỗn hợp ${container.materialGroup} ${container.qrCode} cho lệnh ${container.orderCode}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: assignmentStage, order: container.orderCode, result: 'In QR hỗn hợp' })))
    setTimeout(() => window.print(), 50)
  }

  const createDemoQrData = () => {
    setData((current) => applyDemoQrData(current))
    setWarning('Đã tạo dữ liệu demo QR hỗn hợp.')
  }

  const updateWeight = (order, item, patch, log, stockIssue = null) => setData((current) => {
    const materialLots = normalizeRawMaterialLots(current.rawMaterials || [])
    const issuedQty = num(stockIssue?.quantityOut)
    const nextRawMaterials = stockIssue && issuedQty > 0
      ? materialLots.map((lot) => lot.lotCode === stockIssue.lotCode && lot.materialCode === stockIssue.materialCode
        ? normalizeRawMaterialLot({ ...lot, remainingQty: Math.max(0, num(lot.remainingQty) - issuedQty) })
        : lot)
      : materialLots
    const stockTransactions = stockIssue && issuedQty > 0
      ? [...(current.stockTransactions || []), { id: uid('STK'), transactionType: 'ISSUE_TO_PRODUCTION', createdAt: nowText(), operator: 'Tổ cân', ...stockIssue, quantityOut: issuedQty }]
      : (current.stockTransactions || [])
    return addLogToData({ ...current, rawMaterials: nextRawMaterials, stockTransactions, orders: current.orders.map((currentOrder) => {
      if (currentOrder.id !== order.id) return currentOrder
      if (currentOrder.stage === 'supplement-weighing') {
        const updateTickets = (tickets = []) => tickets.map((ticket) => ticket.id === item.ticketId ? { ...ticket, items: getTicketItems(ticket).map((row) => row.id === item.id ? { ...row, ...patch } : row) } : ticket)
        return {
          ...currentOrder,
          qc2SupplementTickets: updateTickets(getQc2SupplementTickets(currentOrder)),
          qc2AdjustedFormula: updateTickets(getQc2SupplementTickets(currentOrder)),
          updatedAt: nowText(),
        }
      }
      const apply = (rows) => rows.map((row) => row.id === item.id ? { ...row, ...patch } : row)
      return { ...currentOrder, activeProductionFormula: apply(getEffectiveFormula(currentOrder)), qc1AdjustedFormula: currentOrder.qc1AdjustedFormula ? apply(currentOrder.qc1AdjustedFormula) : currentOrder.qc1AdjustedFormula, updatedAt: nowText() }
    }), supplementalWeighing: (current.supplementalWeighing || []).map((ticket) => ticket.id === item.ticketId ? { ...ticket, items: getTicketItems(ticket).map((row) => row.id === item.id ? { ...row, ...patch } : row) } : ticket) }, log, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: assignmentStage, order, result: patch.weighStatus || patch.qrStatus || 'Cập nhật cân' }))
  })

  const finishIfReady = (order) => setData((current) => {
    const sourceOrder = current.orders.find((item) => item.id === order.id) || order
    const weighingType = sourceOrder.stage === 'supplement-weighing' ? 'Cân bổ sung QC thành phẩm' : 'Cân chính'
    const containerItems = sourceOrder.stage === 'supplement-weighing'
      ? getQc2SupplementTickets(sourceOrder).flatMap((ticket) => getTicketItems(ticket)).filter((item) => item.materialGroup === group)
      : getEffectiveFormula(sourceOrder).filter((item) => item.materialGroup === group)
    const normalizedContainers = normalizeWeighedContainers(current.weighedContainers || [])
    const hasContainer = normalizedContainers.some((item) => (
      (item.orderId === sourceOrder.id || item.orderCode === (sourceOrder.orderCode || sourceOrder.id))
      && item.materialGroup === group
      && item.weighingType === weighingType
    ))
    const groupContainer = !hasContainer && containerItems.length && containerItems.every(isDone)
      ? buildWeighedContainer(sourceOrder, group, containerItems, normalizedContainers, weighingType)
      : null
    const orders = current.orders.map((item) => {
      if (item.id !== order.id) return item
      if (item.stage === 'supplement-weighing') {
        const tickets = getQc2SupplementTickets(item).map((ticket) => {
          const done = getTicketItems(ticket).every((row) => row.qrStatus === 'PASS' && row.weighStatus === 'PASS')
          return done ? { ...ticket, status: 'Completed' } : ticket
        })
        const allDone = tickets.every((ticket) => ticket.status === 'Completed')
        const updatedAdjustments = getQc2Adjustments(item).map((adjustment) => {
          const relatedTickets = tickets.filter((ticket) => ticket.adjustmentId === (adjustment.adjustmentId || adjustment.id))
          return relatedTickets.length && relatedTickets.every((ticket) => ticket.status === 'Completed')
            ? { ...adjustment, status: 'SupplementWeighed', Status: 'SupplementWeighed' }
            : adjustment
        })
        return allDone ? {
          ...item,
          qc2Adjustments: updatedAdjustments,
          qc2SupplementTickets: tickets,
          qc2AdjustedFormula: tickets,
          stage: 'finished-qc',
          status: 'Chờ QC thành phẩm',
          qc2Status: 'Chờ QC lại sau cân bổ sung',
          updatedAt: nowText(),
        } : {
          ...item,
          qc2Adjustments: updatedAdjustments,
          qc2SupplementTickets: tickets,
          qc2AdjustedFormula: tickets,
          updatedAt: nowText(),
        }
      }
      const effective = getEffectiveFormula(item)
      const chemDone = effective.filter((row) => row.materialGroup === CHEMICAL).every((row) => row.qrStatus === 'PASS' && row.weighStatus === 'PASS')
      const solidDone = effective.filter((row) => row.materialGroup === SOLID).every((row) => row.qrStatus === 'PASS' && row.weighStatus === 'PASS')
      const bothGroupsDone = chemDone && solidDone
      return {
        ...item,
        ChemicalCompleted: chemDone,
        SolidCompleted: solidDone,
        ReadyMixing: bothGroupsDone,
        chemicalStatus: chemDone ? 'Completed' : item.chemicalStatus,
        solidStatus: solidDone ? 'Completed' : item.solidStatus,
        scaleStatus: { chemical: chemDone ? 'Completed' : item.scaleStatus.chemical, solid: solidDone ? 'Completed' : item.scaleStatus.solid },
        stage: bothGroupsDone ? 'mixing' : item.stage,
        status: bothGroupsDone ? 'Sẵn sàng phối trộn' : 'Đang cân',
        updatedAt: nowText(),
      }
    })
    const supplementalWeighing = (current.supplementalWeighing || []).map((ticket) => {
      if (ticket.orderId !== order.id) return ticket
      const done = getTicketItems(ticket).every((row) => row.qrStatus === 'PASS' && row.weighStatus === 'PASS')
      return done ? { ...ticket, status: 'Completed' } : ticket
    })
    const logText = groupContainer
      ? `${order.stage === 'supplement-weighing' ? `Hoàn thành cân bổ sung ${group === CHEMICAL ? 'hóa' : 'rắn'}, chuyển lại QC thành phẩm.` : `${label} hoàn thành lệnh ${order.id}.`} Tạo QR hỗn hợp ${group} ${groupContainer.qrCode}.`
      : (order.stage === 'supplement-weighing' ? `Hoàn thành cân bổ sung ${group === CHEMICAL ? 'hóa' : 'rắn'}, chuyển lại QC thành phẩm.` : `${label} hoàn thành lệnh ${order.id}.`)
    return addLogToData({
      ...current,
      orders,
      supplementalWeighing,
      weighedContainers: groupContainer ? [...normalizedContainers, groupContainer] : normalizedContainers,
    }, logText, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: assignmentStage, order, result: 'Hoàn thành cân' }))
  })

  const startOrder = (order) => {
    if (activeOrder && activeOrder.id !== order.id) {
      setWarning('Đang có một lệnh cân. Vui lòng hoàn thành lệnh hiện tại trước.')
      return
    }
    if (!groupCompleted(order)) {
      const startedAt = nowText()
      setData((current) => addLogToData({
        ...current,
        orders: current.orders.map((item) => {
          if (item.id !== order.id) return item
          const supplement = item.stage === 'supplement-weighing'
          const currentScaleStatus = item.scaleStatus || {}
          return {
            ...item,
            stage: supplement ? item.stage : 'weighing',
            status: supplement ? 'Đang cân bổ sung' : 'Đang cân',
            [statusKey]: item[statusKey] === 'Completed' ? 'Completed' : 'Active',
            scaleStatus: {
              ...currentScaleStatus,
              [scaleKey]: currentScaleStatus[scaleKey] === 'Completed' ? 'Completed' : 'Active',
            },
            updatedAt: startedAt,
          }
        }),
      }, `${label} bắt đầu cân lệnh ${order.id}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: assignmentStage, order, result: 'Bắt đầu cân' })))
      setActiveOrderId(order.id)
      setWarning('')
    }
  }
  const finishActiveOrder = () => {
    if (!canFinish) return
    finishIfReady(activeOrder)
    setActiveOrderId('')
    setWarning('')
  }

  return (
    <div className="page-content weighing-dispatch-page">
      <section className="weighing-stats">
        <article><span>Đang cân</span><strong>{activeOrder ? '01' : '00'}</strong></article>
        <article><span>Chờ cân</span><strong>{String(waitingOrders.length).padStart(2, '0')}</strong></article>
        <article><span>Hoàn thành</span><strong>{String(completedOrders.length).padStart(2, '0')}</strong></article>
      </section>
      <section className="panel weighing-dispatch-layout">
        <aside className="weighing-order-list">
          <h2>Danh sách lệnh sản xuất</h2>
          <WeighingOrderGroup title="Danh sách lệnh chờ cân" orders={waitingOrders} activeId={activeOrder?.id} onStart={startOrder} showStart />
        </aside>
        <main className="weighing-active-board">
          <div className="section-heading-row">
            <div>
              <span className="section-kicker">{activeOrder?.stage === 'supplement-weighing' ? 'Cân bổ sung QC2' : 'Cân chính'}</span>
              <h2>{activeTitle}</h2>
            </div>
            <div className="action-row">
              <button className="secondary-button weighing-finish-button" type="button" onClick={connectScale}>Kết nối cân</button>
              <button className="secondary-button weighing-finish-button" onClick={createDemoQrData}>Tạo dữ liệu demo QR</button>
              {canFinish && <button className="primary-button weighing-finish-button" onClick={finishActiveOrder}>{group === CHEMICAL ? 'Hoàn thành cân hóa' : 'Hoàn thành cân rắn'}</button>}
            </div>
          </div>
          {warning && <div className="process-alert">{warning}</div>}
          <div className="process-alert assignment-context">
            <strong>Phân công ca hiện tại:</strong> {currentAssignments.length ? currentAssignments.map((item) => `${formatAssignmentEmployees(item)} (${item.shiftCode}${item.productionTeamName ? ` - ${item.productionTeamName}` : ''})`).join(', ') : 'Chưa có phân công.'}
          </div>
          {!scaleSupported && <div className="process-alert">Trình duyệt không hỗ trợ kết nối cân. Vui lòng dùng Chrome hoặc Edge.</div>}
          <div className="scale-serial-panel">
            <div><span>Trạng thái cân</span><strong>{scaleStatus}</strong></div>
            <div><span>Cổng đã kết nối</span><strong>{scalePortLabel}</strong></div>
            <label className="scale-baud-select">
              <span>BaudRate</span>
              <select value={scaleBaudRate} onChange={(event) => setScaleBaudRate(Number(event.target.value))}>
                {[2400, 4800, 9600, 19200].map((rate) => <option key={rate} value={rate}>{rate}</option>)}
              </select>
            </label>
            <div><span>Khối lượng cân hiện tại</span><strong>{scaleWeightKg == null ? '-' : kg(scaleWeightKg)}</strong></div>
            <div className="scale-raw-text"><span>Dữ liệu nhận được</span><strong>{scaleRawText || '-'}</strong></div>
          </div>
          {!activeOrder && <p className="empty-alert">Chưa có lệnh nào đang cân. Vui lòng chọn một lệnh từ danh sách chờ.</p>}
          {activeOrder && (
            <>
              <div className="weighing-order-summary">
                <div><span>Mã lệnh SX</span><strong>{activeOrder.orderCode || activeOrder.id}</strong></div>
                <div><span>Sản phẩm</span><strong>{activeOrder.productName || activeOrder.product}</strong></div>
                <div><span>LOT</span><strong>{activeOrder.lot}</strong></div>
                <div><span>Loại cân</span><strong>{activeOrder.stage === 'supplement-weighing' ? 'Cân bổ sung QC2' : 'Cân chính'}</strong></div>
                <div><span>Tiến độ đã cân</span><strong>Đã cân {doneCount}/{activeItems.length} vật tư</strong></div>
              </div>
              <div className="weighing-progress-card">
                <div><span>Đã cân</span><strong>{doneCount} / {activeItems.length} vật tư</strong></div>
                <div className="weighing-progress"><i style={{ width: `${progress}%` }} /></div>
                <strong>{progress}%</strong>
              </div>
              <SimpleTable headers={['STT', 'Mã VT yêu cầu', 'Tên VT', 'Cần cân', 'QR lô đã quét', 'Lô nhập', 'Tồn trước', 'Thực cân', 'Tồn sau', 'Khớp QR', 'Trạng thái']} rows={activeItems.map((item, index) => (
                <WeighingRow key={`${activeOrder.id}-${item.id}`} order={activeOrder} item={item} index={index} active={item.id === activeItem?.id} updateWeight={updateWeight} rawMaterialLots={normalizeRawMaterialLots(data.rawMaterials || [])} scaleType={scaleKey} setWarning={setWarning} scaleWeightKg={scaleWeightKg} />
              ))} empty={`Không có vật tư nhóm ${group}.`} />
              {activeContainers.map((container) => <WeighedContainerCard key={container.containerId} container={container} onPrint={handlePrintQr} onDetail={handleViewWeighingDetail} />)}
            </>
          )}
          <section className="weighed-container-list">
            <h3>Lệnh cân gần nhất</h3>
            {latestContainer
              ? <WeighedContainerCard container={latestContainer} title="Lệnh cân gần nhất" compact onPrint={handlePrintQr} onDetail={handleViewWeighingDetail} />
              : <p className="muted-text">Chưa có lệnh đã cân của tổ hiện tại.</p>}
          </section>
          <section className="weighed-container-list">
            <h3>Danh sách lệnh đã cân</h3>
            <SimpleTable
              tableClassName="weighed-container-history-table"
              headers={['Mã QR hỗn hợp', 'Mã lệnh SX', 'Sản phẩm', 'LOT', 'Nhóm hỗn hợp', 'Tổng kg', 'Thời gian cân', 'Trạng thái', 'Hành động']}
              rows={completedWeighingContainers.map((container) => (
                <tr key={container.containerId || container.qrCode}>
                  <td>{container.qrCode}</td>
                  <td>{container.orderCode}</td>
                  <td>{container.productName}</td>
                  <td>{container.lot}</td>
                  <td>{container.materialGroup}</td>
                  <td>{kg(container.totalWeight)}</td>
                  <td>{container.completedAt || '-'}</td>
                  <td>{container.status || '-'}</td>
                  <td>
                    <div className="action-row table-action-row">
                      <button type="button" className="secondary-button" onClick={() => handlePrintQr(container)}>In QR</button>
                      <button type="button" className="primary-button" onClick={() => handleViewWeighingDetail(container)}>Xem chi tiết</button>
                    </div>
                  </td>
                </tr>
              ))}
              empty="Chưa có lệnh đã cân của tổ hiện tại."
            />
          </section>
        </main>
      </section>
      {printQrModal && (
        <div className="modal-backdrop" role="presentation">
          <div className="mixing-modal weighed-container-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><span className="section-kicker">Phiếu in</span><h2>QR HỖN HỢP ĐÃ CÂN</h2></div>
              <button type="button" className="icon-button" onClick={() => setPrintQrModal(null)} aria-label="Đóng">×</button>
            </div>
            <QrPrintTicket container={printQrModal} />
            <div className="modal-actions">
              <button type="button" className="secondary-button" onClick={() => printSelectedContainer(printQrModal)}>In QR</button>
              <button type="button" className="primary-button" onClick={() => setPrintQrModal(null)}>Đóng</button>
            </div>
          </div>
        </div>
      )}
      {weighingDetailModal && (
        <div className="modal-backdrop" role="presentation">
          <div className="mixing-modal weighed-container-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><span className="section-kicker">Chi tiết cân</span><h2>{weighingDetailModal.orderCode}</h2></div>
              <button type="button" className="icon-button" onClick={() => setWeighingDetailModal(null)} aria-label="Đóng">×</button>
            </div>
            <div className="weighed-container-detail-grid">
              <div><span>Mã lệnh SX</span><strong>{weighingDetailModal.orderCode}</strong></div>
              <div><span>Sản phẩm</span><strong>{weighingDetailModal.productName}</strong></div>
              <div><span>LOT</span><strong>{weighingDetailModal.lot}</strong></div>
              <div><span>Nhóm hỗn hợp</span><strong>{weighingDetailModal.materialGroup}</strong></div>
              <div><span>Tổng kg</span><strong>{kg(weighingDetailModal.totalWeight)}</strong></div>
              <div><span>Thời gian cân</span><strong>{weighingDetailModal.completedAt || '-'}</strong></div>
              <div><span>Người cân</span><strong>{weighingDetailModal.weighedBy || '-'}</strong></div>
              <div><span>Trạng thái</span><strong>{weighingDetailModal.status || '-'}</strong></div>
            </div>
            <SimpleTable headers={['STT', 'Mã vật tư', 'Tên vật tư', 'Khối lượng theo lệnh', 'QR/Lô đã quét', 'Thực cân', 'Sai lệch', 'Trạng thái khớp QR']} rows={(weighingDetailModal.materials || []).map((item, index) => {
              const requiredKg = num(item.requiredKg)
              const actualWeight = num(item.actualWeight)
              const diff = Number((actualWeight - requiredKg).toFixed(3))
              const qrMatchStatus = item.qrMatchStatus || (item.qrStatus === 'PASS' || item.qrScanned ? 'Khớp QR' : '-')
              return (
                <tr key={item.id || item.materialCode}>
                  <td>{index + 1}</td>
                  <td>{item.materialCode}</td>
                  <td>{item.materialName}</td>
                  <td>{kg(requiredKg)}</td>
                  <td>{item.qrScanned || '-'}</td>
                  <td>{kg(actualWeight)}</td>
                  <td>{kg(diff)}</td>
                  <td>{qrMatchStatus}</td>
                </tr>
              )
            })} />
            <div className="modal-actions">
              <button type="button" className="primary-button" onClick={() => setWeighingDetailModal(null)}>Đóng</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function WeighedContainerCard({ container, onPrint, onDetail, title = 'QR hỗn hợp đã cân', compact = false }) {
  return (
    <article className={`weighed-container-card ${compact ? 'compact' : ''}`}>
      <div className="section-heading-row">
        <div>
          <span className="section-kicker">{title}</span>
          <h3>{container.qrCode}</h3>
        </div>
        <span className="dispatch-badge ready">{container.status}</span>
      </div>
      <div className="weighed-container-grid">
        <div><span>Mã lệnh SX</span><strong>{container.orderCode}</strong></div>
        <div><span>Sản phẩm</span><strong>{container.productName}</strong></div>
        <div><span>LOT</span><strong>{container.lot}</strong></div>
        <div><span>Nhóm hỗn hợp</span><strong>{container.materialGroup}</strong></div>
        <div><span>Tổng kg</span><strong>{kg(container.totalWeight)}</strong></div>
        <div><span>Thời gian tạo</span><strong>{container.completedAt}</strong></div>
      </div>
      <div className="action-row">
        <button type="button" className="secondary-button" onClick={() => onPrint(container)}>In QR</button>
        <button type="button" className="primary-button" onClick={() => onDetail(container)}>Xem chi tiết</button>
      </div>
    </article>
  )
}

function QrPrintTicket({ container }) {
  return (
    <section className="qr-print-ticket">
      <h1>QR HỖN HỢP ĐÃ CÂN</h1>
      <div className="qr-print-code">
        <QRCodeCanvas value={getContainerQrValue(container)} size={220} includeMargin />
      </div>
      <strong className="qr-print-text">{container.qrCode}</strong>
      <div className="qr-print-info">
        <div><span>Mã QR</span><strong>{container.qrCode}</strong></div>
        <div><span>Lệnh SX</span><strong>{container.orderCode}</strong></div>
        <div><span>Sản phẩm</span><strong>{container.productName}</strong></div>
        <div><span>LOT</span><strong>{container.lot}</strong></div>
        <div><span>Nhóm</span><strong>{container.materialGroup}</strong></div>
        <div><span>Tổng kg</span><strong>{kg(container.totalWeight)}</strong></div>
        <div><span>Loại cân</span><strong>{container.weighingType}</strong></div>
        <div><span>Thời gian</span><strong>{container.completedAt}</strong></div>
      </div>
    </section>
  )
}

function WeighingOrderGroup({ title, orders, activeId, onStart, showStart = false }) {
  return (
    <div className="weighing-order-group">
      <h3>{title}</h3>
      {orders.map((order) => (
        <article key={order.id} className={`weighing-order-card ${order.id === activeId ? 'active' : ''}`}>
          <strong>{order.orderCode || order.id}</strong>
          <span>{order.productName || order.product}</span>
          <span>{order.lot}</span>
          <span>{kg(order.requestedWeight ?? order.quantityKg)}</span>
          {showStart && <button className="secondary-button weighing-start-button" onClick={() => onStart(order)}>Bắt đầu cân</button>}
        </article>
      ))}
      {orders.length === 0 && <p className="muted-text">Không có lệnh.</p>}
    </div>
  )
}

function WeighingRow({ order, item, index, active, updateWeight, rawMaterialLots = [], scaleType, setWarning, scaleWeightKg = null }) {
  const [qrInput, setQrInput] = useState(item.qrScanned || '')
  const [weight, setWeight] = useState(item.actualWeight || '')
  const qrPassed = item.qrStatus === 'PASS'
  const weightPassed = item.weighStatus === 'PASS'
  const completed = qrPassed && weightPassed
  const actual = item.actualWeight === '' || item.actualWeight == null ? '' : num(item.actualWeight)
  const selectedLot = rawMaterialLots.find((lot) => lot.lotCode === item.rawMaterialLotCode && lot.materialCode === item.materialCode)
  const remainingBefore = item.rawMaterialRemainingBefore ?? selectedLot?.remainingQty
  const remainingAfter = item.rawMaterialRemainingAfter ?? (actual !== '' && remainingBefore != null ? Math.max(0, num(remainingBefore) - actual) : '')
  const confirmQr = () => {
    if (!active) return
    const result = validateRawMaterialQr(qrInput, item, rawMaterialLots)
    if (!result.ok) {
      setWarning?.(result.message)
      updateWeight(order, item, { qrScanned: qrInput, qrStatus: 'FAIL', qrMatchStatus: result.message, note: result.message }, `QR FAIL ${order.id} - ${item.materialCode}: ${result.message}`)
      return
    }
    setWarning?.('')
    updateWeight(order, item, {
      qrScanned: qrInput,
      rawMaterialQr: qrInput,
      rawMaterialLotCode: result.lot.lotCode,
      rawMaterialRemainingBefore: num(result.lot.remainingQty),
      qrStatus: 'PASS',
      qrMatchStatus: 'Khớp QR',
      note: item.note,
    }, `QR PASS ${order.id} - ${item.materialCode}, lô ${result.lot.lotCode}.`)
  }
  const confirmWeight = (overrideWeight = null) => {
    if (!active) return
    if (!qrPassed || !item.rawMaterialLotCode) {
      setWarning?.('Vui lòng quét QR lô nguyên liệu hợp lệ trước khi cân.')
      return
    }
    const actualWeight = num(overrideWeight ?? weight)
    const lot = rawMaterialLots.find((row) => row.lotCode === item.rawMaterialLotCode && row.materialCode === item.materialCode)
    if (!lot) {
      setWarning?.('Không tìm thấy lô nguyên liệu trong kho.')
      return
    }
    if (num(lot.remainingQty) <= 0) {
      setWarning?.('Lô nguyên liệu đã hết tồn.')
      return
    }
    if (num(lot.remainingQty) < actualWeight) {
      setWarning?.('Lô nguyên liệu không đủ tồn kho.')
      return
    }
    const pass = Math.abs(actualWeight - num(item.requiredKg)) <= num(item.toleranceKg)
    const supplementalText = order.stage === 'supplement-weighing' ? `Cân bổ sung ${item.materialGroup === CHEMICAL ? 'hóa' : 'rắn'}` : 'Cân'
    const nextRemaining = num(lot.remainingQty) - actualWeight
    setWarning?.(pass ? '' : 'Khối lượng cân ngoài dung sai.')
    updateWeight(order, item, {
      actualWeight,
      rawMaterialRemainingBefore: num(lot.remainingQty),
      rawMaterialRemainingAfter: nextRemaining,
      weighStatus: pass ? 'PASS' : 'FAIL',
      confirmedAt: nowText(),
      note: pass ? item.note : 'Ngoài dung sai',
    }, `${supplementalText} ${pass ? 'PASS' : 'FAIL'} ${order.id} - ${item.materialCode}, lô ${lot.lotCode}.`, pass ? {
      productionOrderId: order.id,
      materialCode: item.materialCode,
      lotCode: lot.lotCode,
      quantityOut: actualWeight,
      unit: lot.unit || 'kg',
      scaleType,
    } : null)
  }
  const recordScaleWeight = () => {
    if (scaleWeightKg == null) {
      setWarning?.('Chưa có dữ liệu khối lượng từ cân.')
      return
    }
    setWeight(scaleWeightKg)
    confirmWeight(scaleWeightKg)
  }
  return (
    <tr className={`${active ? 'weighing-active-row' : ''} ${completed ? 'weighing-completed-row' : ''}`}>
      <td>{index + 1}</td>
      <td>{item.materialCode}</td>
      <td>{item.materialName}</td>
      <td>{kg(item.requiredKg)}</td>
      <td><div className="weighing-inline-action">{active && !qrPassed ? <><input value={qrInput} onChange={(event) => setQrInput(event.target.value)} /><button className="secondary-button" onClick={confirmQr}>QR</button></> : <span>{item.qrScanned || '-'}</span>}{qrPassed ? <span className="weighing-check">QR PASS</span> : item.qrStatus === 'FAIL' ? <span className="weighing-fail">QR FAIL</span> : null}</div></td>
      <td>{item.rawMaterialLotCode || '-'}</td>
      <td>{remainingBefore === '' || remainingBefore == null ? '-' : kg(remainingBefore)}</td>
      <td>{active && qrPassed && !weightPassed ? <div className="weighing-inline-action"><input type="number" value={weight} onChange={(event) => setWeight(event.target.value)} /><button className="primary-button" onClick={() => confirmWeight()}>Cân</button><button className="secondary-button" disabled={scaleWeightKg == null} onClick={recordScaleWeight}>Ghi nhận khối lượng</button></div> : actual === '' ? '-' : kg(actual)}</td>
      <td>{remainingAfter === '' || remainingAfter == null ? '-' : kg(remainingAfter)}</td>
      <td>{item.qrMatchStatus || item.qrStatus || '-'}</td>
      <td>{completed ? <span className="weighing-check">✓ Hoàn thành</span> : qrPassed ? <span className="weighing-check">✓ QR đạt</span> : active ? 'Đang thao tác' : '-'}</td>
    </tr>
  )
}

function getMixingDispatchState(order) {
  const chemicalCompleted = order.scaleStatus?.chemical === 'Completed' || order.chemicalStatus === 'Completed' || order.chemicalStatus === 'completed'
  const solidCompleted = order.scaleStatus?.solid === 'Completed' || order.solidStatus === 'Completed' || order.solidStatus === 'completed'
  const chemicalActive = order.scaleStatus?.chemical === 'Active' || order.chemicalStatus === 'Active' || order.chemicalStatus === 'active'
  const solidActive = order.scaleStatus?.solid === 'Active' || order.solidStatus === 'Active' || order.solidStatus === 'active'
  const weighingStarted = chemicalActive || solidActive || chemicalCompleted || solidCompleted || order.status === 'Đang cân'
  if (order.stage === 'completed') return { label: 'Hoàn thành', className: 'done', canStart: false }
  if (order.stage === 'finished-goods') return { label: 'HOÀN THÀNH', className: 'done', canStart: false }
  if (order.stage === 'packaging') return { label: 'HOÀN THÀNH', className: 'packing', canStart: false }
  if (order.stage === 'finished-qc') return { label: 'CHỜ QC THÀNH PHẨM', className: 'qc2', canStart: false }
  if (order.mixing?.status === 'Active' || order.mixingStatus === 'Active') return { label: 'ĐANG PHỐI TRỘN', className: 'mixing', canStart: false }
  if (order.stage === 'mixing-supplement') return { label: 'Chờ phối trộn bổ sung', className: 'ready', canStart: true, supplement: true }
  if (chemicalCompleted && solidCompleted) return { label: 'Sẵn sàng phối trộn', className: 'ready', canStart: true }
  if (weighingStarted || order.stage === 'supplement-weighing') return { label: 'Đang cân', className: 'weighing', canStart: false }
  return { label: 'Chờ cân', className: 'waiting', canStart: false }
}

function getMixingProgress(order) {
  const chemicalCompleted = order.scaleStatus?.chemical === 'Completed' || order.chemicalStatus === 'completed'
  const solidCompleted = order.scaleStatus?.solid === 'Completed' || order.solidStatus === 'completed'
  const steps = [
    Boolean(order.qc1Result || order.stage !== 'qc1'),
    chemicalCompleted,
    solidCompleted,
    Boolean(order.mixing?.status === 'Completed' || order.mixingStatus === 'completed' || ['finished-qc', 'packaging', 'finished-goods', 'completed'].includes(order.stage)),
    Boolean(['OK', 'Đạt'].includes(order.qc2?.result) || ['packaging', 'finished-goods', 'completed'].includes(order.stage)),
    Boolean(order.packaging || ['finished-goods', 'completed'].includes(order.stage)),
  ]
  return Math.round((steps.filter(Boolean).length / 6) * 100)
}

function MixingPage({ data, setData, user }) {
  const machines = normalizeMixingMachines(data.mixingMachines)
  const activeMachines = getActiveMixingMachines(machines)
  const orders = data.orders
  const [qrForms, setQrForms] = useState({})
  const [changeRequestDraft, setChangeRequestDraft] = useState(null)
  const [warning, setWarning] = useState('')
  const currentAssignments = getActiveAssignments(data.productionAssignments || [], 'Phối trộn')
  const assignmentEmployeeText = getAssignmentLogContext(currentAssignments).employee
  const activeMixingOrders = orders.filter((order) => order.mixing?.status === 'Active' || order.mixingStatus === 'Active')
  const readyOrders = orders.filter((order) => getOrderAssignedMachineCode(order))
  const mixingHistory = orders.filter((order) => order.mixingStatus === 'completed' || order.mixing?.status === 'Completed' || order.mixingCompletedAt)
  const getMachineActiveOrder = (machineCode) => activeMixingOrders.find((order) => (order.mixingMachine || order.mixing?.machineCode) === machineCode)
  const machineRows = activeMachines.map((machine) => {
    const activeOrder = getMachineActiveOrder(machine.machineCode)
    return { ...machine, status: activeOrder ? 'Đang chạy' : 'Rảnh', activeOrder }
  })
  const machineLabelByCode = (machineCode) => getMixingMachineLabelByCode(machineCode, machines)
  const durationText = (order) => {
    const start = order.mixingStartAt || order.mixing?.startedAt
    const end = order.mixingCompletedAt || order.mixing?.completedAt
    if (!start || !end) return '-'
    const minutes = Math.max(0, Math.round((new Date(end.replace(' ', 'T')) - new Date(start.replace(' ', 'T'))) / 60000))
    return `${minutes} phút`
  }
  const orderCode = (order) => String(order.orderCode || order.id || '')
  const productionOrders = orders.slice().sort((a, b) => {
    if (a.createdAt && b.createdAt && a.createdAt !== b.createdAt) {
      return String(a.createdAt).localeCompare(String(b.createdAt), 'vi', { numeric: true })
    }
    return orderCode(a).localeCompare(orderCode(b), 'vi', { numeric: true })
  })
  const updateQrForm = (orderId, field, value) => {
    setQrForms((current) => ({ ...current, [orderId]: { ...(current[orderId] || {}), [field]: value } }))
  }
  const getQrForm = (order) => ({
    chemicalQr: qrForms[order.id]?.chemicalQr ?? order.mixingQrConfirmation?.chemicalQr ?? '',
    solidQr: qrForms[order.id]?.solidQr ?? order.mixingQrConfirmation?.solidQr ?? '',
  })
  const validateMixingQr = (order, form) => {
    const chemicalContainer = getContainerByQr(data.weighedContainers || [], form.chemicalQr)
    const solidContainer = getContainerByQr(data.weighedContainers || [], form.solidQr)
    const orderCodeValue = order.orderCode || order.id
    const checks = [
      chemicalContainer,
      solidContainer,
      chemicalContainer?.orderId === order.id || chemicalContainer?.orderCode === orderCodeValue,
      solidContainer?.orderId === order.id || solidContainer?.orderCode === orderCodeValue,
      chemicalContainer?.lot === order.lot,
      solidContainer?.lot === order.lot,
      chemicalContainer?.materialGroup === CHEMICAL,
      solidContainer?.materialGroup === SOLID,
      chemicalContainer?.status === 'Đã cân xong',
      solidContainer?.status === 'Đã cân xong',
    ]
    return { pass: checks.every(Boolean), chemicalContainer, solidContainer }
  }
  const confirmMixingQr = (order) => {
    const form = getQrForm(order)
    const result = validateMixingQr(order, form)
    const confirmedAt = nowText()
    setData((current) => {
      const orders = current.orders.map((item) => item.id === order.id ? {
        ...item,
        mixingQrConfirmation: {
          chemicalQr: form.chemicalQr,
          solidQr: form.solidQr,
          status: result.pass ? 'Đã xác nhận' : 'Không đạt',
          confirmedAt,
          note: result.pass ? 'QR hỗn hợp đúng lệnh sản xuất.' : 'QR hỗn hợp không đúng lệnh sản xuất hoặc không đúng nhóm nguyên liệu.',
        },
        updatedAt: confirmedAt,
      } : item)
      const scanLog = `Quét QR hỗn hợp Hóa ${form.chemicalQr || '-'} và QR hỗn hợp Rắn ${form.solidQr || '-'} cho lệnh ${order.orderCode || order.id}.`
      const resultLog = result.pass
        ? `Xác nhận QR hỗn hợp PASS cho lệnh ${order.orderCode || order.id}.`
        : `Xác nhận QR hỗn hợp FAIL cho lệnh ${order.orderCode || order.id}: QR hỗn hợp không đúng lệnh sản xuất hoặc không đúng nhóm nguyên liệu.`
      const meta = operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'Phối trộn', order, result: result.pass ? 'QR PASS' : 'QR FAIL' })
      return addLogToData(addLogToData({ ...current, orders }, scanLog, meta), resultLog, meta)
    })
    setWarning(result.pass ? '' : 'QR hỗn hợp không đúng lệnh sản xuất hoặc không đúng nhóm nguyên liệu.')
  }
  const createDemoQrData = () => {
    setData((current) => applyDemoQrData(current))
    setWarning('Đã tạo dữ liệu demo QR hỗn hợp. Có sẵn 1 lệnh PASS, 1 lệnh sai LOT và 1 lệnh sai nhóm.')
  }
  const getAssignedMachineCode = (order) => getOrderAssignedMachineCode(order)
  const openMachineChangeRequest = (order) => {
    const currentMachine = getAssignedMachineCode(order)
    setChangeRequestDraft({
      orderId: order.id,
      orderCode: order.orderCode || order.id,
      currentMachine,
      requestedMachine: '',
      reason: '',
      requestedBy: 'Tổ phối trộn',
      requestedAt: nowText(),
    })
  }
  const updateChangeRequestDraft = (field, value) => setChangeRequestDraft((current) => ({ ...(current || {}), [field]: value }))
  const submitMachineChangeRequest = () => {
    if (!changeRequestDraft?.orderId) return
    if (!changeRequestDraft.requestedMachine || !changeRequestDraft.reason.trim() || !changeRequestDraft.requestedBy.trim()) {
      setWarning('Vui lòng nhập đầy đủ máy đề nghị, lý do đổi máy và người đề nghị.')
      return
    }
    const requestedAt = changeRequestDraft.requestedAt || nowText()
    setData((current) => addLogToData({
      ...current,
      orders: current.orders.map((item) => item.id === changeRequestDraft.orderId ? {
        ...item,
        machineChangeRequest: {
          requestedMachine: changeRequestDraft.requestedMachine,
          currentMachine: changeRequestDraft.currentMachine,
          reason: changeRequestDraft.reason.trim(),
          status: 'PENDING',
          requestedAt,
          requestedBy: changeRequestDraft.requestedBy.trim(),
        },
        updatedAt: requestedAt,
      } : item),
    }, `Tổ phối trộn đề nghị đổi máy lệnh ${changeRequestDraft.orderCode} từ ${machineLabelByCode(changeRequestDraft.currentMachine)} sang ${machineLabelByCode(changeRequestDraft.requestedMachine)}. Lý do: ${changeRequestDraft.reason.trim()}`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'Phối trộn', order: changeRequestDraft.orderCode, result: 'Đề nghị đổi máy' })))
    setChangeRequestDraft(null)
    setWarning('Đã gửi đề nghị đổi máy, chờ Phòng SX/Admin xác nhận.')
  }
  const startMixing = (order) => {
    const supplement = order.stage === 'mixing-supplement'
    const machineCode = getAssignedMachineCode(order)
    if (!machineCode) {
      setWarning('Chưa được Phòng SX chỉ định máy.')
      return
    }
    if (!activeMachines.some((machine) => machine.machineCode === machineCode)) {
      setWarning(`Máy ${machineLabelByCode(machineCode)} không ở trạng thái READY.`)
      return
    }
    const dispatchState = getMixingDispatchState(order)
    if (!dispatchState.canStart) {
      setWarning(`Lệnh ${order.orderCode || order.id} đang ở trạng thái ${dispatchState.label}.`)
      return
    }
    if (order.mixingQrConfirmation?.status !== 'Đã xác nhận') {
      setWarning('Vui lòng xác nhận QR hỗn hợp Hóa và Rắn trước khi bắt đầu phối trộn.')
      return
    }
    const runningOrder = getMachineActiveOrder(machineCode)
    if (runningOrder) {
      setWarning(`Máy ${machineLabelByCode(machineCode)} đang thực hiện ${runningOrder.orderCode || runningOrder.id}`)
      return
    }
    setWarning('')
    const startedAt = nowText()
    const qrCodes = [order.mixingQrConfirmation?.chemicalQr, order.mixingQrConfirmation?.solidQr]
    setData((current) => addLogToData({
      ...current,
      weighedContainers: updateContainerStatuses(current.weighedContainers || [], qrCodes, 'Đã chuyển phối trộn'),
      mixingMachines: normalizeMixingMachines(current.mixingMachines),
      orders: current.orders.map((item) => item.id === order.id ? {
        ...item,
        stage: supplement ? 'mixing-supplement' : 'mixing',
        status: supplement ? 'Đang phối trộn bổ sung' : 'Đang phối trộn',
        mixingMachine: machineCode,
        mixingStartAt: startedAt,
        mixingCompletedAt: '',
        mixingStatus: 'Active',
        mixing: { ...(item.mixing || {}), status: 'Active', machineCode, startedAt, operator: 'Tổ phối trộn', supplement },
        machineAssignmentHistory: (item.machineAssignmentHistory || []).length ? item.machineAssignmentHistory.map((history, index, list) => (
          index === list.length - 1 ? { ...history, performedMachine: machineCode, performedAt: startedAt } : history
        )) : [{
          id: uid('MCH'),
          orderId: item.id,
          orderCode: item.orderCode || item.id,
          lot: item.lot,
          assignedMachine: machineCode,
          performedMachine: machineCode,
          performedAt: startedAt,
          changedBy: 'Phòng SX',
          changedAt: item.createdAt || startedAt,
          reason: 'Máy chỉ định trên lệnh sản xuất',
        }],
        updatedAt: startedAt,
      } : item),
    }, supplement ? `Bắt đầu phối trộn bổ sung sau khi xác nhận QR hỗn hợp lệnh ${order.id} trên máy ${machineLabelByCode(machineCode)}.` : `Bắt đầu phối trộn sau khi xác nhận QR hỗn hợp lệnh ${order.id} trên máy ${machineLabelByCode(machineCode)}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'Phối trộn', order, result: 'Bắt đầu phối trộn' })))
  }
  const completeMixing = (order, supplement) => setData((current) => addLogToData({
    ...current,
    weighedContainers: updateContainerStatuses(current.weighedContainers || [], [order.mixingQrConfirmation?.chemicalQr, order.mixingQrConfirmation?.solidQr], 'Đã phối trộn'),
    mixingMachines: normalizeMixingMachines(current.mixingMachines),
    orders: current.orders.map((item) => item.id === order.id ? {
      ...item,
      stage: 'finished-qc',
      status: 'Chờ QC thành phẩm',
      mixingStatus: 'completed',
      mixingCompletedAt: nowText(),
      mixing: { ...(item.mixing || {}), status: 'Completed', finalWeightKg: item.quantityKg, completedAt: nowText(), operator: 'Tổ phối trộn', supplement },
      updatedAt: nowText(),
    } : item),
  }, supplement ? `Hoàn thành phối trộn bổ sung ${order.id} và cập nhật QR hỗn hợp đã phối trộn. Chuyển QC thành phẩm.` : `Hoàn thành phối trộn chính ${order.id} và cập nhật QR hỗn hợp đã phối trộn. Chuyển QC thành phẩm.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'Phối trộn', order, result: 'Hoàn thành phối trộn' })))

  return (
    <div className="page-content mixing-page-v2">
      <section className="panel">
        <div className="panel-header-row">
          <div>
            <h2>Tổ phối trộn</h2>
            <p className="panel-text">Điều độ nhiều máy phối trộn chạy song song. Mỗi máy chỉ nhận một lệnh tại một thời điểm.</p>
          </div>
          <button className="primary-button" onClick={createDemoQrData}>Tạo dữ liệu demo QR</button>
        </div>
        {warning && <div className="process-alert">{warning}</div>}
        <div className="process-alert assignment-context">
          <strong>Phân công ca hiện tại:</strong> {currentAssignments.length ? currentAssignments.map((item) => `${formatAssignmentEmployees(item)} (${item.shiftCode}${item.machineName ? ` - ${item.machineName}` : ''})`).join(', ') : 'Chưa có phân công.'}
        </div>
        <div className="table-wrapper mixing-machine-table-wrapper">
          <table className="mixing-machine-table">
            <thead>
              <tr>
                <th>Máy</th>
                <th>Trạng thái</th>
                <th>Lệnh</th>
                <th>Tiến độ</th>
              </tr>
            </thead>
            <tbody>
              {machineRows.map((machine) => {
                const progress = machine.activeOrder ? getMixingProgress(machine.activeOrder) : 0
                return (
                  <tr key={machine.machineCode}>
                    <td>{formatMixingMachineLabel(machine)}</td>
                    <td><span className={`dispatch-badge ${machine.activeOrder ? 'mixing' : 'ready'}`}>{machine.status}</span></td>
                    <td>{machine.activeOrder ? (machine.activeOrder.orderCode || machine.activeOrder.id) : '-'}</td>
                    <td><div className="mix-progress mixing-progress"><div className="mixing-progress-bar"><i style={{ width: `${progress}%` }} /></div><strong>{progress}%</strong></div></td>
                  </tr>
                )
              })}
            </tbody>
          </table>
        </div>
      </section>
      <section className="mixing-sections-grid">
        <section className="panel">
          <h3>Máy đang phối trộn</h3>
          <div className="table-wrapper mixing-active-table-wrapper">
            <table className="mixing-active-table">
              <thead>
                <tr>
                  <th>Lệnh</th>
                  <th>Sản phẩm</th>
                  <th>LOT</th>
                  <th>Máy phối trộn</th>
                  <th>Giờ bắt đầu</th>
                  <th>Trạng thái</th>
                  <th>Hành động</th>
                </tr>
              </thead>
              <tbody>
                {activeMixingOrders.map((order) => {
                  const supplement = order.stage === 'mixing-supplement' || order.mixing?.supplement
                  return (
                    <tr key={order.id}>
                      <td>{order.orderCode || order.id}</td>
                      <td>{order.productName || order.product}</td>
                      <td>{order.lot}</td>
                      <td>{getOrderAssignedMachineLabel(order, machines)}</td>
                      <td>{order.mixingStartAt || order.mixing?.startedAt || '-'}</td>
                      <td><span className="dispatch-badge mixing">Đang phối trộn</span></td>
                      <td><button className="secondary-button" onClick={() => completeMixing(order, supplement)}>{supplement ? 'Hoàn thành bổ sung' : 'Hoàn thành phối trộn'}</button></td>
                    </tr>
                  )
                })}
                {activeMixingOrders.length === 0 && <tr><td className="empty-row" colSpan="7">Không có máy đang phối trộn.</td></tr>}
              </tbody>
            </table>
          </div>
        </section>
        <section className="panel">
          <h3>Lệnh sẵn sàng phối trộn</h3>
          <div className="ready-mixing-table-wrapper">
            <table className="ready-mixing-table">
              <thead>
                <tr>
                  <th>Lệnh</th>
                  <th>Sản phẩm</th>
                  <th>LOT</th>
                  <th>Loại phối trộn</th>
                  <th><span>QR hỗn hợp</span><span>Hóa</span></th>
                  <th><span>QR hỗn hợp</span><span>Rắn</span></th>
                  <th><span>Xác nhận</span><span>QR</span></th>
                  <th><span>Máy Phòng SX</span><span>chỉ định</span></th>
                  <th>Trạng thái</th>
                  <th>Hành động</th>
                </tr>
              </thead>
              <tbody>
                {readyOrders.map((order) => {
                  const supplement = order.stage === 'mixing-supplement' || order.mixing?.supplement
                  const qrForm = getQrForm(order)
                  const qrStatus = order.mixingQrConfirmation?.status || 'Chưa xác nhận'
                  const qrConfirmed = qrStatus === 'Đã xác nhận'
                  const assignedMachineCode = getAssignedMachineCode(order)
                  const assignedMachine = machines.find((machine) => machine.machineCode === assignedMachineCode)
                  const pendingChange = order.machineChangeRequest?.status === 'PENDING'
                  const state = getMixingDispatchState(order)
                  return (
                    <tr key={order.id}>
                      <td>{order.orderCode || order.id}</td>
                      <td>{order.productName || order.product}</td>
                      <td>{order.lot}</td>
                      <td>{supplement ? 'Phối trộn bổ sung QC2' : 'Phối trộn chính'}</td>
                      <td><input className="mixing-qr-input qr-input" value={qrForm.chemicalQr} onChange={(event) => updateQrForm(order.id, 'chemicalQr', event.target.value)} /></td>
                      <td><input className="mixing-qr-input qr-input" value={qrForm.solidQr} onChange={(event) => updateQrForm(order.id, 'solidQr', event.target.value)} /></td>
                      <td><div className="mixing-qr-confirm"><button className="secondary-button confirm-qr-button" onClick={() => confirmMixingQr(order)}>Xác nhận QR hỗn hợp</button><span className={`dispatch-badge ${qrConfirmed ? 'ready' : qrStatus === 'Không đạt' ? 'fail' : 'waiting'}`}>{qrStatus}</span></div></td>
                      <td>
                        <div className="mixing-qr-confirm">
                          {assignedMachineCode ? <strong>{assignedMachine ? formatMixingMachineLabel(assignedMachine) : getOrderAssignedMachineLabel(order, machines)}</strong> : <span className="process-alert">Chưa được Phòng SX chỉ định máy</span>}
                          <button type="button" className="secondary-button" onClick={() => openMachineChangeRequest(order)}>Đề nghị đổi máy</button>
                          {pendingChange && <span className="dispatch-badge waiting">Chờ xác nhận đổi máy</span>}
                        </div>
                      </td>
                      <td><span className={`dispatch-badge ${state.className}`}>{state.label}</span></td>
                      <td><button className="primary-button start-mixing-button" disabled={!state.canStart || !qrConfirmed || !assignedMachineCode || pendingChange} onClick={() => startMixing(order)}>{supplement ? 'Bắt đầu phối trộn bổ sung QC2' : 'Bắt đầu phối trộn'}</button></td>
                    </tr>
                  )
                })}
                {readyOrders.length === 0 && <tr><td className="empty-row" colSpan="10">Không có lệnh sẵn sàng phối trộn.</td></tr>}
              </tbody>
            </table>
          </div>
        </section>
        <section className="panel mixing-dispatch">
          <h3>Danh sách lệnh sản xuất</h3>
          <SimpleTable headers={['Lệnh', 'Sản phẩm', 'LOT', 'Máy phối trộn', 'Giờ bắt đầu', 'Giờ kết thúc', 'Thời gian phối trộn', 'Trạng thái']} rows={productionOrders.map((order) => {
            const state = getMixingDispatchState(order)
            return (
              <tr key={order.id}>
                <td>{order.orderCode || order.id}</td>
                <td>{order.productName || order.product}</td>
                <td>{order.lot}</td>
                <td>{getOrderAssignedMachineLabel(order, machines)}</td>
                <td>{order.mixingStartAt || order.mixing?.startedAt || '-'}</td>
                <td>{order.mixingCompletedAt || order.mixing?.completedAt || '-'}</td>
                <td>{durationText(order)}</td>
                <td><span className={`dispatch-badge ${state.className}`}>{state.label}</span></td>
              </tr>
            )
          })} />
        </section>
        <section className="panel mixing-history-panel">
          <h3>Lịch sử phối trộn</h3>
          <SimpleTable headers={['Lệnh', 'Sản phẩm', 'Máy phối trộn', 'Giờ bắt đầu', 'Giờ kết thúc', 'Thời gian phối trộn', 'Trạng thái']} rows={mixingHistory.slice().reverse().map((order) => (
            <tr key={order.id}>
              <td>{order.orderCode || order.id}</td>
              <td>{order.productName || order.product}</td>
              <td>{getOrderAssignedMachineLabel(order, machines)}</td>
              <td>{order.mixingStartAt || order.mixing?.startedAt || '-'}</td>
              <td>{order.mixingCompletedAt || order.mixing?.completedAt || '-'}</td>
              <td>{durationText(order)}</td>
              <td><span className="dispatch-badge done">Hoàn thành</span></td>
            </tr>
          ))} empty="Chưa có lịch sử phối trộn." />
        </section>
      </section>
      {changeRequestDraft && (
        <div className="modal-backdrop" role="presentation">
          <div className="mixing-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><span className="section-kicker">Đề nghị đổi máy</span><h2>{changeRequestDraft.orderCode}</h2></div>
              <button type="button" className="icon-button" onClick={() => setChangeRequestDraft(null)} aria-label="Đóng">×</button>
            </div>
            <div className="production-form-grid order-create-form">
              <label>Máy hiện tại<input value={changeRequestDraft.currentMachine ? machineLabelByCode(changeRequestDraft.currentMachine) : 'Chưa được Phòng SX chỉ định máy'} disabled /></label>
              <label>Máy đề nghị<select value={changeRequestDraft.requestedMachine} onChange={(event) => updateChangeRequestDraft('requestedMachine', event.target.value)}><option value="">Chọn máy</option>{activeMachines.map((machine) => <option key={machine.machineCode} value={machine.machineCode}>{mixingMachineOptionLabel(machine)}</option>)}</select></label>
              <label>Người đề nghị<input value={changeRequestDraft.requestedBy} onChange={(event) => updateChangeRequestDraft('requestedBy', event.target.value)} /></label>
              <label>Thời gian đề nghị<input value={changeRequestDraft.requestedAt} disabled /></label>
              <label className="wide-field">Lý do đổi máy<textarea value={changeRequestDraft.reason} onChange={(event) => updateChangeRequestDraft('reason', event.target.value)} /></label>
            </div>
            <div className="modal-actions">
              <button type="button" className="secondary-button" onClick={() => setChangeRequestDraft(null)}>Hủy</button>
              <button type="button" className="primary-button" onClick={submitMachineChangeRequest}>Gửi đề nghị</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function PackagingPage({ data, setData, user }) {
  const orders = data.orders
    .filter((order) => (
      order.stage === 'packaging'
      || ['Chờ đóng gói', 'Đang đóng gói', 'Đóng gói hoàn thành', 'Chờ nhập kho thành phẩm'].includes(order.status)
      || ['pending', 'active', 'completed', 'Pending', 'Active', 'Completed'].includes(order.packingStatus || order.packagingStatus)
    ))
    .slice()
    .sort(sortOldestOrders)
  const [activeOrderId, setActiveOrderId] = useState(orders[0]?.id || '')
  const [forms, setForms] = useState({})
  const [warning, setWarning] = useState('')
  const activeOrder = orders.find((order) => order.id === activeOrderId) || orders[0]
  const form = activeOrder ? getPackingForm(forms, activeOrder) : { details: defaultPackingDetails() }
  const totals = activeOrder ? packingTotals(activeOrder, form) : { qcWeight: 0, totalPackedWeight: 0, remainingWeight: 0, differenceWeight: 0, totalTolerance: 0 }
  const currentAssignments = getActiveAssignments(data.productionAssignments || [], 'Đóng gói')
  const assignmentEmployeeText = getAssignmentLogContext(currentAssignments).employee

  const updateForm = (patch) => {
    if (!activeOrder) return
    setForms((current) => ({ ...current, [activeOrder.id]: { ...getPackingForm(current, activeOrder), ...patch } }))
  }
  const updateDetail = (id, field, value) => {
    if (!activeOrder) return
    setForms((current) => {
      const nextForm = getPackingForm(current, activeOrder)
      return {
        ...current,
        [activeOrder.id]: {
          ...nextForm,
          details: nextForm.details.map((item) => item.id === id ? { ...item, [field]: field === 'note' ? value : Number(value) || 0 } : item),
        },
      }
    })
  }
  const buildPackingLog = (order, nextForm, status) => {
    const nextTotals = packingTotals(order, nextForm)
    return {
      packingId: uid('PKG'),
      orderId: order.id,
      orderCode: order.orderCode || order.id,
      productName: order.productName || order.product,
      lot: order.lot,
      qc2FinalWeight: nextTotals.qcWeight,
      packingDetails: nextForm.details.map((item) => ({
        spec: item.label,
        sizeKg: item.sizeKg,
        boxes: num(item.boxes),
        convertedWeight: Number((num(item.sizeKg) * num(item.boxes)).toFixed(3)),
        toleranceKg: item.toleranceKg,
        actualWeight: num(item.actualWeight),
        note: item.note || '',
      })),
      totalPackedWeight: nextTotals.totalPackedWeight,
      remainingWeight: nextTotals.remainingWeight,
      differenceWeight: nextTotals.differenceWeight,
      packer: nextForm.packer,
      startedAt: nextForm.startedAt,
      completedAt: nextForm.completedAt,
      status,
      notes: nextForm.notes,
    }
  }
  const beginPackaging = () => {
    if (!activeOrder) return
    const startedAt = form.startedAt || nowText()
    const nextForm = { ...form, startedAt }
    setForms((current) => ({ ...current, [activeOrder.id]: nextForm }))
    setData((current) => addLogToData({
      ...current,
      orders: current.orders.map((order) => order.id === activeOrder.id ? {
        ...order,
        stage: 'packaging',
        status: 'Đang đóng gói',
        packingStatus: 'active',
        packagingStatus: 'Active',
        packaging: nextForm,
        updatedAt: startedAt,
      } : order),
    }, `Bắt đầu đóng gói lệnh ${activeOrder.id}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'Đóng gói', order: activeOrder, result: 'Bắt đầu đóng gói' })))
    setWarning('')
  }
  const saveDraft = () => {
    if (!activeOrder) return
    const nextForm = getPackingForm(forms, activeOrder)
    setData((current) => addLogToData({
      ...current,
      orders: current.orders.map((order) => order.id === activeOrder.id ? {
        ...order,
        stage: 'packaging',
        status: order.status === 'Chờ đóng gói' ? 'Đang đóng gói' : order.status,
        packingStatus: order.packingStatus === 'completed' ? 'completed' : 'active',
        packagingStatus: order.packagingStatus === 'Completed' ? 'Completed' : 'Active',
        packaging: nextForm,
        updatedAt: nowText(),
      } : order),
    }, `Lưu tạm đóng gói lệnh ${activeOrder.id}.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'Đóng gói', order: activeOrder, result: 'Lưu tạm đóng gói' })))
    setWarning('')
  }
  const completePackaging = () => {
    if (!activeOrder) return
    const completedAt = form.completedAt || nowText()
    const nextForm = { ...form, completedAt }
    const nextTotals = packingTotals(activeOrder, nextForm)
    if (Math.abs(nextTotals.differenceWeight) > nextTotals.totalTolerance) {
      setWarning(`Sai lệch ${kg(nextTotals.differenceWeight)} vượt sai số tổng cho phép ${kg(nextTotals.totalTolerance)}.`)
      return
    }
    const packingLog = buildPackingLog(activeOrder, nextForm, 'completed')
    setForms((current) => ({ ...current, [activeOrder.id]: nextForm }))
    setData((current) => addLogToData({
      ...current,
      orders: current.orders.map((order) => order.id === activeOrder.id ? {
        ...order,
        stage: 'finished-goods',
        status: 'Chờ nhập kho thành phẩm',
        orderStatus: 'Chờ nhập kho thành phẩm',
        packingStatus: 'completed',
        packagingStatus: 'Completed',
        packaging: { ...nextForm, ...nextTotals, packingLogId: packingLog.packingId },
        updatedAt: completedAt,
      } : order),
      packingLogs: [...(current.packingLogs || []), packingLog],
    }, `Hoàn tất đóng gói lệnh ${activeOrder.id}, chuyển Chờ nhập kho thành phẩm.`, operationLogMeta(user, { assignments: currentAssignments, employee: assignmentEmployeeText, stage: 'Đóng gói', order: activeOrder, result: 'Hoàn tất đóng gói' })))
    setWarning('')
  }

  return (
    <div className="page-content packaging-page">
      <section className="panel packaging-layout">
        <aside className="packaging-list">
          <h2>Danh sách lệnh chờ đóng gói</h2>
          <SimpleTable headers={['Mã lệnh SX', 'Sản phẩm', 'LOT', 'Khối lượng sau QC thành phẩm', 'Trạng thái', 'Hành động']} rows={orders.map((order) => (
            <tr key={order.id} className={order.id === activeOrder?.id ? 'current-row' : ''}>
              <td>{order.orderCode || order.id}</td>
              <td>{order.productName || order.product}</td>
              <td>{order.lot}</td>
              <td>{kg(qc2FinalWeight(order))}</td>
              <td><span className={`dispatch-badge ${order.packingStatus === 'completed' || order.packagingStatus === 'Completed' ? 'done' : order.status === 'Đang đóng gói' ? 'mixing' : 'waiting'}`}>{order.status}</span></td>
              <td><button className="secondary-button" onClick={() => { setActiveOrderId(order.id); setWarning('') }}>Chọn</button></td>
            </tr>
          ))} empty="Không có lệnh chờ đóng gói." />
        </aside>

        <main className="packaging-detail">
          {!activeOrder && <p className="empty-alert">Chưa có lệnh đóng gói.</p>}
          {activeOrder && (
            <>
              <div className="section-heading-row">
                <div>
                  <span className="section-kicker">Chi tiết lệnh đang đóng gói</span>
                  <h2>{activeOrder.orderCode || activeOrder.id}</h2>
                </div>
                <div className="action-row touch-actions">
                  <button className="secondary-button touch-button" onClick={beginPackaging}>Bắt đầu đóng gói</button>
                  <button className="secondary-button touch-button" onClick={saveDraft}>Lưu tạm</button>
                  <button className="primary-button touch-button" onClick={completePackaging}>Hoàn tất đóng gói</button>
                </div>
              </div>
              {warning && <div className="process-alert">{warning}</div>}
              <div className="qc-order-summary">
                <div><span>Mã lệnh SX</span><strong>{activeOrder.orderCode || activeOrder.id}</strong></div>
                <div><span>Sản phẩm</span><strong>{activeOrder.productName || activeOrder.product}</strong></div>
                <div><span>LOT</span><strong>{activeOrder.lot}</strong></div>
                <div><span>Khối lượng sau QC thành phẩm</span><strong>{kg(totals.qcWeight)}</strong></div>
                <div><span>Khối lượng đã đóng gói</span><strong>{kg(totals.totalPackedWeight)}</strong></div>
                <div><span>Khối lượng còn lại</span><strong>{kg(totals.remainingWeight)}</strong></div>
                <div><span>Sai lệch</span><strong>{kg(totals.differenceWeight)}</strong></div>
                <div><span>Sai số tổng cho phép</span><strong>{kg(totals.totalTolerance)}</strong></div>
              </div>
              <section className="v3-card">
                <h3>Bảng quy cách đóng gói</h3>
                <SimpleTable headers={['Quy cách', 'Số thùng', 'Khối lượng quy đổi', 'Sai số cho phép', 'Khối lượng thực tế', 'Ghi chú']} rows={form.details.map((item) => (
                  <tr key={item.id}>
                    <td>{item.label}</td>
                    <td><input type="number" value={item.boxes} onChange={(event) => updateDetail(item.id, 'boxes', event.target.value)} /></td>
                    <td>{kg(num(item.sizeKg) * num(item.boxes))}</td>
                    <td>±{kg(num(item.toleranceKg) * num(item.boxes))}</td>
                    <td><input type="number" step="0.001" value={item.actualWeight} onChange={(event) => updateDetail(item.id, 'actualWeight', event.target.value)} /></td>
                    <td><input value={item.note} onChange={(event) => updateDetail(item.id, 'note', event.target.value)} /></td>
                  </tr>
                ))} />
              </section>
              <section className="v3-card">
                <h3>Thông tin đóng gói</h3>
                <div className="production-form-grid">
                  <label>Người đóng gói<input value={form.packer} onChange={(event) => updateForm({ packer: event.target.value })} /></label>
                  <label>Thời gian bắt đầu đóng gói<input value={form.startedAt} onChange={(event) => updateForm({ startedAt: event.target.value })} /></label>
                  <label>Thời gian hoàn thành đóng gói<input value={form.completedAt} onChange={(event) => updateForm({ completedAt: event.target.value })} /></label>
                  <label className="wide-field">Ghi chú đóng gói<input value={form.notes} onChange={(event) => updateForm({ notes: event.target.value })} /></label>
                </div>
              </section>
            </>
          )}
        </main>
      </section>
    </div>
  )
}

function FinishedGoodsPage({ data, setData }) {
  const packingLogs = data.packingLogs || []
  const finishedGoods = normalizeFinishedGoodsData(data.finishedGoods || [])
  const waitingOrders = data.orders
    .filter((order) => (
      (order.packingStatus === 'completed' || order.packagingStatus === 'Completed')
      && !['completed', 'Completed'].includes(order.finishedGoodsStatus)
      && (order.orderStatus === 'Chờ nhập kho thành phẩm' || order.status === 'Chờ nhập kho thành phẩm' || order.stage === 'finished-goods')
    ))
    .slice()
    .sort(sortOldestOrders)
  const [activeOrder, setActiveOrder] = useState(null)
  const [filters, setFilters] = useState({ fromDate: '', toDate: '', orderCode: '', product: '', lot: '', location: '' })
  const [form, setForm] = useState(null)
  const [notice, setNotice] = useState('')
  const filteredFinishedGoods = filterFinishedGoods(finishedGoods, filters)

  const createDemoData = () => {
    const payload = createFinishedGoodsDemoPayload(data)
    setData((current) => ({
      ...current,
      orders: [...(current.orders || []), ...payload.orders],
      packingLogs: [...(current.packingLogs || []), ...payload.packingLogs],
      finishedGoods: [...normalizeFinishedGoodsData(current.finishedGoods || []), ...payload.finishedGoods],
      productionLogs: [...(current.productionLogs || current.logs || []), ...payload.productionLogs],
      logs: [...(current.logs || []), ...payload.productionLogs],
    }))
    setNotice('Đã tạo dữ liệu demo kho TP')
  }

  const openImportForm = (order) => {
    const packingLog = getLatestPackingLog(order, packingLogs)
    setActiveOrder(order)
    setForm({
      finishedCode: nextFinishedCode(finishedGoods),
      orderCode: order.orderCode || order.id,
      productName: order.productName || order.product,
      lot: order.lot,
      weight: packingLog?.totalPackedWeight || order.packaging?.totalPackedWeight || qc2FinalWeight(order),
      boxes: packingLog ? totalPackingBoxes(packingLog.packingDetails) : totalPackingBoxes(order.packaging?.details || []),
      spec: packingLog ? packingSpecSummary(packingLog.packingDetails) : packingSpecSummary(order.packaging?.details || []),
      location: 'KTP-A01',
      receiver: '',
      importDate: todayText(),
      note: '',
    })
  }
  const updateForm = (field, value) => setForm((current) => ({ ...current, [field]: ['weight', 'boxes'].includes(field) ? value : value }))
  const completeImport = () => {
    if (!activeOrder || !form) return
    const completedAt = nowText()
    const finishedItem = {
      id: uid('FG'),
      finishedCode: form.finishedCode,
      orderId: activeOrder.id,
      orderCode: form.orderCode,
      productName: form.productName,
      product: form.productName,
      lot: form.lot,
      spec: form.spec,
      boxes: num(form.boxes),
      weight: num(form.weight),
      importDate: form.importDate,
      location: form.location,
      receiver: form.receiver,
      note: form.note,
      status: 'Hoàn thành',
    }
    setData((current) => addLogToData({
      ...current,
      finishedGoods: [finishedItem, ...normalizeFinishedGoodsData(current.finishedGoods || [])],
      orders: current.orders.map((order) => order.id === activeOrder.id ? {
        ...order,
        stage: 'completed',
        status: 'Hoàn thành',
        orderStatus: 'Hoàn thành',
        finishedGoodsStatus: 'completed',
        finishedGoodsCode: form.finishedCode,
        completedAt,
        updatedAt: completedAt,
      } : order),
    }, `Nhập kho thành phẩm ${activeOrder.id}, mã ${form.finishedCode}. Hoàn thành lệnh SX.`))
    setActiveOrder(null)
    setForm(null)
  }
  const exportRows = filteredFinishedGoods.map((item) => ({
    'Mã TP': item.finishedCode,
    'Mã lệnh SX': item.orderCode || item.orderId,
    'Sản phẩm': item.productName || item.product,
    LOT: item.lot,
    'Quy cách': item.spec,
    'Số thùng': item.boxes,
    'Khối lượng': item.weight,
    'Ngày nhập': item.importDate,
    'Vị trí kho': item.location,
    'Người nhập': item.receiver,
    'Trạng thái': item.status,
  }))
  const exportExcel = () => {
    const sheet = XLSX.utils.json_to_sheet(exportRows)
    const book = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(book, sheet, 'Kho thanh pham')
    XLSX.writeFile(book, 'kho-thanh-pham.xlsx')
  }
  const exportPdf = () => {
    const doc = new jsPDF({ orientation: 'landscape' })
    doc.text('Bao cao kho thanh pham', 14, 14)
    autoTable(doc, {
      head: [['Ma TP', 'Lenh SX', 'San pham', 'LOT', 'Quy cach', 'So thung', 'Khoi luong', 'Ngay nhap', 'Vi tri', 'Nguoi nhap', 'Trang thai']],
      body: filteredFinishedGoods.map((item) => [item.finishedCode, item.orderCode || item.orderId, item.productName || item.product, item.lot, item.spec, item.boxes, item.weight, item.importDate, item.location, item.receiver, item.status]),
      startY: 20,
    })
    doc.save('kho-thanh-pham.pdf')
  }

  return (
    <div className="page-content finished-goods-page">
      <section className="panel">
        <div className="section-heading-row">
          <h2>Danh sách lệnh chờ nhập kho TP</h2>
          <button className="primary-button" onClick={createDemoData}>Tạo dữ liệu demo kho TP</button>
        </div>
        {notice && <div className="process-alert success-alert">{notice}</div>}
        <div className="finished-warehouse-table-wrapper warehouse-table-wrapper">
          <table className="finished-warehouse-table warehouse-table">
            <thead>
              <tr>
                <th>Mã lệnh SX</th>
                <th>Sản phẩm</th>
                <th>LOT</th>
                <th>Khách hàng</th>
                <th><span>Khối lượng sau</span><span>QC thành phẩm</span></th>
                <th><span>Tổng khối lượng</span><span>đã đóng gói</span></th>
                <th><span>Quy cách</span><span>đóng gói</span></th>
                <th>Số thùng</th>
                <th>Người đóng gói</th>
                <th><span>Thời gian</span><span>hoàn thành</span></th>
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {waitingOrders.map((order) => {
                const packingLog = getLatestPackingLog(order, packingLogs)
                const details = packingLog?.packingDetails || order.packaging?.details || []
                return (
                  <tr key={order.id}>
                    <td>{order.orderCode || order.id}</td>
                    <td>{order.productName || order.product}</td>
                    <td>{order.lot}</td>
                    <td>{order.customer || '-'}</td>
                    <td>{kg(qc2FinalWeight(order))}</td>
                    <td>{kg(packingLog?.totalPackedWeight || order.packaging?.totalPackedWeight || 0)}</td>
                    <td>{packingSpecSummary(details)}</td>
                    <td>{totalPackingBoxes(details)}</td>
                    <td>{packingLog?.packer || order.packaging?.packer || '-'}</td>
                    <td>{packingLog?.completedAt || order.packaging?.completedAt || '-'}</td>
                    <td><button className="primary-button warehouse-import-button" onClick={() => openImportForm(order)}>Nhập kho thành phẩm</button></td>
                  </tr>
                )
              })}
              {waitingOrders.length === 0 && <tr><td className="empty-row" colSpan="11">Không có lệnh chờ nhập kho thành phẩm.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>

      <section className="panel">
        <div className="section-heading-row">
          <h2>Thành phẩm đã nhập kho</h2>
          <div className="action-row">
            <button className="secondary-button" onClick={exportExcel}>Tải Excel</button>
            <button className="secondary-button" onClick={exportPdf}>Tải PDF</button>
          </div>
        </div>
        <div className="production-form-grid finished-goods-filters">
          <label>Từ ngày<input type="date" value={filters.fromDate} onChange={(event) => setFilters({ ...filters, fromDate: event.target.value })} /></label>
          <label>Đến ngày<input type="date" value={filters.toDate} onChange={(event) => setFilters({ ...filters, toDate: event.target.value })} /></label>
          <label>Mã lệnh SX<input value={filters.orderCode} onChange={(event) => setFilters({ ...filters, orderCode: event.target.value })} /></label>
          <label>Sản phẩm<input value={filters.product} onChange={(event) => setFilters({ ...filters, product: event.target.value })} /></label>
          <label>LOT<input value={filters.lot} onChange={(event) => setFilters({ ...filters, lot: event.target.value })} /></label>
          <label>Vị trí kho<input value={filters.location} onChange={(event) => setFilters({ ...filters, location: event.target.value })} /></label>
        </div>
        <div className="finished-goods-table-wrapper warehouse-table-wrapper">
          <table className="finished-goods-table warehouse-table">
            <thead>
              <tr>
                <th>Mã TP</th>
                <th>Mã lệnh SX</th>
                <th>Sản phẩm</th>
                <th>LOT</th>
                <th>Quy cách</th>
                <th>Số thùng</th>
                <th>Khối lượng</th>
                <th>Ngày nhập</th>
                <th>Vị trí kho</th>
                <th>Người nhập</th>
                <th>Trạng thái</th>
              </tr>
            </thead>
            <tbody>
              {filteredFinishedGoods.map((item) => (
                <tr key={item.id}>
                  <td>{item.finishedCode}</td>
                  <td>{item.orderCode || item.orderId}</td>
                  <td>{item.productName || item.product}</td>
                  <td>{item.lot}</td>
                  <td>{item.spec}</td>
                  <td>{item.boxes}</td>
                  <td>{kg(item.weight)}</td>
                  <td>{item.importDate}</td>
                  <td>{item.location}</td>
                  <td>{item.receiver || '-'}</td>
                  <td>{item.status}</td>
                </tr>
              ))}
              {filteredFinishedGoods.length === 0 && <tr><td className="empty-row" colSpan="11">Chưa có thành phẩm nhập kho.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>

      {form && activeOrder && (
        <div className="modal-backdrop" role="presentation">
          <div className="mixing-modal finished-goods-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><span className="section-kicker">Nhập kho thành phẩm</span><h2>{activeOrder.orderCode || activeOrder.id}</h2></div>
              <button type="button" className="icon-button" onClick={() => { setForm(null); setActiveOrder(null) }} aria-label="Đóng">×</button>
            </div>
            <div className="production-form-grid">
              <label>Mã thành phẩm<input value={form.finishedCode} onChange={(event) => updateForm('finishedCode', event.target.value)} /></label>
              <label>Mã lệnh SX<input readOnly value={form.orderCode} /></label>
              <label>Sản phẩm<input readOnly value={form.productName} /></label>
              <label>LOT<input readOnly value={form.lot} /></label>
              <label>Khối lượng nhập kho<input type="number" value={form.weight} onChange={(event) => updateForm('weight', event.target.value)} /></label>
              <label>Số thùng nhập kho<input type="number" value={form.boxes} onChange={(event) => updateForm('boxes', event.target.value)} /></label>
              <label>Quy cách<input value={form.spec} onChange={(event) => updateForm('spec', event.target.value)} /></label>
              <label>Vị trí kho<input value={form.location} onChange={(event) => updateForm('location', event.target.value)} /></label>
              <label>Người nhập kho<input value={form.receiver} onChange={(event) => updateForm('receiver', event.target.value)} /></label>
              <label>Ngày nhập kho<input type="date" value={form.importDate} onChange={(event) => updateForm('importDate', event.target.value)} /></label>
              <label className="wide-field">Ghi chú<input value={form.note} onChange={(event) => updateForm('note', event.target.value)} /></label>
            </div>
            <div className="modal-actions">
              <button type="button" className="secondary-button" onClick={() => { setForm(null); setActiveOrder(null) }}>Hủy</button>
              <button type="button" className="primary-button" onClick={completeImport}>Hoàn tất nhập kho</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

function DashboardPage({ data }) {
  const [filters, setFilters] = useState({ fromDate: '', toDate: '', productGroup: 'all', machine: 'all', stage: 'all' })
  const updateFilter = (field, value) => setFilters((current) => ({ ...current, [field]: value }))
  const machines = normalizeMixingMachines(data.mixingMachines)
  const activeMachines = getActiveMixingMachines(machines)
  const getOrderGroup = (order = {}) => String(order.productName || order.product || 'Khác').split(/[.\s-]/)[0] || 'Khác'
  const productGroups = Array.from(new Set((data.orders || []).map(getOrderGroup))).filter(Boolean)
  const stageOptions = [
    ['qc1', 'QC sản xuất thử'],
    ['weighing', 'Cân'],
    ['mixing', 'Phối trộn'],
    ['finished-qc', 'QC thành phẩm'],
    ['packaging', 'Đóng gói'],
    ['finished-goods', 'Kho TP'],
    ['completed', 'Hoàn thành'],
  ]
  const filteredOrders = (data.orders || []).filter((order) => {
    const dateText = String(order.updatedAt || order.createdAt || '').slice(0, 10)
    const machineCode = order.mixingMachine || order.mixing?.machineCode || ''
    if (filters.fromDate && dateText && dateText < filters.fromDate) return false
    if (filters.toDate && dateText && dateText > filters.toDate) return false
    if (filters.productGroup !== 'all' && getOrderGroup(order) !== filters.productGroup) return false
    if (filters.machine !== 'all' && machineCode !== filters.machine) return false
    if (filters.stage !== 'all' && order.stage !== filters.stage) return false
    return true
  })
  const runningMixingOrders = filteredOrders.filter((o) => o.mixing?.status === 'Active' || o.mixingStatus === 'Active')
  const runningMachines = activeMachines.filter((machine) => runningMixingOrders.some((order) => (order.mixingMachine || order.mixing?.machineCode) === machine.machineCode)).length
  const qc2AdjustmentRows = filteredOrders.flatMap((order) => getQc2Adjustments(order).map((ticket) => ({ order, ticket })))
  const today = todayText()
  const finishedGoods = normalizeFinishedGoodsData(data.finishedGoods || []).filter((item) => (
    filteredOrders.some((order) => order.id === item.orderId || order.orderCode === item.orderCode)
  ))
  const qc2Orders = filteredOrders.filter((o) => o.qc2 || getQc2Adjustments(o).length > 0 || ['finished-qc', 'packaging', 'finished-goods', 'completed'].includes(o.stage))
  const qc2FirstPass = qc2Orders.filter((o) => o.qc2?.result === 'OK' && getQc2Adjustments(o).length === 0).length
  const firstPassRate = qc2Orders.length ? Math.round((qc2FirstPass / qc2Orders.length) * 100) : 0
  const delayedOrders = filteredOrders.filter((order) => ['qc1', 'weighing', 'mixing', 'finished-qc', 'packaging', 'finished-goods'].includes(order.stage) && String(order.createdAt || '').slice(0, 10) < today).length
  const stoppedMachines = machines.filter((machine) => ['Dừng', 'Bảo trì', 'Lỗi', 'Stopped'].includes(machine.status)).length
  const repeatedFormulaAdjustments = countBy(qc2AdjustmentRows, ({ order }) => order.formulaCode || order.originalFormulaId).filter(([, value]) => value >= 2).length
  const plannedOutput = filteredOrders.reduce((sum, order) => sum + num(order.quantityKg || order.requestedWeight), 0)
  const actualOutput = finishedGoods.reduce((sum, item) => sum + num(item.weight), 0) || filteredOrders.filter((order) => ['completed', 'finished-goods', 'packaging'].includes(order.stage)).reduce((sum, order) => sum + num(order.quantityKg || order.requestedWeight), 0)
  const completionRate = plannedOutput ? Math.round((actualOutput / plannedOutput) * 100) : 0
  const capacityUse = activeMachines.length ? Math.round((runningMachines / activeMachines.length) * 100) : 0
  const oee = Math.round((capacityUse / 100) * Math.min(completionRate, 100) * (firstPassRate / 100))
  const packingLogs = data.packingLogs || []
  const materialLossKg = Math.abs(packingLogs.reduce((sum, log) => sum + num(log.differenceWeight), 0)) + qc2AdjustmentRows.reduce((sum, { ticket }) => sum + getAdjustmentItems(ticket).reduce((lineSum, item) => lineSum + Math.max(0, num(item.adjustmentKg ?? item.requiredKg)), 0), 0)
  const productionValue = actualOutput * 45000
  const inventoryValue = finishedGoods.reduce((sum, item) => sum + num(item.weight) * 45000, 0) + (data.rawMaterials || []).reduce((sum, item) => sum + num(item.weight) * 18000, 0)
  const lossLimitKg = Math.max(1, plannedOutput * 0.015)
  const productionKpis = [
    ['Sản lượng thực tế', kg(actualOutput), 'normal'],
    ['% hoàn thành kế hoạch', `${completionRate}%`, completionRate >= 90 ? 'normal' : 'watch'],
    ['Công suất sử dụng', `${capacityUse}%`, 'normal'],
    ['OEE', `${oee}%`, oee >= 60 ? 'normal' : 'watch'],
    ['Hao hụt nguyên liệu', kg(materialLossKg), materialLossKg > lossLimitKg ? 'risk' : 'normal'],
    ['Pass QC lần đầu', `${firstPassRate}%`, firstPassRate >= 85 ? 'normal' : 'watch'],
    ['Giá trị sản xuất', `${Math.round(productionValue / 1000000).toLocaleString('vi-VN')} tr`, 'normal'],
    ['Giá trị tồn kho', `${Math.round(inventoryValue / 1000000).toLocaleString('vi-VN')} tr`, 'normal'],
  ]
  const alertKpis = [
    ['Chậm tiến độ', delayedOrders, 'risk'],
    ['Máy dừng bất thường', stoppedMachines, 'risk'],
    ['Hao hụt vượt định mức', materialLossKg > lossLimitKg ? 1 : 0, materialLossKg > lossLimitKg ? 'risk' : 'normal'],
    ['QC phải điều chỉnh nhiều lần', repeatedFormulaAdjustments, 'watch'],
  ]
  const outputByTime = countBy(filteredOrders, (order) => String(order.updatedAt || order.createdAt || '').slice(5, 10) || 'N/A', (order) => num(order.quantityKg || order.requestedWeight)).slice(-5)
  const capacityByMachine = activeMachines.map((machine) => {
    const active = filteredOrders.filter((order) => (order.mixingMachine || order.mixing?.machineCode) === machine.machineCode).reduce((sum, order) => sum + num(order.quantityKg || order.requestedWeight), 0)
    return [formatMixingMachineLabel(machine), Math.min(100, Math.round((active / Math.max(1, num(machine.capacityKg))) * 100))]
  })
  const outputByGroup = countBy(filteredOrders, getOrderGroup, (order) => num(order.quantityKg || order.requestedWeight)).slice(0, 4)
  const lossRows = [
    ['Định mức', lossLimitKg],
    ['Thực tế', materialLossKg],
  ]
  const renderExecutiveMetrics = (items) => items.map(([label, value, tone]) => <div className={`ceo-metric ${tone}`} key={label}><span>{label}</span><strong>{value}</strong></div>)
  const renderBars = (items, maxValue = Math.max(1, ...items.map(([, value]) => num(value)))) => items.map(([label, value]) => (
    <div className="ceo-bar-row" key={label}>
      <span>{label}</span>
      <i style={{ width: `${Math.max(4, Math.round((num(value) / maxValue) * 100))}%` }} />
      <strong>{typeof value === 'number' ? Math.round(value).toLocaleString('vi-VN') : value}</strong>
    </div>
  ))
  const renderDonutLegend = (items, total = items.reduce((sum, [, value]) => sum + num(value), 0)) => items.map(([label, value], index) => (
    <div className="ceo-legend-row" key={label}>
      <span className={`dot dot-${index + 1}`} />
      <strong>{label}</strong>
      <em>{total ? Math.round((num(value) / total) * 100) : 0}%</em>
    </div>
  ))
  return (
    <div className="page-content ceo-dashboard">
      <section className="ceo-filter-strip">
        <label>Từ ngày<input type="date" value={filters.fromDate} onChange={(event) => updateFilter('fromDate', event.target.value)} /></label>
        <label>Đến ngày<input type="date" value={filters.toDate} onChange={(event) => updateFilter('toDate', event.target.value)} /></label>
        <label>Nhóm sản phẩm<select value={filters.productGroup} onChange={(event) => updateFilter('productGroup', event.target.value)}><option value="all">Tất cả</option>{productGroups.map((group) => <option key={group} value={group}>{group}</option>)}</select></label>
        <label>Máy<select value={filters.machine} onChange={(event) => updateFilter('machine', event.target.value)}><option value="all">Tất cả</option>{activeMachines.map((machine) => <option key={machine.machineCode} value={machine.machineCode}>{mixingMachineOptionLabel(machine)}</option>)}</select></label>
        <label>Công đoạn<select value={filters.stage} onChange={(event) => updateFilter('stage', event.target.value)}><option value="all">Tất cả</option>{stageOptions.map(([stage, label]) => <option key={stage} value={stage}>{label}</option>)}</select></label>
      </section>
      <section className="ceo-kpi-strip">
        {renderExecutiveMetrics(productionKpis)}
      </section>
      <section className="ceo-chart-grid">
        <article className="ceo-chart-card"><h3>Sản lượng theo thời gian</h3><div className="ceo-bars">{renderBars(outputByTime)}</div></article>
        <article className="ceo-chart-card"><h3>Công suất theo máy</h3><div className="ceo-bars">{renderBars(capacityByMachine, 100)}</div></article>
        <article className="ceo-chart-card"><h3>Cơ cấu sản lượng</h3><div className="ceo-donut"><div className="donut-ring" /><div>{renderDonutLegend(outputByGroup)}</div></div></article>
        <article className="ceo-chart-card"><h3>Hao hụt thực tế vs định mức</h3><div className="ceo-bars loss">{renderBars(lossRows)}</div></article>
      </section>
      <section className="ceo-alert-strip">
        {alertKpis.map(([label, value, tone]) => <div className={`ceo-alert-item ${tone}`} key={label}><span>{label}</span><strong>{value}</strong></div>)}
      </section>
    </div>
  )
}

function normalizeProductionHistory(data = {}) {
  const orders = normalizeProductionOrders(data.orders || [], data.formulas || [])
  const machines = normalizeMixingMachines(data.mixingMachines || [])
  const productionLogs = data.productionLogs || data.logs || []
  const qc2Logs = data.qc2Logs || []
  const supplementalWeighing = data.supplementalWeighing || []
  const packingLogs = data.packingLogs || []
  const finishedGoods = normalizeFinishedGoodsData(data.finishedGoods || [])
  return orders.map((order) => {
    const orderKey = order.orderCode || order.id
    const qc1Rows = order.qc1Adjustments || order.qc1Logs || []
    const qc2Rows = getQc2Adjustments(order)
    const packingLog = getLatestPackingLog(order, packingLogs)
    const finishedRows = finishedGoods.filter((item) => item.orderId === order.id || item.orderCode === orderKey)
    const supplementRows = supplementalWeighing.filter((ticket) => ticket.orderId === order.id || ticket.orderCode === orderKey)
    const relatedLogs = productionLogs.filter((log) => String(log.entry || '').includes(orderKey) || log.orderId === order.id)
    const totalSupplementKg = qc2Rows.reduce((sum, ticket) => sum + getAdjustmentItems(ticket).reduce((lineSum, item) => lineSum + Math.max(0, num(item.adjustmentKg ?? item.requiredKg)), 0), 0)
    const mixingMachineLabel = getMixingMachineLabelByCode(order.mixingMachine || order.mixing?.machineCode || getOrderAssignedMachineCode(order), machines)
    const currentStage = ({
      qc1: 'QC sản xuất thử',
      weighing: 'Cân',
      'supplement-weighing': 'Cân bổ sung QC2',
      mixing: 'Phối trộn',
      'mixing-supplement': 'Phối trộn bổ sung QC2',
      'finished-qc': 'QC thành phẩm',
      packaging: 'Đóng gói',
      'finished-goods': 'Kho thành phẩm',
      completed: 'Hoàn thành',
    })[order.stage] || order.stage || '-'
    const timeline = [
      { time: order.createdAt, stage: 'Tạo lệnh SX', actor: order.createdBy || '-', content: `Tạo lệnh ${orderKey}`, status: order.qc1Status || order.status },
      ...qc1Rows.map((row, index) => ({ time: row.time || row.createdAt, stage: 'QC sản xuất thử', actor: row.createdBy || row.operator || 'QC', content: `QC sản xuất thử lần ${index + 1}: ${displayQcTrialText(row.result)}`, status: displayQcTrialText(row.result) })),
      ...(order.mixingStartAt || order.mixing?.startedAt ? [{ time: order.mixingStartAt || order.mixing?.startedAt, stage: 'Phối trộn', actor: order.mixing?.operator || 'Tổ phối trộn', content: `Bắt đầu phối trộn trên ${mixingMachineLabel}`, status: 'Bắt đầu' }] : []),
      ...(order.mixingCompletedAt || order.mixing?.completedAt ? [{ time: order.mixingCompletedAt || order.mixing?.completedAt, stage: 'Phối trộn', actor: order.mixing?.operator || 'Tổ phối trộn', content: 'Hoàn thành phối trộn', status: 'Hoàn thành' }] : []),
      ...qc2Rows.map((ticket) => ({ time: ticket.createdAt, stage: 'QC thành phẩm', actor: ticket.createdBy || 'QC', content: `${ticket.adjustmentId || ticket.id}: ${ticket.reason || 'Điều chỉnh QC2'}`, status: displayQc2Status(ticket.status) })),
      ...qc2Logs.filter((log) => log.orderId === order.id).map((log) => ({ time: log.time, stage: 'QC thành phẩm', actor: log.actor || 'QC', content: log.action || '-', status: log.result || '-' })),
      ...supplementRows.map((ticket) => ({ time: ticket.createdAt, stage: 'Cân bổ sung QC2', actor: ticket.createdBy || '-', content: `Phiếu cân ${ticket.id}`, status: displayQc2Status(ticket.status) })),
      ...(packingLog ? [{ time: packingLog.completedAt || packingLog.startedAt, stage: 'Đóng gói', actor: packingLog.packer || '-', content: `Đóng gói ${kg(packingLog.totalPackedWeight)}`, status: packingLog.status === 'completed' ? 'Hoàn thành' : packingLog.status }] : []),
      ...finishedRows.map((item) => ({ time: item.importDate, stage: 'Nhập kho TP', actor: item.receiver || '-', content: `Nhập kho ${item.finishedCode}`, status: item.status })),
      ...relatedLogs.map((log) => ({ time: log.time, stage: log.processName || log.stage || 'Nhật ký hệ thống', actor: log.employeeName || log.employee || log.user || log.actor || '-', employeeCode: log.employeeCode || '', employeeName: log.employeeName || log.employee || '', userAccount: log.userAccount || log.username || '', machine: log.machineCode || log.machineName || '', machineName: log.machineName || '', content: log.entry || log.action || log.actionDescription || '-', status: log.resultStatus || log.result || log.status || '-' })),
    ].filter((item) => item.time || item.content).sort((a, b) => String(a.time || '').localeCompare(String(b.time || ''), 'vi', { numeric: true }))
    return { order, machines, qc1Rows, qc2Rows, supplementRows, packingLog, finishedRows, timeline, totalSupplementKg, currentStage }
  })
}

function getHistoryFormulaRows(record) {
  const { order, qc2Rows } = record
  const original = order.originalFormulaSnapshot || order.productionFormulaSnapshot || []
  const qc1 = order.activeProductionFormula || order.qc1AdjustedFormula || order.productionFormulaSnapshot || original
  const codes = Array.from(new Set([...original, ...qc1, ...qc2Rows.flatMap(getAdjustmentItems)].map((item) => item.materialCode).filter(Boolean)))
  return codes.map((code) => {
    const originalItem = findMaterialByCode(original, code) || {}
    const qc1Item = findMaterialByCode(qc1, code) || originalItem
    const adds = qc2Rows.map((ticket) => getAdjustmentItems(ticket).filter((item) => item.materialCode === code).reduce((sum, item) => sum + num(item.adjustmentKg ?? item.requiredKg), 0))
    const totalAdd = adds.reduce((sum, value) => sum + value, 0)
    const qc1Kg = num(qc1Item.requiredKg ?? originalItem.requiredKg)
    return {
      materialCode: code,
      materialName: qc1Item.materialName || originalItem.materialName || code,
      materialGroup: qc1Item.materialGroup || originalItem.materialGroup || '-',
      originalKg: num(originalItem.requiredKg),
      qc1Kg,
      add1: adds[0] || 0,
      add2: adds[1] || 0,
      totalAdd,
      finalKg: Number((qc1Kg + totalAdd).toFixed(3)),
      note: qc2Rows.flatMap(getAdjustmentItems).filter((item) => item.materialCode === code).map((item) => item.note || item.reason).filter(Boolean).join('; ') || '-',
    }
  })
}

function getHistoryWeighingRows(record, group) {
  const mainRows = (record.order.activeProductionFormula || record.order.productionFormulaSnapshot || []).filter((item) => item.materialGroup === group).map((item) => ({ ...item, weighingType: 'Cân chính' }))
  const supplementRows = record.supplementRows.flatMap((ticket) => getTicketItems(ticket).filter((item) => item.materialGroup === group).map((item) => ({ ...item, weighingType: 'Cân bổ sung QC2' })))
  return [...mainRows, ...supplementRows]
}

function getHistoryMixingRows(record) {
  const order = record.order
  return (order.mixing || order.mixingStatus || order.mixingCompletedAt) ? [{
    no: 1,
    type: order.mixing?.supplement ? 'Bổ sung QC2' : 'Chính',
    machine: getMixingMachineLabelByCode(order.mixingMachine || order.mixing?.machineCode || getOrderAssignedMachineCode(order), record.machines || []),
    operator: order.mixing?.operator || 'Tổ phối trộn',
    startedAt: order.mixingStartAt || order.mixing?.startedAt || '-',
    completedAt: order.mixingCompletedAt || order.mixing?.completedAt || '-',
    duration: order.mixingStartAt && order.mixingCompletedAt ? `${Math.max(0, Math.round((new Date(order.mixingCompletedAt.replace(' ', 'T')) - new Date(order.mixingStartAt.replace(' ', 'T'))) / 60000))} phút` : '-',
    note: order.mixing?.note || '-',
  }] : []
}

function historyInfoRows(record) {
  const order = record.order
  const formulaRows = getHistoryFormulaRows(record)
  const packed = record.packingLog?.totalPackedWeight || order.packaging?.totalPackedWeight || 0
  const imported = record.finishedRows.reduce((sum, item) => sum + num(item.weight), 0)
  return [
    ['Mã lệnh SX', order.orderCode || order.id],
    ['Công thức gốc', order.formulaCode || order.originalFormulaId || '-'],
    ['Version', order.formulaVersion || order.originalFormulaVersion || '-'],
    ['Khách hàng', order.customer || '-'],
    ['Sản phẩm', order.productName || order.product || '-'],
    ['LOT', order.lot || '-'],
    ['Khối lượng yêu cầu', kg(order.requestedWeight ?? order.quantityKg)],
    ['Khối lượng sau QC1', kg(formulaRows.reduce((sum, item) => sum + item.qc1Kg, 0))],
    ['Khối lượng sau QC2', kg(formulaRows.reduce((sum, item) => sum + item.finalKg, 0))],
    ['Khối lượng sau phối trộn', kg(qc2FinalWeight(order))],
    ['Khối lượng đóng gói', kg(packed)],
    ['Khối lượng nhập kho', kg(imported)],
    ['Trạng thái cuối', order.orderStatus || order.status || '-'],
  ]
}

function inferLogStage(log = {}) {
  if (log.stage) return log.stage
  const text = String(log.entry || log.action || '')
  if (/cân hóa|Hóa/i.test(text)) return 'Cân hóa'
  if (/cân rắn|Rắn/i.test(text)) return 'Cân rắn'
  if (/phối trộn/i.test(text)) return 'Phối trộn'
  if (/QC|kiểm tra/i.test(text)) return 'QC'
  if (/đóng gói/i.test(text)) return 'Đóng gói'
  if (/nhập kho|thành phẩm/i.test(text)) return 'Kho thành phẩm'
  return '-'
}

function inferLogOrder(log = {}) {
  if (log.orderCode || log.orderId) return log.orderCode || log.orderId
  const text = String(log.entry || log.action || '')
  return text.match(/\b(?:LSX|ORD|HSS)-[A-Z0-9-]+/i)?.[0] || '-'
}

function normalizeSystemLogs(data = {}) {
  const byId = new Map()
  ;[...(data.logs || []), ...(data.productionLogs || [])].forEach((log) => {
    if (!log) return
    const id = log.id || `${log.time}-${log.entry || log.action}`
    byId.set(id, {
      id,
      time: log.time || log.createdAt || '-',
      username: log.username || log.user || log.actor || '-',
      employee: log.employeeName || log.employee || '-',
      employeeCode: log.employeeCode || '',
      role: log.role || '-',
      stage: log.processName || inferLogStage(log),
      orderCode: log.productionOrderCode || inferLogOrder(log),
      machine: log.machineName || log.machineCode || '',
      content: log.actionDescription || log.entry || log.action || '-',
      result: log.resultStatus || log.result || log.status || '-',
    })
  })
  return Array.from(byId.values()).sort((a, b) => String(b.time).localeCompare(String(a.time), 'vi', { numeric: true }))
}

function SystemLogsPage({ data }) {
  const [filters, setFilters] = useState({ username: '', employee: '', role: '', stage: '', orderCode: '' })
  const updateFilter = (field, value) => setFilters((current) => ({ ...current, [field]: value }))
  const logs = normalizeSystemLogs(data)
  const filtered = logs.filter((log) => (
    (!filters.username || String(log.username).toLowerCase().includes(filters.username.toLowerCase()))
    && (!filters.employee || String(log.employee).toLowerCase().includes(filters.employee.toLowerCase()))
    && (!filters.role || String(log.role).toLowerCase().includes(filters.role.toLowerCase()))
    && (!filters.stage || String(log.stage).toLowerCase().includes(filters.stage.toLowerCase()))
    && (!filters.orderCode || String(log.orderCode).toLowerCase().includes(filters.orderCode.toLowerCase()))
  ))
  return (
    <div className="page-content system-logs-page">
      <section className="panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Quản trị hệ thống</span><h2>Nhật ký hệ thống</h2></div>
        </div>
        <div className="production-form-grid">
          <label>Người dùng<input value={filters.username} onChange={(event) => updateFilter('username', event.target.value)} /></label>
          <label>Nhân viên<input value={filters.employee} onChange={(event) => updateFilter('employee', event.target.value)} /></label>
          <label>Vai trò<input value={filters.role} onChange={(event) => updateFilter('role', event.target.value)} /></label>
          <label>Công đoạn<input value={filters.stage} onChange={(event) => updateFilter('stage', event.target.value)} /></label>
          <label>Mã lệnh SX<input value={filters.orderCode} onChange={(event) => updateFilter('orderCode', event.target.value)} /></label>
        </div>
      </section>
      <section className="panel">
        <SimpleTable headers={['Người dùng', 'Nhân viên', 'Vai trò', 'Công đoạn', 'Lệnh sản xuất', 'Thời gian', 'Nội dung thao tác', 'Kết quả']} rows={filtered.map((log) => (
          <tr key={log.id}>
            <td>{log.username}</td>
            <td>{log.employee}</td>
            <td>{log.role}</td>
            <td>{log.stage}</td>
            <td>{log.orderCode}</td>
            <td>{log.time}</td>
            <td>{log.content}</td>
            <td>{log.result}</td>
          </tr>
        ))} empty="Chưa có nhật ký thao tác." />
      </section>
    </div>
  )
}

function LogsPage({ data }) {
  const [filters, setFilters] = useState({ fromDate: '', toDate: '', orderCode: '', product: '', lot: '', customer: '', status: '', stage: '', actor: '', employee: '', machine: '', userAccount: '' })
  const [selectedId, setSelectedId] = useState('')
  const [tab, setTab] = useState('info')
  const history = normalizeProductionHistory(data)
  const filtered = history.filter((record) => {
    const order = record.order
    if (filters.fromDate && String(order.createdAt || '').slice(0, 10) < filters.fromDate) return false
    if (filters.toDate && String(order.createdAt || '').slice(0, 10) > filters.toDate) return false
    if (filters.orderCode && !String(order.orderCode || order.id).toLowerCase().includes(filters.orderCode.toLowerCase())) return false
    if (filters.product && !String(order.productName || order.product || '').toLowerCase().includes(filters.product.toLowerCase())) return false
    if (filters.lot && !String(order.lot || '').toLowerCase().includes(filters.lot.toLowerCase())) return false
    if (filters.customer && !String(order.customer || '').toLowerCase().includes(filters.customer.toLowerCase())) return false
    if (filters.status && !String(order.status || order.orderStatus || '').toLowerCase().includes(filters.status.toLowerCase())) return false
    if (filters.stage && !String(record.currentStage || '').toLowerCase().includes(filters.stage.toLowerCase())) return false
    if (filters.actor && !record.timeline.some((item) => String(item.actor || '').toLowerCase().includes(filters.actor.toLowerCase()))) return false
    if (filters.employee && !record.timeline.some((item) => `${item.actor || ''} ${item.employeeName || ''} ${item.employeeCode || ''}`.toLowerCase().includes(filters.employee.toLowerCase()))) return false
    if (filters.machine && !record.timeline.some((item) => `${item.machine || ''} ${item.machineName || ''} ${item.content || ''}`.toLowerCase().includes(filters.machine.toLowerCase()))) return false
    if (filters.userAccount && !record.timeline.some((item) => `${item.userAccount || ''} ${item.actor || ''}`.toLowerCase().includes(filters.userAccount.toLowerCase()))) return false
    return true
  })
  const selected = history.find((record) => record.order.id === selectedId)
  const summaryRows = filtered.map((record, index) => {
    const order = record.order
    return {
      STT: index + 1,
      'Mã lệnh SX': order.orderCode || order.id,
      'Ngày tạo': order.createdAt || '-',
      'Khách hàng': order.customer || '-',
      'Sản phẩm': order.productName || order.product || '-',
      LOT: order.lot || '-',
      'Khối lượng yêu cầu': order.requestedWeight ?? order.quantityKg,
      'Trạng thái hiện tại': displayQcTrialText(order.orderStatus || order.status || '-'),
      'Công đoạn hiện tại': record.currentStage,
      'Số lần QC sản xuất thử điều chỉnh': record.qc1Rows.filter((row) => (row.changes || []).length).length,
      'Số lần QC2 điều chỉnh': record.qc2Rows.length,
      'Tổng kg bổ sung': record.totalSupplementKg,
      'Thời gian hoàn thành': order.completedAt || '-',
    }
  })
  const exportRecord = selected || filtered[0]
  const makeExportSheets = (record) => {
    if (!record) return { Tong_hop_lenh: summaryRows }
    return {
      Tong_hop_lenh: summaryRows,
      Cong_thuc_dieu_chinh: getHistoryFormulaRows(record),
      QC1: record.qc1Rows,
      Can_hoa: getHistoryWeighingRows(record, CHEMICAL),
      Can_ran: getHistoryWeighingRows(record, SOLID),
      Phoi_tron: getHistoryMixingRows(record),
      QC2: record.qc2Rows,
      Dong_goi: record.packingLog ? [record.packingLog] : [],
      Nhap_kho_TP: record.finishedRows,
      Timeline: record.timeline,
    }
  }
  const exportExcel = () => {
    const book = XLSX.utils.book_new()
    Object.entries(makeExportSheets(exportRecord)).forEach(([name, rows]) => XLSX.utils.book_append_sheet(book, XLSX.utils.json_to_sheet(rows), name.slice(0, 31)))
    XLSX.writeFile(book, 'nhat-ky-san-xuat.xlsx')
  }
  const exportPdf = () => {
    const doc = new jsPDF({ orientation: 'landscape' })
    doc.text('NHAT KY SAN XUAT', 14, 14)
    if (exportRecord) {
      autoTable(doc, { startY: 20, head: [['Thong tin', 'Gia tri']], body: historyInfoRows(exportRecord) })
      autoTable(doc, { head: [['Thoi gian', 'Cong doan', 'Nguoi thuc hien', 'Noi dung', 'Trang thai']], body: exportRecord.timeline.map((item) => [item.time || '-', item.stage || '-', item.actor || '-', item.content || '-', item.status || '-']) })
      autoTable(doc, { head: [['Ma VT', 'Nhom', 'Goc', 'Sau QC1', 'Bo sung QC2', 'Cuoi']], body: getHistoryFormulaRows(exportRecord).map((item) => [item.materialCode, item.materialGroup, item.originalKg, item.qc1Kg, item.totalAdd, item.finalKg]) })
    }
    doc.save('nhat-ky-san-xuat.pdf')
  }
  return (
    <div className="page-content production-history-page">
      <section className="panel">
        <div className="section-heading-row"><h2>Nhật ký sản xuất</h2><div className="action-row"><button className="secondary-button" onClick={exportExcel}>Tải Excel</button><button className="secondary-button" onClick={exportPdf}>Tải PDF</button></div></div>
        <div className="production-form-grid">
          <label>Từ ngày<input type="date" value={filters.fromDate} onChange={(event) => setFilters({ ...filters, fromDate: event.target.value })} /></label>
          <label>Đến ngày<input type="date" value={filters.toDate} onChange={(event) => setFilters({ ...filters, toDate: event.target.value })} /></label>
          <label>Mã lệnh SX<input value={filters.orderCode} onChange={(event) => setFilters({ ...filters, orderCode: event.target.value })} /></label>
          <label>Sản phẩm<input value={filters.product} onChange={(event) => setFilters({ ...filters, product: event.target.value })} /></label>
          <label>LOT<input value={filters.lot} onChange={(event) => setFilters({ ...filters, lot: event.target.value })} /></label>
          <label>Khách hàng<input value={filters.customer} onChange={(event) => setFilters({ ...filters, customer: event.target.value })} /></label>
          <label>Trạng thái<input value={filters.status} onChange={(event) => setFilters({ ...filters, status: event.target.value })} /></label>
          <label>Công đoạn<input value={filters.stage} onChange={(event) => setFilters({ ...filters, stage: event.target.value })} /></label>
          <label>Người thực hiện<input value={filters.actor} onChange={(event) => setFilters({ ...filters, actor: event.target.value })} /></label>
          <label>Nhân viên<input value={filters.employee} onChange={(event) => setFilters({ ...filters, employee: event.target.value })} /></label>
          <label>Máy<input value={filters.machine} onChange={(event) => setFilters({ ...filters, machine: event.target.value })} /></label>
          <label>Tài khoản thao tác<input value={filters.userAccount} onChange={(event) => setFilters({ ...filters, userAccount: event.target.value })} /></label>
        </div>
      </section>
      <section className="panel">
        <h3>Bảng tổng hợp lệnh SX</h3>
        <SimpleTable headers={['STT', 'Mã lệnh SX', 'Ngày tạo', 'Khách hàng', 'Sản phẩm', 'LOT', 'Khối lượng yêu cầu', 'Trạng thái hiện tại', 'Công đoạn hiện tại', 'Số lần QC sản xuất thử điều chỉnh', 'Số lần QC2 điều chỉnh', 'Tổng kg bổ sung', 'Thời gian hoàn thành', 'Hành động']} rows={filtered.map((record, index) => {
          const order = record.order
          return <tr key={order.id}><td>{index + 1}</td><td>{order.orderCode || order.id}</td><td>{order.createdAt || '-'}</td><td>{order.customer || '-'}</td><td>{order.productName || order.product || '-'}</td><td>{order.lot || '-'}</td><td>{kg(order.requestedWeight ?? order.quantityKg)}</td><td>{displayQcTrialText(order.orderStatus || order.status || '-')}</td><td>{record.currentStage}</td><td>{record.qc1Rows.filter((row) => (row.changes || []).length).length}</td><td>{record.qc2Rows.length}</td><td>{kg(record.totalSupplementKg)}</td><td>{order.completedAt || '-'}</td><td><button className="primary-button" onClick={() => { setSelectedId(order.id); setTab('info') }}>Xem chi tiết</button></td></tr>
        })} />
      </section>
      {selected && <ProductionHistoryModal record={selected} tab={tab} setTab={setTab} onClose={() => setSelectedId('')} />}
    </div>
  )
}

function ProductionHistoryModal({ record, tab, setTab, onClose }) {
  const tabs = [
    ['info', 'Thông tin lệnh'],
    ['formula', 'Công thức & điều chỉnh'],
    ['qc1', 'Nhật ký QC sản xuất thử'],
    ['chemical', 'Nhật ký cân hóa'],
    ['solid', 'Nhật ký cân rắn'],
    ['mixing', 'Nhật ký phối trộn'],
    ['qc2', 'Nhật ký QC thành phẩm'],
    ['packing', 'Nhật ký đóng gói'],
    ['warehouse', 'Nhật ký nhập kho TP'],
    ['timeline', 'Timeline sự kiện'],
  ]
  const order = record.order
  const formulaRows = getHistoryFormulaRows(record)
  const renderTab = () => {
    if (tab === 'info') return <div className="production-log-grid">{historyInfoRows(record).map(([label, value]) => <div key={label}><span>{label}</span><strong>{value}</strong></div>)}</div>
    if (tab === 'formula') return <SimpleTable headers={['Mã VT', 'Tên VT', 'Nhóm', 'Khối lượng gốc', 'Sau QC1', 'QC2 bổ sung lần 1', 'QC2 bổ sung lần 2', 'Tổng bổ sung QC2', 'Khối lượng cuối', 'Ghi chú']} rows={formulaRows.map((item) => <tr key={item.materialCode}><td>{item.materialCode}</td><td>{item.materialName}</td><td>{item.materialGroup}</td><td>{kg(item.originalKg)}</td><td>{kg(item.qc1Kg)}</td><td>{kg(item.add1)}</td><td>{kg(item.add2)}</td><td>{kg(item.totalAdd)}</td><td>{kg(item.finalKg)}</td><td>{item.note}</td></tr>)} />
    if (tab === 'qc1') return <SimpleTable headers={['Thời gian QC sản xuất thử', 'Người QC', 'Kết quả', 'Các điều chỉnh nếu có', 'Ghi chú']} rows={record.qc1Rows.map((row, index) => <tr key={row.id || index}><td>{row.time || row.createdAt || '-'}</td><td>{row.createdBy || row.operator || 'QC'}</td><td>{displayQcTrialText(row.result)}</td><td>{(row.changes || []).map((item) => `${item.materialCode}: ${kg(item.diff || item.requiredKg || 0)}`).join(', ') || '-'}</td><td>{row.note || '-'}</td></tr>)} />
    if (tab === 'chemical' || tab === 'solid') {
      const rows = getHistoryWeighingRows(record, tab === 'chemical' ? CHEMICAL : SOLID)
      return <SimpleTable headers={['Mã VT', 'Khối lượng yêu cầu', 'QR quét', 'Trạng thái QR', 'Khối lượng thực cân', 'Sai lệch', 'Kết quả', 'Người cân', 'Thời gian cân', 'Loại cân']} rows={rows.map((item, index) => {
        const actual = item.actualWeight === '' || item.actualWeight == null ? 0 : num(item.actualWeight)
        const diff = Number((actual - num(item.requiredKg)).toFixed(3))
        return <tr key={`${item.id}-${index}`}><td>{item.materialCode}</td><td>{kg(item.requiredKg)}</td><td>{item.qrScanned || '-'}</td><td>{item.qrStatus || '-'}</td><td>{actual ? kg(actual) : '-'}</td><td>{actual ? kg(diff) : '-'}</td><td>{item.weighStatus || '-'}</td><td>{item.weigher || '-'}</td><td>{item.confirmedAt || '-'}</td><td>{item.weighingType}</td></tr>
      })} />
    }
    if (tab === 'mixing') return <SimpleTable headers={['Lần phối trộn', 'Loại phối trộn', 'Máy phối trộn', 'Người thực hiện', 'Giờ bắt đầu', 'Giờ kết thúc', 'Thời gian phối trộn', 'Ghi chú']} rows={getHistoryMixingRows(record).map((item) => <tr key={item.no}><td>{item.no}</td><td>{item.type}</td><td>{item.machine}</td><td>{item.operator}</td><td>{item.startedAt}</td><td>{item.completedAt}</td><td>{item.duration}</td><td>{item.note}</td></tr>)} />
    if (tab === 'qc2') return <SimpleTable headers={['Lần QC2', 'Người QC', 'Màu sắc', 'pH', 'Độ nhớt', 'Tỷ trọng', 'Độ phủ', 'Độ mịn', 'Kết quả', 'Lý do điều chỉnh', 'Tổng kg bổ sung', 'Ghi chú']} rows={record.qc2Rows.map((ticket, index) => {
      const qc = ticket.qc2Record || order.qc2 || {}
      return <tr key={ticket.id || index}><td>{ticket.qc2No || index + 1}</td><td>{ticket.createdBy || 'QC'}</td><td>{qc.color || '-'}</td><td>{qc.ph || '-'}</td><td>{qc.viscosity || '-'}</td><td>{qc.density || '-'}</td><td>{qc.coverage || '-'}</td><td>{qc.fineness || '-'}</td><td>{qc.result || displayQc2Status(ticket.status)}</td><td>{ticket.reason || '-'}</td><td>{kg(ticket.totalSupplementKg || 0)}</td><td>{qc.note || '-'}</td></tr>
    })} />
    if (tab === 'packing') {
      const log = record.packingLog
      const rows = log?.packingDetails || []
      return <SimpleTable headers={['Quy cách 25kg / 10kg / 5kg', 'Số thùng', 'Khối lượng đóng gói', 'Sai lệch', 'Người đóng gói', 'Thời gian hoàn tất']} rows={rows.map((item, index) => <tr key={index}><td>{item.spec || `${item.sizeKg} kg`}</td><td>{item.boxes}</td><td>{kg(item.actualWeight)}</td><td>{kg((item.actualWeight || 0) - (item.convertedWeight || 0))}</td><td>{log?.packer || '-'}</td><td>{log?.completedAt || '-'}</td></tr>)} />
    }
    if (tab === 'warehouse') return <SimpleTable headers={['Mã TP', 'LOT', 'Quy cách', 'Số thùng', 'Khối lượng nhập kho', 'Vị trí kho', 'Người nhập', 'Ngày nhập kho']} rows={record.finishedRows.map((item) => <tr key={item.id}><td>{item.finishedCode}</td><td>{item.lot}</td><td>{item.spec}</td><td>{item.boxes}</td><td>{kg(item.weight)}</td><td>{item.location}</td><td>{item.receiver || '-'}</td><td>{item.importDate}</td></tr>)} />
    return <SimpleTable headers={['Thời gian', 'Công đoạn', 'Người thực hiện', 'Nội dung', 'Trạng thái']} rows={record.timeline.map((item, index) => <tr key={index}><td>{item.time || '-'}</td><td>{item.stage || '-'}</td><td>{item.actor || '-'}</td><td>{item.content || '-'}</td><td>{item.status || '-'}</td></tr>)} />
  }
  return (
    <div className="modal-backdrop" role="presentation">
      <div className="production-history-modal" role="dialog" aria-modal="true">
        <div className="modal-header">
          <div><span className="section-kicker">Truy xuất vòng đời lệnh SX</span><h2>{order.orderCode || order.id}</h2></div>
          <button type="button" className="icon-button" onClick={onClose} aria-label="Đóng">×</button>
        </div>
        <div className="log-tabs">{tabs.map(([id, label]) => <button key={id} className={tab === id ? 'active' : ''} onClick={() => setTab(id)}>{label}</button>)}</div>
        {renderTab()}
      </div>
    </div>
  )
}

function ReportsPage({ data, initialTab = 'production', lockedTab = false }) {
  const [tab, setTab] = useState(initialTab)
  const [traceLot, setTraceLot] = useState('')
  useEffect(() => {
    setTab(initialTab)
  }, [initialTab])
  const orders = normalizeProductionOrders(data.orders || [], data.formulas || [])
  const packingLogs = data.packingLogs || []
  const finishedGoods = normalizeFinishedGoodsData(data.finishedGoods || [])
  const machines = normalizeMixingMachines(data.mixingMachines)
  const activeMachines = getActiveMixingMachines(machines)
  const weighedContainers = normalizeWeighedContainers(data.weighedContainers || [])
  const qc2AdjustmentRows = orders.flatMap((order) => getQc2Adjustments(order).map((ticket) => ({ order, ticket })))
  const qc2AdjustmentItems = qc2AdjustmentRows.flatMap(({ order, ticket }) => getAdjustmentItems(ticket).map((item) => ({ ...item, order, ticket })))
  const topFormula = countBy(qc2AdjustmentRows, ({ order }) => order.formulaCode || order.originalFormulaId).slice(0, 5)
  const topMaterials = countBy(qc2AdjustmentItems, (item) => item.materialCode, (item) => Math.max(0, num(item.adjustmentKg ?? item.requiredKg))).slice(0, 5)
  const topQc = countBy(qc2AdjustmentRows, ({ ticket }) => ticket.createdBy).slice(0, 5)
  const qrFailLogs = (data.productionLogs || data.logs || []).filter((log) => String(log.entry || '').includes('QR') && String(log.entry || '').includes('FAIL'))
  const traceRecords = normalizeProductionHistory(data)
    .filter((record) => !traceLot || String(record.order.lot || '').toLowerCase().includes(traceLot.toLowerCase()))
  const pipelineStages = [
    ['qc1', 'QC sản xuất thử'],
    ['weighing', 'Tổ cân'],
    ['mixing', 'Phối trộn'],
    ['finished-qc', 'QC thành phẩm'],
    ['packaging', 'Đóng gói'],
    ['finished-goods', 'Kho thành phẩm'],
    ['completed', 'Hoàn thành'],
  ]
  const stageLabelMap = {
    qc1: 'QC sản xuất thử',
    weighing: 'Tổ cân nguyên liệu',
    mixing: 'Tổ phối trộn',
    'finished-qc': 'QC thành phẩm',
    packaging: 'Đóng gói',
    'finished-goods': 'Kho thành phẩm',
    completed: 'Hoàn thành',
  }
  const tabs = [
    ['production', 'Sản xuất'],
    ['qc', 'QC'],
    ['warehouse', 'Kho'],
    ['machines', 'Máy móc'],
    ['qr', 'Truy xuất QR'],
    ['trace', 'Truy xuất lô'],
  ]
  const reportKpis = {
    production: [
      ['Tổng lệnh', orders.length],
      ['Đang chạy', orders.filter((order) => !['completed', 'cancelled'].includes(order.stage)).length],
      ['Hoàn thành', orders.filter((order) => order.stage === 'completed').length],
      ['Tổng sản lượng', kg(orders.reduce((sum, order) => sum + num(order.quantityKg), 0))],
    ],
    qc: [
      ['Lệnh có QC2', orders.filter((order) => order.qc2 || getQc2Adjustments(order).length).length],
      ['QC thành phẩm OK', orders.filter((order) => order.qc2?.result === 'OK').length],
      ['Lệnh chỉnh màu', orders.filter((order) => getQc2Adjustments(order).length > 0).length],
      ['Lần chỉnh màu', qc2AdjustmentRows.length],
    ],
    warehouse: [
      ['Phiếu đóng gói', packingLogs.length],
      ['Kg đã đóng gói', kg(packingLogs.reduce((sum, log) => sum + num(log.totalPackedWeight), 0))],
      ['Mã TP nhập kho', finishedGoods.length],
      ['Kg TP nhập kho', kg(finishedGoods.reduce((sum, item) => sum + num(item.weight), 0))],
    ],
    machines: [
      ['Tổng máy', machines.length],
      ['Máy đang chạy', activeMachines.filter((machine) => orders.some((order) => (order.mixingMachine || order.mixing?.machineCode) === machine.machineCode && (order.mixing?.status === 'Active' || order.mixingStatus === 'Active'))).length],
      ['Lệnh chờ phối trộn', orders.filter((order) => getMixingDispatchState(order).canStart).length],
      ['Lệnh đang phối trộn', orders.filter((order) => ['mixing', 'mixing-supplement'].includes(order.stage)).length],
    ],
    qr: [
      ['QR hỗn hợp', weighedContainers.length],
      ['QR hóa', weighedContainers.filter((item) => item.materialGroup === CHEMICAL).length],
      ['QR rắn', weighedContainers.filter((item) => item.materialGroup === SOLID).length],
      ['Lỗi QR', qrFailLogs.length],
    ],
  }
  const renderKpis = (items) => <section className="report-kpi-grid compact">{items.map(([label, value]) => <article className="report-kpi-card" key={label}><span>{label}</span><strong>{value}</strong></article>)}</section>
  const renderTab = () => {
    if (tab === 'production') return (
      <>
        {renderKpis(reportKpis.production)}
        <section className="panel"><h2>Pipeline sản xuất</h2><div className="pipeline-flow report-pipeline">{pipelineStages.map(([stage, label]) => {
          const stageOrders = orders.filter((order) => order.stage === stage)
          const stageWeight = stageOrders.reduce((sum, order) => sum + num(order.quantityKg), 0)
          return (
            <div className="pipeline-step pipeline-card" key={stage}>
              <div className="pipeline-card-title">{stageLabelMap[stage] || label}</div>
              <div className="pipeline-card-count">{stageOrders.length}</div>
              <div className="pipeline-card-weight">{kg(stageWeight)}</div>
            </div>
          )
        })}</div></section>
        <section className="panel report-table-panel"><h2>Báo cáo lệnh sản xuất V3</h2><SimpleTable tableClassName="report-wide-table" headers={['Lệnh', 'Sản phẩm', 'LOT', 'Công thức', 'QC sản xuất thử', 'QC2', 'Đóng gói', 'Kho TP', 'Trạng thái']} rows={orders.map((order) => <tr key={order.id}><td>{order.id}</td><td>{order.product}</td><td>{order.lot}</td><td>{order.originalFormulaId}/{order.originalFormulaVersion}</td><td>{displayQcTrialText(order.qc1Result) || '-'}</td><td>{order.qc2?.result || '-'}</td><td>{order.packaging ? 'Đã đóng gói' : '-'}</td><td>{order.stage === 'completed' ? 'Đã nhập' : '-'}</td><td>{displayQcTrialText(order.status)}</td></tr>)} /></section>
      </>
    )
    if (tab === 'qc') return (
      <>
        {renderKpis(reportKpis.qc)}
        <section className="report-analysis-grid">
          <article className="panel report-table-panel"><h3>Top công thức phải điều chỉnh</h3><SimpleTable headers={['Công thức', 'Số lần']} rows={topFormula.map(([name, value]) => <tr key={name}><td>{name}</td><td>{value}</td></tr>)} /></article>
          <article className="panel report-table-panel"><h3>Top nguyên liệu bổ sung</h3><SimpleTable headers={['Nguyên liệu', 'Kg bổ sung']} rows={topMaterials.map(([name, value]) => <tr key={name}><td>{name}</td><td>{kg(value)}</td></tr>)} /></article>
          <article className="panel report-table-panel"><h3>Top QC điều chỉnh</h3><SimpleTable headers={['Người điều chỉnh', 'Số phiếu']} rows={topQc.map(([name, value]) => <tr key={name || '-'}><td>{name || '-'}</td><td>{value}</td></tr>)} /></article>
        </section>
        <section className="panel report-table-panel"><h2>Chi tiết QC</h2><SimpleTable tableClassName="report-wide-table" headers={['Lệnh', 'LOT', 'QC1', 'QC2', 'Số lần chỉnh màu', 'Kg bổ sung', 'Trạng thái']} rows={orders.map((order) => {
          const adjustments = getQc2Adjustments(order)
          const totalSupplement = adjustments.reduce((sum, ticket) => sum + getAdjustmentItems(ticket).reduce((lineSum, item) => lineSum + Math.max(0, num(item.adjustmentKg ?? item.requiredKg)), 0), 0)
          return <tr key={order.id}><td>{order.id}</td><td>{order.lot}</td><td>{displayQcTrialText(order.qc1Result) || '-'}</td><td>{order.qc2?.result || '-'}</td><td>{adjustments.length}</td><td>{kg(totalSupplement)}</td><td>{displayQcTrialText(order.status)}</td></tr>
        })} /></section>
      </>
    )
    if (tab === 'warehouse') return (
      <>
        {renderKpis(reportKpis.warehouse)}
        <section className="panel report-table-panel"><h2>Báo cáo đóng gói</h2><SimpleTable tableClassName="report-wide-table" headers={['Phiếu đóng gói', 'Lệnh SX', 'Sản phẩm', 'LOT', 'Khối lượng QC2', 'Đã đóng gói', 'Còn lại', 'Sai lệch', 'Người đóng gói', 'Trạng thái']} rows={packingLogs.map((log) => <tr key={log.packingId}><td>{log.packingId}</td><td>{log.orderCode || log.orderId}</td><td>{log.productName}</td><td>{log.lot}</td><td>{kg(log.qc2FinalWeight)}</td><td>{kg(log.totalPackedWeight)}</td><td>{kg(log.remainingWeight)}</td><td>{kg(log.differenceWeight)}</td><td>{log.packer || '-'}</td><td>{log.status === 'completed' ? 'Hoàn thành' : log.status}</td></tr>)} /></section>
        <section className="panel report-table-panel"><h2>Báo cáo kho thành phẩm</h2><SimpleTable tableClassName="report-wide-table" headers={['Mã TP', 'Lệnh SX', 'Sản phẩm', 'LOT', 'Quy cách', 'Số thùng', 'Khối lượng', 'Ngày nhập', 'Vị trí', 'Người nhập']} rows={finishedGoods.map((item) => <tr key={item.id}><td>{item.finishedCode}</td><td>{item.orderCode || item.orderId}</td><td>{item.productName}</td><td>{item.lot}</td><td>{item.spec}</td><td>{item.boxes}</td><td>{kg(item.weight)}</td><td>{item.importDate}</td><td>{item.location}</td><td>{item.receiver || '-'}</td></tr>)} /></section>
      </>
    )
    if (tab === 'machines') return (
      <>
        {renderKpis(reportKpis.machines)}
        <section className="panel report-table-panel"><h2>Máy phối trộn</h2><SimpleTable headers={['Mã máy', 'Tên máy', 'Công suất motor', 'Công suất kg/mẻ', 'Trạng thái', 'Lệnh đang chạy']} rows={machines.map((machine) => {
          const activeOrder = orders.find((order) => (order.mixingMachine || order.mixing?.machineCode) === machine.machineCode && (order.mixing?.status === 'Active' || order.mixingStatus === 'Active'))
          return <tr key={machine.machineCode}><td>{machine.machineCode}</td><td>{formatMixingMachineLabel(machine)}</td><td>{machine.motorPower || '-'}</td><td>{machine.capacityKg}</td><td>{machine.status}</td><td>{activeOrder?.id || '-'}</td></tr>
        })} /></section>
        <section className="panel report-table-panel"><h2>Lệnh phối trộn</h2><SimpleTable tableClassName="report-wide-table" headers={['Lệnh', 'Sản phẩm', 'LOT', 'Máy', 'Trạng thái phối trộn', 'Bắt đầu', 'Hoàn thành']} rows={orders.filter((order) => order.mixing || order.mixingStatus || ['mixing', 'mixing-supplement'].includes(order.stage)).map((order) => <tr key={order.id}><td>{order.id}</td><td>{order.product}</td><td>{order.lot}</td><td>{getOrderAssignedMachineLabel(order, machines)}</td><td>{order.mixingStatus || order.mixing?.status || '-'}</td><td>{order.mixingStartAt || order.mixing?.startedAt || '-'}</td><td>{order.mixingCompletedAt || order.mixing?.completedAt || '-'}</td></tr>)} /></section>
      </>
    )
    if (tab === 'trace') return (
      <>
        <section className="panel">
          <div className="section-heading-row"><h2>Truy xuất lô sản xuất</h2></div>
          <div className="production-form-grid">
            <label>Nhập LOT<input value={traceLot} onChange={(event) => setTraceLot(event.target.value)} placeholder="VD: LOT-HNS-G1-001" /></label>
          </div>
        </section>
        <section className="panel report-table-panel">
          <SimpleTable tableClassName="report-wide-table" headers={['LOT', 'Lệnh sản xuất', 'Sản phẩm', 'Người cân hóa', 'Người cân rắn', 'Người phối trộn', 'Người QC', 'Người đóng gói', 'Người nhập kho TP', 'Thời gian công đoạn', 'Máy phối trộn']} rows={traceRecords.map((record) => {
            const byStage = (name) => record.timeline.filter((item) => String(item.stage || '').includes(name))
            const actorText = (name) => byStage(name).map((item) => item.actor).filter(Boolean).join(', ') || '-'
            const timeText = record.timeline.map((item) => `${item.stage}: ${item.time || '-'}`).join(' | ')
            const machineText = getHistoryMixingRows(record).map((item) => item.machine).filter(Boolean).join(', ') || '-'
            return (
              <tr key={record.order.id}>
                <td>{record.order.lot || '-'}</td>
                <td>{record.order.orderCode || record.order.id}</td>
                <td>{record.order.productName || record.order.product || '-'}</td>
                <td>{actorText('Cân hóa')}</td>
                <td>{actorText('Cân rắn')}</td>
                <td>{actorText('Phối trộn')}</td>
                <td>{actorText('QC')}</td>
                <td>{actorText('Đóng gói')}</td>
                <td>{actorText('kho')}</td>
                <td>{timeText || '-'}</td>
                <td>{machineText}</td>
              </tr>
            )
          })} empty="Chưa có dữ liệu truy xuất cho LOT đã chọn." />
        </section>
      </>
    )
    return (
      <>
        {renderKpis(reportKpis.qr)}
        <section className="panel report-table-panel"><h2>QR hỗn hợp đã tạo</h2><SimpleTable tableClassName="report-wide-table" headers={['QR', 'Lệnh', 'Nhóm', 'Khối lượng', 'Trạng thái', 'Thời gian']} rows={weighedContainers.map((item) => <tr key={item.id || item.qrCode}><td>{item.qrCode || item.id}</td><td>{item.orderCode || item.orderId || '-'}</td><td>{item.materialGroup || '-'}</td><td>{kg(item.totalWeight || item.weight)}</td><td>{item.status || '-'}</td><td>{item.createdAt || item.confirmedAt || '-'}</td></tr>)} /></section>
        <section className="panel report-table-panel"><h2>Lỗi xác nhận QR</h2><SimpleTable headers={['Thời gian', 'Nội dung']} rows={qrFailLogs.map((log) => <tr key={log.id}><td>{log.time || '-'}</td><td>{log.entry}</td></tr>)} /></section>
      </>
    )
  }
  return (
    <div className="page-content reports-page">
      <section className="panel report-shell">
        <div className="section-heading-row">
          <div><span className="section-kicker">Phân tích chi tiết</span><h2>Báo cáo sản xuất</h2></div>
        </div>
        {!lockedTab && <div className="log-tabs report-tabs">{tabs.map(([id, label]) => <button key={id} className={tab === id ? 'active' : ''} onClick={() => setTab(id)}>{label}</button>)}</div>}
      </section>
      {renderTab()}
    </div>
  )
}

function StaffPerformanceReportPage({ data }) {
  const employees = data.employeeCatalog || []
  const assignments = data.productionAssignments || []
  const logs = normalizeSystemLogs(data)
  const rows = employees.map((employee) => {
    const employeeAssignments = assignments.filter((item) => assignmentEmployeeCodes(item).includes(employee.code) || assignmentEmployeeNames(item).includes(employee.name))
    const employeeLogs = logs.filter((log) => (Array.isArray(log.employeeCodes) && log.employeeCodes.includes(employee.code)) || log.employeeCode === employee.code || (Array.isArray(log.employeeNames) && log.employeeNames.includes(employee.name)) || log.employee === employee.name)
    return {
      employee,
      assigned: employeeAssignments.length,
      completed: employeeAssignments.filter((item) => item.status === 'Hoàn thành').length,
      actions: employeeLogs.length,
      lastAction: employeeLogs[0]?.time || employeeAssignments[0]?.assignedAt || '-',
    }
  })
  return (
    <div className="page-content reports-page">
      <section className="panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Báo cáo</span><h2>Hiệu suất nhân sự</h2></div>
        </div>
        <SimpleTable
          tableClassName="report-wide-table"
          headers={['Mã NV', 'Nhân viên', 'Tổ sản xuất', 'Vai trò vận hành', 'Số phân công', 'Hoàn thành', 'Số thao tác', 'Thao tác gần nhất']}
          rows={rows.map(({ employee, assigned, completed, actions, lastAction }) => (
            <tr key={employee.code}>
              <td>{employee.code}</td>
              <td>{employee.name}</td>
              <td>{employee.productionTeam}</td>
              <td>{employee.operationRole}</td>
              <td>{assigned}</td>
              <td>{completed}</td>
              <td>{actions}</td>
              <td>{lastAction}</td>
            </tr>
          ))}
        />
      </section>
    </div>
  )
}

function MachinePerformanceReportPage({ data }) {
  const machines = normalizeMixingMachines(data.mixingMachines || [])
  const orders = normalizeProductionOrders(data.orders || [], data.formulas || [])
  const rows = machines.map((machine) => {
    const machineOrders = orders.filter((order) => (order.mixingMachine || order.mixing?.machineCode) === machine.machineCode)
    return {
      machine,
      orderCount: machineOrders.length,
      activeCount: machineOrders.filter((order) => order.mixing?.status === 'Active' || order.mixingStatus === 'Active').length,
      completedCount: machineOrders.filter((order) => order.mixingCompletedAt || order.mixing?.completedAt).length,
      lastOrder: machineOrders[0]?.orderCode || machineOrders[0]?.id || '-',
    }
  })
  return (
    <div className="page-content reports-page">
      <section className="panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Báo cáo</span><h2>Hiệu suất máy</h2></div>
        </div>
        <SimpleTable
          tableClassName="report-wide-table"
          headers={['Mã máy', 'Tên máy', 'Công suất kg/mẻ', 'Trạng thái', 'Số lệnh', 'Đang chạy', 'Đã hoàn thành phối trộn', 'Lệnh gần nhất']}
          rows={rows.map(({ machine, orderCount, activeCount, completedCount, lastOrder }) => (
            <tr key={machine.machineCode}>
              <td>{machine.machineCode}</td>
              <td>{formatMixingMachineLabel(machine)}</td>
              <td>{machine.capacityKg || '-'}</td>
              <td>{machine.status || '-'}</td>
              <td>{orderCount}</td>
              <td>{activeCount}</td>
              <td>{completedCount}</td>
              <td>{lastOrder}</td>
            </tr>
          ))}
        />
      </section>
    </div>
  )
}

function MixingMachineCatalogPage({ data, setData, user, permissions = [] }) {
  const emptyMachineDraft = { machineCode: '', machineName: '', motorPower: '', capacityKg: '', department: 'Tổ phối trộn', productionTeam: 'Tổ phối trộn', status: 'READY', note: '' }
  const [machineDraft, setMachineDraft] = useState(emptyMachineDraft)
  const [editingMachineCode, setEditingMachineCode] = useState('')
  const [notice, setNotice] = useState('')
  const machines = normalizeMixingMachines(data.mixingMachines)
  const pendingMachineRequests = (data.orders || []).filter((order) => order.machineChangeRequest?.status === 'PENDING')
  const canCreateMachine = hasPermission(permissions, 'master.machine.create')
  const canEditMachine = hasPermission(permissions, 'master.machine.edit')
  const canDeleteMachine = hasPermission(permissions, 'master.machine.delete')
  const canApproveMachineChange = user?.role === 'Admin' || user?.role === 'Sản xuất'
  const updateMachineDraft = (field, value) => setMachineDraft((current) => ({ ...current, [field]: value }))
  const resetMachineDraft = () => {
    setMachineDraft(emptyMachineDraft)
    setEditingMachineCode('')
  }
  const editMachine = (machine) => {
    if (!canEditMachine) return
    setMachineDraft({
      machineCode: machine.machineCode,
      machineName: machine.machineName,
      motorPower: machine.motorPower || '',
      capacityKg: machine.capacityKg,
      department: machine.department || machine.productionTeam || 'Tổ phối trộn',
      productionTeam: machine.productionTeam || machine.department || 'Tổ phối trộn',
      status: normalizeMixingMachineStatus(machine.status),
      note: machine.note || '',
    })
    setEditingMachineCode(machine.machineCode)
  }
  const saveMachine = () => {
    if (!(editingMachineCode ? canEditMachine : canCreateMachine)) {
      setNotice('Bạn chưa có quyền tạo hoặc sửa danh mục máy phối trộn.')
      return
    }
    const machine = { ...normalizeMixingMachine(machineDraft), catalogOverride: true }
    if (!machine.machineCode || !machine.machineName || !machine.capacityKg) {
      setNotice('Vui lòng nhập Mã máy, Tên máy và Công suất kg.')
      return
    }
    const duplicate = machines.some((item) => item.machineCode === machine.machineCode && item.machineCode !== editingMachineCode)
    if (duplicate) {
      setNotice(`Mã máy ${machine.machineCode} đã tồn tại.`)
      return
    }
    setData((current) => {
      const currentMachines = normalizeMixingMachines(current.mixingMachines)
      const exists = currentMachines.some((item) => item.machineCode === editingMachineCode)
      const nextMachines = exists
        ? currentMachines.map((item) => item.machineCode === editingMachineCode ? machine : item)
        : [...currentMachines, machine]
      return addLogToData({ ...current, mixingMachines: normalizeMixingMachines(nextMachines) }, `${editingMachineCode ? 'Cập nhật' : 'Thêm'} máy phối trộn ${formatMixingMachineLabel(machine)}.`)
    })
    setNotice(`${editingMachineCode ? 'Đã cập nhật' : 'Đã thêm'} máy phối trộn ${formatMixingMachineLabel(machine)}.`)
    resetMachineDraft()
  }
  const deactivateMachine = (machineCode) => {
    if (!canDeleteMachine) {
      setNotice('Bạn chưa có quyền ngừng sử dụng máy phối trộn.')
      return
    }
    const machineLabel = getMixingMachineLabelByCode(machineCode, machines)
    setData((current) => addLogToData({
      ...current,
      mixingMachines: normalizeMixingMachines(current.mixingMachines).map((machine) => machine.machineCode === machineCode ? { ...machine, status: 'INACTIVE' } : machine),
    }, `Ngừng sử dụng máy phối trộn ${machineLabel}.`))
    setNotice(`Đã chuyển ${machineLabel} sang INACTIVE.`)
    if (editingMachineCode === machineCode) resetMachineDraft()
  }
  const resolveMachineRequest = (order, approved) => {
    const request = order.machineChangeRequest
    if (!request) return
    if (!canApproveMachineChange) {
      setNotice('Chỉ Admin hoặc Sản xuất được duyệt đổi máy.')
      return
    }
    const nextMachine = machines.find((machine) => machine.machineCode === request.requestedMachine && machine.status === 'READY')
    if (approved && !nextMachine) {
      setNotice(`Máy ${getMixingMachineLabelByCode(request.requestedMachine, machines)} không tồn tại hoặc không ở trạng thái READY.`)
      return
    }
    const rejectReason = approved ? '' : window.prompt('Lý do từ chối:', '')?.trim()
    if (!approved && !rejectReason) return
    const resolvedAt = nowText()
    const approvedBy = user?.fullName || user?.username || user?.role || 'Admin'
    setData((current) => addLogToData({
      ...current,
      orders: (current.orders || []).map((item) => item.id === order.id ? {
        ...item,
        mixerMachine: approved ? nextMachine.machineCode : item.mixerMachine,
        assignedMixingMachine: approved ? nextMachine.machineCode : item.assignedMixingMachine,
        assignedMachineCode: approved ? nextMachine.machineCode : item.assignedMachineCode,
        assignedMachineName: approved ? nextMachine.machineName : item.assignedMachineName,
        assignedMachineCapacityKg: approved ? nextMachine.capacityKg : item.assignedMachineCapacityKg,
        assignedMachineMotorPower: approved ? nextMachine.motorPower : item.assignedMachineMotorPower,
        assignedMachineDepartment: approved ? (nextMachine.department || nextMachine.productionTeam) : item.assignedMachineDepartment,
        machineChangeHistory: [
          ...(item.machineChangeHistory || []),
          {
            id: uid('MCH'),
            orderId: item.id,
            orderCode: item.orderCode || item.id,
            lot: item.lot,
            oldMachine: request.currentMachine || '',
            newMachine: request.requestedMachine,
            assignedMachine: approved ? nextMachine.machineCode : getOrderAssignedMachineCode(item),
            performedMachine: item.mixingMachine || item.mixing?.machineCode || '',
            reason: request.reason,
            requestedBy: request.requestedBy,
            requestedAt: request.requestedAt,
            approved,
            approvedBy,
            approvedAt: resolvedAt,
            changedBy: approvedBy,
            changedAt: resolvedAt,
            rejectedReason: rejectReason,
          },
        ],
        machineAssignmentHistory: approved ? [
          ...(item.machineAssignmentHistory || []),
          {
            id: uid('MCH'),
            orderId: item.id,
            orderCode: item.orderCode || item.id,
            lot: item.lot,
            assignedMachine: nextMachine.machineCode,
            performedMachine: item.mixingMachine || item.mixing?.machineCode || '',
            changedBy: approvedBy,
            changedAt: resolvedAt,
            reason: request.reason,
            requestedBy: request.requestedBy,
          },
        ] : (item.machineAssignmentHistory || []),
        machineChangeRequest: {
          ...request,
          status: approved ? 'APPROVED' : 'REJECTED',
          resolvedAt,
          resolvedBy: approvedBy,
          rejectedReason: rejectReason,
        },
        updatedAt: resolvedAt,
      } : item),
    }, `${approved ? 'Duyệt' : 'Từ chối'} đề nghị đổi máy lệnh ${order.orderCode || order.id}: ${getMixingMachineLabelByCode(request.currentMachine, machines)} -> ${getMixingMachineLabelByCode(request.requestedMachine, machines)}. Lý do: ${request.reason}. Người đề nghị: ${request.requestedBy}. Người duyệt: ${approvedBy}.`))
    setNotice(`${approved ? 'Đã duyệt' : 'Đã từ chối'} đề nghị đổi máy ${order.orderCode || order.id}.`)
  }
  return (
    <div className="page-content admin-page">
      <section className="panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Dữ liệu gốc</span><h2>Danh mục máy phối trộn</h2></div>
          <div className="action-row">
            <button type="button" className="secondary-button" disabled={!canCreateMachine} onClick={resetMachineDraft}>Nhập mới</button>
            <button type="button" className="primary-button" disabled={editingMachineCode ? !canEditMachine : !canCreateMachine} onClick={saveMachine}>{editingMachineCode ? 'Lưu máy' : 'Thêm máy'}</button>
          </div>
        </div>
        {notice && <p className="empty-alert">{notice}</p>}
        <div className="production-form-grid order-create-form">
          <label>Mã máy<input value={machineDraft.machineCode} onChange={(event) => updateMachineDraft('machineCode', event.target.value.toUpperCase())} placeholder="M01" disabled={Boolean(editingMachineCode) || !(editingMachineCode ? canEditMachine : canCreateMachine)} /></label>
          <label>Tên máy<input value={machineDraft.machineName} readOnly={!(editingMachineCode ? canEditMachine : canCreateMachine)} onChange={(event) => updateMachineDraft('machineName', event.target.value)} placeholder="Máy phối trộn 500kg" /></label>
          <label>Công suất motor<input value={machineDraft.motorPower} readOnly={!(editingMachineCode ? canEditMachine : canCreateMachine)} onChange={(event) => updateMachineDraft('motorPower', event.target.value)} placeholder="7.5HP" /></label>
          <label>Công suất kg<input type="number" value={machineDraft.capacityKg} readOnly={!(editingMachineCode ? canEditMachine : canCreateMachine)} onChange={(event) => updateMachineDraft('capacityKg', event.target.value)} placeholder="500" /></label>
          <label>Tổ sản xuất/phân xưởng<input value={machineDraft.department} readOnly={!(editingMachineCode ? canEditMachine : canCreateMachine)} onChange={(event) => { updateMachineDraft('department', event.target.value); updateMachineDraft('productionTeam', event.target.value) }} placeholder="Tổ phối trộn" /></label>
          <label>Trạng thái<select value={machineDraft.status} disabled={!(editingMachineCode ? canEditMachine : canCreateMachine)} onChange={(event) => updateMachineDraft('status', event.target.value)}><option value="READY">READY</option><option value="INACTIVE">INACTIVE</option><option value="MAINTENANCE">MAINTENANCE</option></select></label>
          <label className="wide-field">Ghi chú<textarea value={machineDraft.note} readOnly={!(editingMachineCode ? canEditMachine : canCreateMachine)} onChange={(event) => updateMachineDraft('note', event.target.value)} /></label>
        </div>
        <div className="table-wrapper">
          <table className="admin-wide-table">
            <thead>
              <tr>
                <th>Mã máy</th>
                <th>Tên máy</th>
                <th>Công suất motor</th>
                <th>Công suất kg</th>
                <th>Tổ sản xuất/phân xưởng</th>
                <th>Trạng thái</th>
                <th>Ghi chú</th>
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {machines.map((machine) => (
                <tr key={machine.machineCode}>
                  <td><strong>{machine.machineCode}</strong></td>
                  <td>{formatMixingMachineLabel(machine)}</td>
                  <td>{machine.motorPower || '-'}</td>
                  <td>{machine.capacityKg}</td>
                  <td>{machine.department || machine.productionTeam || '-'}</td>
                  <td><span className={`dispatch-badge ${machine.status === 'READY' ? 'ready' : machine.status === 'MAINTENANCE' ? 'waiting' : 'fail'}`}>{machine.status}</span></td>
                  <td>{machine.note || '-'}</td>
                  <td>
                    <div className="user-action-group">
                      <button type="button" className="secondary-button" disabled={!canEditMachine} onClick={() => editMachine(machine)}>Sửa</button>
                      <button type="button" className="danger-button" disabled={!canDeleteMachine || machine.status === 'INACTIVE'} onClick={() => deactivateMachine(machine.machineCode)}>Ngừng sử dụng</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
      <section className="panel">
        <h3>Đề nghị đổi máy</h3>
        {!canApproveMachineChange && <div className="process-alert">Chỉ Admin hoặc Sản xuất được duyệt đổi máy.</div>}
        <SimpleTable headers={['Lệnh', 'Máy hiện tại', 'Máy đề nghị', 'Lý do', 'Thời gian', 'Hành động']} rows={pendingMachineRequests.map((order) => (
          <tr key={order.id}>
            <td>{order.orderCode || order.id}</td>
            <td>{getMixingMachineLabelByCode(order.machineChangeRequest.currentMachine, machines)}</td>
            <td>{getMixingMachineLabelByCode(order.machineChangeRequest.requestedMachine, machines)}</td>
            <td>{order.machineChangeRequest.reason}</td>
            <td>{order.machineChangeRequest.requestedAt || '-'}</td>
            <td>
              <div className="user-action-group">
                <button type="button" className="primary-button" disabled={!canApproveMachineChange} onClick={() => resolveMachineRequest(order, true)}>Duyệt</button>
                <button type="button" className="secondary-button" disabled={!canApproveMachineChange} onClick={() => resolveMachineRequest(order, false)}>Từ chối</button>
              </div>
            </td>
          </tr>
        ))} empty="Không có đề nghị đổi máy đang chờ." />
      </section>
    </div>
  )
}

function ProductionAssignmentPage({ data, setData, user, permissions = [] }) {
  const canCreate = hasPermission(permissions, 'production.assignment.create')
  const canEdit = hasPermission(permissions, 'production.assignment.edit')
  const canDelete = hasPermission(permissions, 'production.assignment.delete')
  const visibleStages = roleAssignmentStageMap[user?.role] || productionAssignmentStages
  const activeEmployees = (data.employeeCatalog || []).filter((employee) => employee.status !== 'Ngừng sử dụng')
  const shifts = data.shiftCatalog || []
  const teams = data.teamCatalog || []
  const machines = getActiveMixingMachines(normalizeMixingMachines(data.mixingMachines || []))
  const employeeByCode = (code) => activeEmployees.find((employee) => employee.code === code)
  const teamByCode = (code) => teams.find((team) => team.code === code)
  const shiftByCode = (code) => shifts.find((shift) => shift.code === code)
  const machineByCode = (code) => machines.find((machine) => machine.machineCode === code)
  const teamStageOptions = {
    TP1: ['Phối trộn'],
    TP2: ['Phối trộn'],
    TH: ['Cân hóa'],
    TC: ['Cân rắn'],
    QC: ['QC sản xuất thử', 'QC thành phẩm'],
    DG: ['Đóng gói'],
    KTP: ['Kho thành phẩm'],
  }
  const teamEmployeeAliases = {
    TP1: ['TP1', 'Tổ trộn 1', 'Tổ phối trộn 1'],
    TP2: ['TP2', 'Tổ trộn 2', 'Tổ phối trộn 2'],
    TH: ['TH', 'Tổ Hóa'],
    TC: ['TC', 'Tổ Cát'],
    QC: ['QC'],
    DG: ['DG', 'Đóng gói'],
    KTP: ['KTP', 'Kho thành phẩm'],
  }
  const getAllowedStagesForTeam = (teamCode) => {
    const stages = teamStageOptions[teamCode] || visibleStages
    return stages.filter((stage) => visibleStages.includes(stage))
  }
  const availableTeams = teams.filter((team) => getAllowedStagesForTeam(team.code).length)
  const defaultTeamCode = availableTeams[0]?.code || teams[0]?.code || ''
  const defaultStage = getAllowedStagesForTeam(defaultTeamCode)[0] || visibleStages[0] || productionAssignmentStages[0]
  const employeesByTeam = (teamNameOrCode) => {
    const aliases = new Set([teamNameOrCode, teamByCode(teamNameOrCode)?.name, ...(teamEmployeeAliases[teamNameOrCode] || [])].filter(Boolean))
    return activeEmployees.filter((employee) => aliases.has(employee.productionTeam))
  }
  const [notice, setNotice] = useState('')
  const [form, setForm] = useState({
    date: todayText(),
    shiftCode: shifts[0]?.code || '',
    teamCode: defaultTeamCode,
    stage: defaultStage,
    machineCodes: [],
    employeeCodes: [],
  })
  const updateForm = (field, value) => {
    setForm((current) => {
      const next = { ...current, [field]: value }
      const nextAllowedStages = getAllowedStagesForTeam(next.teamCode)
      if (field === 'teamCode') {
        next.stage = nextAllowedStages.includes(current.stage) ? current.stage : nextAllowedStages[0] || ''
        next.employeeCodes = []
      }
      if (field === 'stage' && !nextAllowedStages.includes(value)) next.stage = nextAllowedStages[0] || ''
      if (next.stage !== 'Phối trộn') next.machineCodes = []
      return next
    })
  }
  const allowedStages = getAllowedStagesForTeam(form.teamCode)
  const selectableEmployees = employeesByTeam(form.teamCode)
  useEffect(() => {
    setForm((current) => {
      const normalizedTeamCode = getAllowedStagesForTeam(current.teamCode).length ? current.teamCode : defaultTeamCode
      const nextAllowedStages = getAllowedStagesForTeam(normalizedTeamCode)
      const nextEmployees = employeesByTeam(normalizedTeamCode)
      const nextStage = nextAllowedStages.includes(current.stage) ? current.stage : nextAllowedStages[0] || ''
      const currentEmployeeCodes = Array.isArray(current.employeeCodes) ? current.employeeCodes : [current.employeeCode].filter(Boolean)
      const nextEmployeeCodes = currentEmployeeCodes.filter((code) => nextEmployees.some((employee) => employee.code === code))
      const currentMachineCodes = Array.isArray(current.machineCodes) ? current.machineCodes : [current.machineCode].filter(Boolean)
      const nextMachineCodes = nextStage === 'Phối trộn' ? currentMachineCodes.filter((code) => machines.some((machine) => machine.machineCode === code)) : []
      if (
        normalizedTeamCode === current.teamCode
        && nextStage === current.stage
        && nextEmployeeCodes.join('|') === currentEmployeeCodes.join('|')
        && nextMachineCodes.join('|') === currentMachineCodes.join('|')
      ) return current
      return { ...current, teamCode: normalizedTeamCode, stage: nextStage, employeeCodes: nextEmployeeCodes, machineCodes: nextMachineCodes }
    })
  }, [data.employeeCatalog, data.teamCatalog, defaultTeamCode, visibleStages])
  const buildAssignment = ({ workDate, shift, team, processName, employees = [], assignedMachines = [], status = 'Chưa bắt đầu', assignmentNote = '' }) => {
    const assignedAt = nowText()
    const employeeCodes = employees.map((employee) => employee.code)
    const employeeNames = employees.map((employee) => employee.name)
    const employeeQrs = employees.map((employee) => employee.qrEmployee || employee.qr || '').filter(Boolean)
    const machineCodes = assignedMachines.map((machine) => machine.machineCode)
    const machineNames = assignedMachines.map((machine) => formatMixingMachineLabel(machine))
    const firstEmployee = employees[0] || {}
    return {
      id: uid('ASSIGN'),
      assignmentId: uid('ASSIGN'),
      date: workDate,
      workDate,
      shiftCode: shift.code,
      shiftName: shift.name,
      teamCode: team?.code || firstEmployee.productionTeam || '',
      teamName: team?.name || firstEmployee.productionTeam || '',
      stage: processName,
      processCode: processCodeByName[processName] || processName,
      processName,
      machineCodes,
      machineNames,
      machineCode: machineCodes.join(', '),
      machineName: machineNames.join(', '),
      employeeCodes,
      employeeNames,
      employeeQrs,
      employeeCode: employeeCodes.join(', '),
      employeeName: employeeNames.join(', '),
      employeeQr: employeeQrs.join(', '),
      productionTeam: team?.code || firstEmployee.productionTeam || '',
      productionTeamName: team?.name || firstEmployee.productionTeam || '',
      assignedBy: user?.fullName || user?.username || 'Hệ thống',
      assignedAt,
      note: assignmentNote,
      assignmentNote,
      status,
    }
  }
  const visibleAssignments = getStageAssignmentsForRole(data.productionAssignments || [], user?.role)
    .slice()
    .sort((a, b) => `${b.date} ${b.assignedAt}`.localeCompare(`${a.date} ${a.assignedAt}`, 'vi', { numeric: true }))
  const saveAssignment = () => {
    if (!canCreate) return
    const selectedEmployees = (form.employeeCodes || []).map(employeeByCode).filter(Boolean)
    const shift = shiftByCode(form.shiftCode)
    const team = teamByCode(form.teamCode)
    const selectedMachines = form.stage === 'Phối trộn' ? (form.machineCodes || []).map(machineByCode).filter(Boolean) : []
    const teamEmployees = employeesByTeam(form.teamCode)
    const teamStages = getAllowedStagesForTeam(form.teamCode)
    if (!team || selectedEmployees.length === 0 || !shift || !form.date || !form.stage) {
      setNotice('Vui lòng chọn đủ ngày, ca, tổ sản xuất, công đoạn và nhân viên.')
      return
    }
    if (!teamStages.includes(form.stage)) {
      setNotice('Công đoạn không phù hợp với tổ sản xuất đã chọn.')
      return
    }
    if (selectedEmployees.some((employee) => !teamEmployees.some((item) => item.code === employee.code))) {
      setNotice('Nhân viên không thuộc tổ sản xuất đã chọn.')
      return
    }
    if (form.stage === 'Phối trộn' && selectedMachines.length === 0) {
      setNotice('Vui lòng chọn ít nhất một máy phối trộn cho công đoạn Phối trộn.')
      return
    }
    const assignment = buildAssignment({ workDate: form.date, shift, team, processName: form.stage, employees: selectedEmployees, assignedMachines: selectedMachines })
    setData((current) => addLogToData({
      ...current,
      productionAssignments: [assignment, ...(current.productionAssignments || [])],
    }, `Phân công ${formatAssignmentEmployees(assignment)} vào ${assignment.stage} ca ${assignment.shiftCode} ngày ${assignment.date}.`))
    setNotice('Đã lưu phân công.')
  }
  const updateAssignmentStatus = (assignmentId, status) => {
    if (!canEdit) return
    setData((current) => ({
      ...current,
      productionAssignments: (current.productionAssignments || []).map((item) => (item.id || item.assignmentId) === assignmentId ? { ...item, status } : item),
    }))
  }
  const deleteAssignment = (assignmentId) => {
    if (!canDelete) return
    setData((current) => ({
      ...current,
      productionAssignments: (current.productionAssignments || []).filter((item) => (item.id || item.assignmentId) !== assignmentId),
    }))
  }
  const selectWholeTeam = () => updateForm('employeeCodes', selectableEmployees.map((employee) => employee.code))
  const clearSelectedEmployees = () => updateForm('employeeCodes', [])
  const selectAllMachines = () => updateForm('machineCodes', machines.map((machine) => machine.machineCode))
  const clearSelectedMachines = () => updateForm('machineCodes', [])
  const toggleEmployee = (employeeCode) => {
    updateForm('employeeCodes', (form.employeeCodes || []).includes(employeeCode)
      ? (form.employeeCodes || []).filter((code) => code !== employeeCode)
      : [...(form.employeeCodes || []), employeeCode])
  }
  const toggleMachine = (machineCode) => {
    updateForm('machineCodes', (form.machineCodes || []).includes(machineCode)
      ? (form.machineCodes || []).filter((code) => code !== machineCode)
      : [...(form.machineCodes || []), machineCode])
  }

  return (
    <div className="page-content production-assignment-page">
      <section className="panel">
        <div className="section-heading-row">
          <div>
            <span className="section-kicker">Sản xuất</span>
            <h2>Phân công nhân sự</h2>
            <p className="panel-text">Dữ liệu phục vụ nhận diện nhân sự đang trực ở công đoạn, sẵn sàng nối QR nhân viên và cân điện tử sau này.</p>
          </div>
          <button className="primary-button" type="button" disabled={!canCreate} onClick={saveAssignment}>Lưu phân công</button>
        </div>
        <div className="production-form-grid order-create-form">
          <label>Ngày<input type="date" value={form.date} onChange={(event) => updateForm('date', event.target.value)} /></label>
          <label>Ca làm việc<select value={form.shiftCode} onChange={(event) => updateForm('shiftCode', event.target.value)}>{shifts.map((shift) => <option key={shift.code} value={shift.code}>{shift.code} / {shift.name}</option>)}</select></label>
          <label>Tổ sản xuất<select value={form.teamCode} onChange={(event) => updateForm('teamCode', event.target.value)}>{availableTeams.map((team) => <option key={team.code} value={team.code}>{team.code} / {team.name}</option>)}</select></label>
          <label>Công đoạn<select value={form.stage} onChange={(event) => updateForm('stage', event.target.value)}>{allowedStages.map((stage) => <option key={stage} value={stage}>{stage}</option>)}</select></label>
          {form.stage === 'Phối trộn' && (
            <div className="assignment-employee-field">
              <span>Máy phối trộn</span>
              <div className="assignment-employee-checklist">
                {machines.length === 0 && <div className="assignment-empty">Chưa có máy phối trộn khả dụng</div>}
                {machines.map((machine) => (
                  <label className="employee-row" key={machine.machineCode}>
                    <input
                      type="checkbox"
                      checked={(form.machineCodes || []).includes(machine.machineCode)}
                      onChange={() => toggleMachine(machine.machineCode)}
                    />
                    <span>{mixingMachineOptionLabel(machine)}</span>
                  </label>
                ))}
              </div>
              <div className="assignment-mini-actions">
                <button className="secondary-button" type="button" disabled={!canCreate || machines.length === 0} onClick={selectAllMachines}>Chọn tất cả máy</button>
                <button className="secondary-button" type="button" disabled={!canCreate || (form.machineCodes || []).length === 0} onClick={clearSelectedMachines}>Bỏ chọn máy</button>
              </div>
            </div>
          )}
          <div className={`assignment-employee-field ${form.stage === 'Phối trộn' ? '' : 'wide-field'}`}>
            <span>Nhân viên</span>
            <div className="assignment-employee-checklist">
              {selectableEmployees.length === 0 && <div className="assignment-empty">Chưa có nhân viên phù hợp</div>}
              {selectableEmployees.map((employee) => (
                <label className="employee-row" key={employee.code}>
                  <input
                    type="checkbox"
                    checked={(form.employeeCodes || []).includes(employee.code)}
                    onChange={() => toggleEmployee(employee.code)}
                  />
                  <span>{employee.code} / {employee.name}</span>
                </label>
              ))}
            </div>
          </div>
        </div>
        <div className="action-row">
          <button className="secondary-button" type="button" disabled={!canCreate || selectableEmployees.length === 0} onClick={selectWholeTeam}>Chọn cả tổ</button>
          <button className="secondary-button" type="button" disabled={!canCreate || (form.employeeCodes || []).length === 0} onClick={clearSelectedEmployees}>Bỏ chọn</button>
        </div>
        {notice && <div className="process-alert">{notice}</div>}
      </section>
      <section className="panel">
        <h3>Bảng phân công</h3>
        <SimpleTable tableClassName="production-assignment-table" headers={['Ngày', 'Ca', 'Tổ sản xuất', 'Công đoạn', 'Máy được phân công', 'Nhân viên', 'Người phân công', 'Thời gian phân công', 'Trạng thái', 'Hành động']} rows={visibleAssignments.map((item) => (
          <tr key={item.id || item.assignmentId}>
            <td>{item.workDate || item.date}</td>
            <td>{item.shiftCode} / {item.shiftName}</td>
            <td>{item.teamName || item.productionTeamName || item.productionTeam || '-'}</td>
            <td>{item.processName || item.stage}</td>
            <td>{formatAssignmentMachines(item)}</td>
            <td>{formatAssignmentEmployees(item)}</td>
            <td>{item.assignedBy}</td>
            <td>{item.assignedAt}</td>
            <td>
              <select value={item.status} disabled={!canEdit} onChange={(event) => updateAssignmentStatus(item.id || item.assignmentId, event.target.value)}>
                <option>Chưa bắt đầu</option>
                <option>Đang thực hiện</option>
                <option>Hoàn thành</option>
                <option>Hủy</option>
              </select>
            </td>
            <td><button className="danger-button" type="button" disabled={!canDelete} onClick={() => deleteAssignment(item.id || item.assignmentId)}>Xóa</button></td>
          </tr>
        ))} empty="Chưa có phân công trong phạm vi được xem." />
      </section>
    </div>
  )
}

function MasterCatalogPage({ title, storageKey, fields, labels, data, setData, permissions = [], permissionKey }) {
  const rows = data[storageKey] || []
  const key = permissionKey || storageKey.replace('Catalog', '')
  const canCreate = hasPermission(permissions, `master.${key}.create`)
  const canEdit = hasPermission(permissions, `master.${key}.edit`)
  const canDelete = hasPermission(permissions, `master.${key}.delete`)
  const updateRow = (rowId, field, value) => {
    if (!canEdit) return
    setData((current) => ({
      ...current,
      [storageKey]: (current[storageKey] || []).map((row) => row.id === rowId ? { ...row, [field]: value } : row),
    }))
  }
  const addRow = () => {
    if (!canCreate) return
    const next = {
      id: uid(storageKey.replace('Catalog', '').toUpperCase()),
      code: '',
      name: '',
      status: 'Hoạt động',
      note: '',
    }
    fields.forEach((field) => {
      if (next[field] == null) next[field] = ''
    })
    setData((current) => ({ ...current, [storageKey]: [next, ...(current[storageKey] || [])] }))
  }
  const deleteRow = (rowId) => {
    if (!canDelete) return
    setData((current) => ({ ...current, [storageKey]: (current[storageKey] || []).filter((row) => row.id !== rowId) }))
  }

  return (
    <div className="page-content">
      <section className="panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Dữ liệu gốc</span><h2>{title}</h2></div>
          <button className="primary-button" type="button" disabled={!canCreate} onClick={addRow}>Thêm mới</button>
        </div>
        <div className="table-wrapper">
          <table className="admin-wide-table">
            <thead>
              <tr>
                {labels.map((label) => <th key={label}>{label}</th>)}
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.id}>
                  {fields.map((field, index) => (
                    <td key={field} data-label={labels[index]}>
                      <input value={row[field] || ''} readOnly={!canEdit} onChange={(event) => updateRow(row.id, field, event.target.value)} />
                    </td>
                  ))}
                  <td data-label="Hành động">
                    <button className="danger-button" type="button" disabled={!canDelete} onClick={() => deleteRow(row.id)}>Xóa</button>
                  </td>
                </tr>
              ))}
              {rows.length === 0 && <tr><td className="empty-row" colSpan={fields.length + 1}>Chưa có dữ liệu.</td></tr>}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  )
}

function AdminPage({ authData, setAuthData, section = 'users' }) {
  const [showCreateRole, setShowCreateRole] = useState(false)
  const [newRoleName, setNewRoleName] = useState('')
  const [notice, setNotice] = useState('')
  const baseRoles = [
    ['Admin', 'Admin'],
    ['Kho NL', 'Kho NL'],
    ['Kỹ thuật', 'Kỹ thuật'],
    ['Sản xuất', 'Sản xuất'],
    ['QC', 'QC'],
    ['Cân hóa', 'Cân hóa'],
    ['Cân rắn', 'Cân rắn'],
    ['Phối trộn', 'Phối trộn'],
    ['Đóng gói', 'Đóng gói'],
    ['Kho TP', 'Kho TP'],
    ['Quản đốc', 'Quản đốc'],
    ['Ban giám đốc', 'Ban giám đốc'],
  ]
  const roles = [
    ...baseRoles.filter(([role]) => authData.roles?.[role]),
  ]
  const operationPermissionRows = defaultNavItems
    .filter((item) => item.permission && item.type !== 'group')
    .map((item) => ({ label: item.label, permissions: [{ action: 'Truy cập', id: item.permission }] }))
  const masterPermissionRows = masterPermissionGroups.map(([key, label]) => ({
    label,
    permissions: CRUD_ACTIONS.map((action) => ({ action: action[0].toUpperCase() + action.slice(1), id: `master.${key}.${action}` })),
  }))
  const productionPermissionRows = productionPermissionGroups.map(([key, label]) => ({
    label,
    permissions: CRUD_ACTIONS.map((action) => ({ action: action[0].toUpperCase() + action.slice(1), id: `production.${key}.${action}` })),
  }))
  const productionExtraPermissionRows = [
    { label: 'Sản xuất / Truy xuất lô sản xuất', permissions: [{ action: 'View', id: 'production.trace.view' }] },
    { label: 'Sản xuất / Nhật ký thao tác', permissions: [{ action: 'View', id: 'production.log.view' }] },
  ]
  const systemPermissionRows = [
    { label: 'Quản trị hệ thống / Người dùng', permissions: [{ action: 'Truy cập', id: 'admin' }] },
    { label: 'Quản trị hệ thống / Vai trò', permissions: [{ action: 'Truy cập', id: 'admin' }] },
    { label: 'Quản trị hệ thống / Ma trận phân quyền', permissions: [{ action: 'Truy cập', id: 'admin' }] },
    { label: 'Quản trị hệ thống / Nhật ký hệ thống', permissions: [{ action: 'Truy cập', id: 'admin' }] },
    { label: 'Bảo mật công thức / Xem tỷ lệ đầy đủ', permissions: [{ action: 'Secure view', id: 'formula.secure.view' }] },
  ]
  const permissionRows = [...operationPermissionRows, ...productionPermissionRows, ...productionExtraPermissionRows, ...masterPermissionRows, ...systemPermissionRows]
  const allPermissionIds = allSystemPermissionIds
  const withAdminPermissions = (nextAuth) => ({
    ...nextAuth,
    roles: {
      ...nextAuth.roles,
      Admin: allSystemPermissionIds,
    },
  })
  const updateAuth = (nextAuth, message = 'Đã cập nhật phân quyền.') => {
    setAuthData(withAdminPermissions(nextAuth))
    setNotice(message)
  }
  const updateUser = (username, updates) => {
    const nextUsers = (authData.users || []).map((user) => (
      user.username === username
        ? { ...user, ...updates, department: updates.role || user.department || user.role }
        : user
    ))
    updateAuth({ ...authData, users: nextUsers }, 'Đã cập nhật người dùng.')
  }
  const addUser = () => {
    const nextNo = (authData.users || []).length + 1
    const username = `user${nextNo}`
    updateAuth({
      ...authData,
      users: [
        { username, password: DEFAULT_PASSWORD, role: 'Sản xuất', fullName: 'Người dùng mới', department: 'Sản xuất', status: ACTIVE_STATUS },
        ...(authData.users || []),
      ],
    }, `Đã thêm người dùng ${username}.`)
  }
  const toggleUserLock = (username) => {
    const user = (authData.users || []).find((item) => item.username === username)
    if (!user || user.username === 'admin') return
    updateUser(username, { status: user.status === ACTIVE_STATUS ? LOCKED_STATUS : ACTIVE_STATUS })
  }
  const resetPassword = (username) => {
    const nextPassword = window.prompt('Mật khẩu mới:', DEFAULT_PASSWORD)?.trim()
    if (!nextPassword) return
    updateUser(username, { password: nextPassword })
    setNotice(`Đã đặt lại mật khẩu cho ${username}.`)
  }
  const resetAllPasswords = () => {
    updateAuth({
      ...authData,
      users: (authData.users || []).map((user) => ({ ...user, password: DEFAULT_PASSWORD })),
    }, 'Đã đặt lại toàn bộ mật khẩu về 123456')
  }
  const togglePermission = (role, id) => {
    if (role === 'Admin') return
    const permissions = authData.roles[role] || []
    const nextPermissions = permissions.includes(id) ? permissions.filter((item) => item !== id) : [...permissions, id]
    updateAuth({ ...authData, roles: { ...authData.roles, [role]: nextPermissions } })
  }
  const setRolePermissions = (role, permissions) => {
    if (role === 'Admin') return
    updateAuth({ ...authData, roles: { ...authData.roles, [role]: permissions } })
  }
  const createRole = () => {
    const role = newRoleName.trim()
    if (!role || authData.roles[role]) return
    updateAuth({ ...authData, roles: { ...authData.roles, [role]: [] } }, `Đã tạo vai trò ${role}.`)
    setNewRoleName('')
    setShowCreateRole(false)
  }
  const renameRole = (role) => {
    if (role === 'Admin') return
    const nextName = window.prompt('Tên vai trò mới:', role)?.trim()
    if (!nextName || nextName === role || authData.roles[nextName]) return
    const nextRoles = {}
    Object.entries(authData.roles).forEach(([name, permissions]) => {
      nextRoles[name === role ? nextName : name] = permissions
    })
    updateAuth({
      ...authData,
      roles: nextRoles,
      users: (authData.users || []).map((user) => user.role === role ? { ...user, role: nextName, department: user.department === role ? nextName : user.department } : user),
    }, `Đã đổi tên vai trò ${role} thành ${nextName}.`)
  }
  const deleteRole = (role) => {
    if (role === 'Admin') return
    if (!window.confirm(`Xóa vai trò ${role}? Người dùng thuộc vai trò này sẽ không còn quyền menu cho tới khi đổi vai trò.`)) return
    const nextRoles = Object.fromEntries(Object.entries(authData.roles).filter(([name]) => name !== role))
    updateAuth({ ...authData, roles: nextRoles }, `Đã xóa vai trò ${role}.`)
  }
  const savePermissions = () => {
    localStorage.setItem(AUTH_KEY, JSON.stringify(withAdminPermissions(authData)))
    setNotice('Đã lưu phân quyền. Thay đổi có hiệu lực khi người dùng đăng nhập lại.')
  }
  return (
    <div className={`page-content admin-page ${section === 'permissions' ? 'permission-matrix-page' : ''}`}>
      {section === 'users' && <section className="panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Quản trị hệ thống</span><h2>Danh sách người dùng</h2></div>
          <div className="action-row">
            <button className="secondary-button" type="button" onClick={resetAllPasswords}>Đặt lại toàn bộ mật khẩu</button>
            <button className="primary-button" type="button" onClick={addUser}>Thêm người dùng</button>
          </div>
        </div>
        {notice && <p className="empty-alert">{notice}</p>}
        <div className="table-wrapper">
          <table className="admin-wide-table user-admin-table">
            <thead>
              <tr>
                <th>Tài khoản</th>
                <th>Mật khẩu</th>
                <th>Họ tên</th>
                <th>Vai trò</th>
                <th>Trạng thái</th>
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {(authData.users || []).map((user) => (
                <tr key={user.username}>
                  <td data-label="Tài khoản">
                    <input
                      value={user.username}
                      disabled={defaultUsers.some((item) => item.username === user.username)}
                      onChange={(event) => updateUser(user.username, { username: event.target.value.trim() })}
                    />
                  </td>
                  <td data-label="Mật khẩu"><input value={user.password} onChange={(event) => updateUser(user.username, { password: event.target.value })} /></td>
                  <td data-label="Họ tên"><input value={user.fullName || ''} onChange={(event) => updateUser(user.username, { fullName: event.target.value })} /></td>
                  <td data-label="Vai trò">
                    <select value={user.role} onChange={(event) => updateUser(user.username, { role: event.target.value })}>
                      {officialRoleNames.map((role) => <option key={role} value={role}>{role}</option>)}
                    </select>
                  </td>
                  <td data-label="Trạng thái"><span className={`status-pill ${user.status === ACTIVE_STATUS ? 'pass' : 'locked'}`}>{user.status}</span></td>
                  <td data-label="Hành động">
                    <div className="user-action-group">
                      <button type="button" className="secondary-button" onClick={() => resetPassword(user.username)}>Đặt lại mật khẩu</button>
                      <button type="button" className={user.status === ACTIVE_STATUS ? 'danger-button' : 'secondary-button'} disabled={user.username === 'admin'} onClick={() => toggleUserLock(user.username)}>
                        {user.status === ACTIVE_STATUS ? 'Khóa' : 'Mở khóa'}
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>}
      {section === 'roles' && <section className="panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Quản trị hệ thống</span><h2>Vai trò</h2></div>
          <button className="primary-button" type="button" onClick={() => setShowCreateRole(true)}>Thêm vai trò</button>
        </div>
        {notice && <p className="empty-alert">{notice}</p>}
        <div className="table-wrapper">
          <table className="admin-wide-table">
            <thead>
              <tr>
                <th>Vai trò</th>
                <th>Số quyền</th>
                <th>Số người dùng</th>
                <th>Hành động</th>
              </tr>
            </thead>
            <tbody>
              {roles.map(([role, label]) => (
                <tr key={role}>
                  <td>{label}</td>
                  <td>{role === 'Admin' ? allPermissionIds.length : (authData.roles[role] || []).length}</td>
                  <td>{(authData.users || []).filter((user) => user.role === role).length}</td>
                  <td>
                    <div className="user-action-group">
                      <button type="button" className="secondary-button" disabled={role === 'Admin'} onClick={() => renameRole(role)}>Sửa tên</button>
                      <button type="button" className="danger-button" disabled={role === 'Admin'} onClick={() => deleteRole(role)}>Xóa</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>}
      {section === 'permissions' && <section className="panel permission-matrix-panel">
        <div className="section-heading-row">
          <div><span className="section-kicker">Quản trị hệ thống</span><h2>Phân quyền V3</h2></div>
          <div className="permission-toolbar">
            <button className="primary-button" type="button" onClick={savePermissions}>Lưu phân quyền</button>
          </div>
        </div>
        {notice && <p className="empty-alert">{notice}</p>}
        <div className="permission-matrix-wrapper">
          <table className="permission-matrix">
            <thead>
              <tr>
                <th>Chức năng</th>
                {roles.map(([role, label]) => (
                  <th key={role}>
                    <div className="role-column-header">
                      <strong>{label}</strong>
                    </div>
                    <div className="role-quick-actions">
                      <button type="button" disabled={role === 'Admin'} onClick={() => setRolePermissions(role, allPermissionIds)}>✓ Chọn tất cả</button>
                      <button type="button" disabled={role === 'Admin'} onClick={() => setRolePermissions(role, [])}>✕ Bỏ tất cả</button>
                    </div>
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {permissionRows.map((item) => (
                <tr key={item.label}>
                  <td>
                    <strong>{item.label}</strong>
                    <div className="permission-action-row">
                      {item.permissions.map((permission) => <span key={permission.id}>{permission.action}</span>)}
                    </div>
                  </td>
                  {roles.map(([role]) => {
                    const checked = role === 'Admin' || item.permissions.every((permission) => (authData.roles[role] || []).includes(permission.id))
                    const disabled = role === 'Admin'
                    return (
                      <td key={`${role}-${item.label}`}>
                        <div className="permission-action-checks">
                          {item.permissions.map((permission) => (
                            <input
                              key={permission.id}
                              type="checkbox"
                              checked={role === 'Admin' || (authData.roles[role] || []).includes(permission.id)}
                              disabled={disabled}
                              title={`${permission.action}: ${permission.id}`}
                              onChange={() => togglePermission(role, permission.id)}
                            />
                          ))}
                        </div>
                      </td>
                    )
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>}
      {section === 'roles' && showCreateRole && (
        <div className="modal-backdrop" role="presentation">
          <div className="mixing-modal role-modal" role="dialog" aria-modal="true">
            <div className="modal-header">
              <div><span className="section-kicker">Vai trò</span><h2>Tạo vai trò mới</h2></div>
              <button type="button" className="icon-button" onClick={() => setShowCreateRole(false)} aria-label="Đóng">×</button>
            </div>
            <label>Tên vai trò mới<input value={newRoleName} onChange={(event) => setNewRoleName(event.target.value)} autoFocus /></label>
            <div className="modal-actions">
              <button type="button" className="secondary-button" onClick={() => setShowCreateRole(false)}>Hủy</button>
              <button type="button" className="primary-button" onClick={createRole}>Tạo</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}

const pageMeta = {
  dashboard: ['Dashboard', 'Báo cáo nhanh quy trình sản xuất V3'],
  'raw-materials': ['Kho nguyên liệu', 'Nhập NVL, tạo QR/Barcode và lưu localStorage'],
  formulas: ['Công thức gốc', 'Phòng kỹ thuật quản lý công thức gốc'],
  orders: ['Lệnh sản xuất', 'Tạo lệnh từ công thức gốc và chờ QC sản xuất thử'],
  'production-assignments': ['Phân công nhân sự', 'Theo dõi nhân sự trực công đoạn trong từng ca sản xuất'],
  qc: ['QC sản xuất thử', 'QC sản xuất thử - Kiểm tra và hiệu chỉnh lệnh SX trước sản xuất chính thức'],
  chemical: ['Tổ cân hóa', 'Cân hóa chất chính và bổ sung'],
  solid: ['Tổ cân rắn', 'Cân nguyên liệu rắn chính và bổ sung'],
  mixing: ['Tổ phối trộn', 'Phối trộn chính và phối trộn bổ sung'],
  'finished-qc': ['QC thành phẩm', 'QC2 sau phối trộn và điều chỉnh màu'],
  packaging: ['Đóng gói', 'Nhận lệnh QC thành phẩm OK để đóng gói'],
  'finished-goods': ['Kho thành phẩm', 'Nhập kho TP và hoàn thành lệnh'],
  logs: ['Nhật ký sản xuất', 'Lịch sử thao tác toàn quy trình'],
  reports: ['Báo cáo', 'Dashboard và báo cáo V3'],
  'reports-production': ['Báo cáo sản xuất', 'Tổng hợp pipeline, lệnh sản xuất và sản lượng'],
  'reports-trace': ['Truy xuất lô sản xuất', 'Tra cứu người thao tác, công đoạn, thời gian và máy theo LOT'],
  'reports-staff-performance': ['Hiệu suất nhân sự', 'Theo dõi phân công và thao tác của nhân sự sản xuất'],
  'reports-machine-performance': ['Hiệu suất máy', 'Theo dõi trạng thái và lệnh phối trộn theo máy'],
  admin: ['Quản trị hệ thống', 'Vai trò và phân quyền'],
  'admin-users': ['Người dùng', 'Quản lý tài khoản hệ thống'],
  'admin-roles': ['Vai trò', 'Quản lý vai trò hệ thống'],
  'admin-permissions': ['Ma trận phân quyền', 'Phân quyền chi tiết theo chức năng'],
  'admin-system-logs': ['Nhật ký hệ thống', 'Truy xuất nhật ký thao tác theo người dùng, nhân viên và công đoạn'],
  'master-materials': ['Danh mục vật tư', 'Dữ liệu gốc vật tư dùng trong sản xuất'],
  'master-products': ['Danh mục sản phẩm', 'Dữ liệu gốc sản phẩm'],
  'master-suppliers': ['Danh mục nhà cung cấp', 'Dữ liệu gốc nhà cung cấp'],
  'master-customers': ['Danh mục khách hàng', 'Dữ liệu gốc khách hàng'],
  'master-employees': ['Danh sách nhân viên', 'Nhân sự vận hành trong phạm vi sản xuất'],
  'master-teams': ['Danh mục tổ sản xuất', 'Tổ vận hành dùng cho phân công sản xuất'],
  'master-shifts': ['Danh mục ca làm việc', 'Ca làm việc dùng cho phân công sản xuất'],
  'admin-machines': ['Danh mục máy phối trộn', 'Khai báo máy, trạng thái và đề nghị đổi máy'],
}

function App() {
  const [data, setData] = useState(() => {
    const seed = seedData()
    const saved = loadStored(DATA_KEY, seed)
    const storedFormulas = loadStored(FORMULAS_KEY, null)
    const formulas = nonEmptyArray(storedFormulas, saved.formulas, seed.formulas)
    const productionLogs = nonEmptyArray(loadStored(PRODUCTION_LOGS_KEY, null), saved.productionLogs, saved.logs, seed.productionLogs)
    const qc2Logs = nonEmptyArray(loadStored(QC2_LOGS_KEY, null), saved.qc2Logs)
    const qc2AdjustmentTickets = nonEmptyArray(loadStored(QC2_ADJUSTMENTS_KEY, null), saved.qc2AdjustmentTickets, saved.qc2Adjustments)
    const supplementalWeighing = nonEmptyArray(loadStored(SUPPLEMENTAL_WEIGHING_KEY, null), saved.supplementalWeighing)
    const weighedContainers = normalizeWeighedContainers(nonEmptyArray(loadStored(WEIGHED_CONTAINERS_KEY, null), saved.weighedContainers))
    const packingLogs = nonEmptyArray(loadStored(PACKING_LOGS_KEY, null), saved.packingLogs)
    const finishedGoods = normalizeFinishedGoodsData(nonEmptyArray(loadStored(FINISHED_GOODS_KEY, null), saved.finishedGoods))
    const rawMaterials = normalizeRawMaterialLots(nonEmptyArray(saved.rawMaterials, seed.rawMaterials))
    const materialCatalog = deriveMaterialCatalog({
      formulas,
      rawMaterials,
      materialCatalog: nonEmptyArray(loadStored(MATERIAL_CATALOG_KEY, null), saved.materialCatalog),
    })
    const storedOrders = loadStored(PRODUCTION_ORDERS_KEY, null)
    const orderSource = nonEmptyArray(storedOrders, saved.productionOrders, saved.orders, seed.orders)
    const orders = ensureQcDemoOrders(
      normalizeProductionOrders(orderSource, formulas),
      normalizeProductionOrders(seed.orders, formulas),
    )
    const baseData = {
      ...seed,
      ...saved,
      formulas,
      orders,
      formulaVersions: saved.formulaVersions || [],
      qc2Logs,
      qc2AdjustmentTickets,
      supplementalWeighing,
      weighedContainers,
      packingLogs,
      finishedGoods,
      rawMaterials,
      materialCatalog,
      employeeCatalog: productionEmployeeCatalog,
      teamCatalog: [
        ...(saved.teamCatalog || []),
        ...seed.teamCatalog.filter((seedTeam) => !(saved.teamCatalog || []).some((team) => team.code === seedTeam.code)),
      ],
      stockTransactions: saved.stockTransactions || [],
      mixingMachines: normalizeMixingMachines(saved.mixingMachines),
      productionLogs,
    }
    return weighedContainers.length ? baseData : applyDemoQrData(baseData)
  })
  const [authData, setAuthData] = useState(() => {
    return normalizeAuthData(loadStored(AUTH_KEY, defaultAuth))
  })
  const [currentUser, setCurrentUser] = useState(() => {
    const session = loadStored(SESSION_KEY, null)
    return session ? normalizeAuthData(loadStored(AUTH_KEY, defaultAuth)).users.find((user) => user.username === session.username) : null
  })
  const [loginError, setLoginError] = useState('')
  const initialPage = new URLSearchParams(window.location.search).get('page') || 'dashboard'
  const [selectedPage, setSelectedPage] = useState(initialPage)
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false)

  useEffect(() => {
    const orders = normalizeProductionOrders(data.orders || [], data.formulas || [])
    localStorage.setItem(PRODUCTION_ORDERS_KEY, JSON.stringify(orders))
    localStorage.setItem(FORMULAS_KEY, JSON.stringify(data.formulas || []))
    localStorage.setItem(PRODUCTION_LOGS_KEY, JSON.stringify(data.productionLogs || data.logs || []))
    localStorage.setItem(QC2_LOGS_KEY, JSON.stringify(data.qc2Logs || []))
    localStorage.setItem(QC2_ADJUSTMENTS_KEY, JSON.stringify(data.qc2AdjustmentTickets || data.qc2Adjustments || []))
    localStorage.setItem(SUPPLEMENTAL_WEIGHING_KEY, JSON.stringify(data.supplementalWeighing || []))
    localStorage.setItem(WEIGHED_CONTAINERS_KEY, JSON.stringify(normalizeWeighedContainers(data.weighedContainers || [])))
    localStorage.setItem(PACKING_LOGS_KEY, JSON.stringify(data.packingLogs || []))
    localStorage.setItem(FINISHED_GOODS_KEY, JSON.stringify(normalizeFinishedGoodsData(data.finishedGoods || [])))
    localStorage.setItem(MATERIAL_CATALOG_KEY, JSON.stringify(deriveMaterialCatalog(data)))
    localStorage.setItem(DATA_KEY, JSON.stringify({ ...data, materialCatalog: deriveMaterialCatalog(data), rawMaterials: normalizeRawMaterialLots(data.rawMaterials || []), orders, productionOrders: orders }))
  }, [data])
  useEffect(() => { localStorage.setItem(AUTH_KEY, JSON.stringify(authData)) }, [authData])

  const user = currentUser && authData.users.find((item) => item.username === currentUser.username)
  const navItems = useMemo(() => {
    const permissions = authData.roles[user?.role] || []
    const visibleChildren = new Set(defaultNavItems
      .filter((item) => item.parentId && hasPermission(permissions, pagePermission(item)))
      .map((item) => item.parentId))
    return defaultNavItems.filter((item) => {
      if (!item.parentId && item.type === 'group') return visibleChildren.has(item.id)
      return hasPermission(permissions, pagePermission(item))
    })
  }, [authData.roles, user])
  const userPermissions = authData.roles[user?.role] || []
  const visiblePageIds = useMemo(() => navItems.filter((item) => item.type !== 'group').map((item) => item.id), [navItems])
  const page = visiblePageIds.includes(selectedPage) ? selectedPage : visiblePageIds[0] || 'dashboard'
  const [title, subtitle] = pageMeta[page] || pageMeta.dashboard

  const login = (username, password) => {
    const normalizedUsername = username.trim()
    const found = authData.users.find((item) => item.username === normalizedUsername && item.password === password && item.status === ACTIVE_STATUS)
    if (!found) { setLoginError('Sai tài khoản, mật khẩu hoặc tài khoản bị khóa.'); return }
    localStorage.setItem(SESSION_KEY, JSON.stringify({ username: found.username, loggedAt: nowText() }))
    setAuthData((current) => ({
      ...current,
      accessLogs: [
        ...(current.accessLogs || []),
        { id: uid('access'), time: nowText(), username: found.username, action: 'Đăng nhập', role: found.role },
      ],
    }))
    setCurrentUser(found)
    setLoginError('')
  }
  const logout = () => {
    localStorage.removeItem(SESSION_KEY)
    setCurrentUser(null)
    setMobileMenuOpen(false)
  }
  const changePage = (nextPage) => {
    setSelectedPage(nextPage)
    setMobileMenuOpen(false)
  }

  if (!user) return <LoginPage onLogin={login} error={loginError} />

  const pages = {
    dashboard: <DashboardPage data={data} />,
    'raw-materials': <RawMaterialsPage data={data} setData={setData} />,
    formulas: <FormulasPage data={data} setData={setData} permissions={userPermissions} />,
    orders: <OrdersPage data={data} setData={setData} permissions={userPermissions} />,
    'production-assignments': <ProductionAssignmentPage data={data} setData={setData} user={user} permissions={userPermissions} />,
    qc: <QCPage data={data} setData={setData} user={user} />,
    chemical: <WeighingPage data={data} setData={setData} group={CHEMICAL} user={user} />,
    solid: <WeighingPage data={data} setData={setData} group={SOLID} user={user} />,
    mixing: <MixingPage data={data} setData={setData} user={user} />,
    'finished-qc': <FinishedProductQcPage data={data} setData={setData} user={user} />,
    packaging: <PackagingPage data={data} setData={setData} user={user} />,
    'finished-goods': <FinishedGoodsPage data={data} setData={setData} />,
    logs: <LogsPage data={data} />,
    reports: <ReportsPage data={data} />,
    'reports-production': <ReportsPage data={data} initialTab="production" lockedTab />,
    'reports-trace': <ReportsPage data={data} initialTab="trace" lockedTab />,
    'reports-staff-performance': <StaffPerformanceReportPage data={data} />,
    'reports-machine-performance': <MachinePerformanceReportPage data={data} />,
    admin: <AdminPage authData={authData} setAuthData={setAuthData} section="users" />,
    'admin-users': <AdminPage authData={authData} setAuthData={setAuthData} section="users" />,
    'admin-roles': <AdminPage authData={authData} setAuthData={setAuthData} section="roles" />,
    'admin-permissions': <AdminPage authData={authData} setAuthData={setAuthData} section="permissions" />,
    'admin-system-logs': <SystemLogsPage data={data} />,
    'master-materials': <MasterCatalogPage title="Danh mục vật tư" storageKey="materialCatalog" fields={['materialCode', 'materialName', 'materialGroup', 'unit']} labels={['Mã vật tư', 'Tên vật tư', 'Nhóm', 'Đơn vị']} data={data} setData={setData} permissions={userPermissions} />,
    'master-products': <MasterCatalogPage title="Danh mục sản phẩm" storageKey="productCatalog" fields={['code', 'name', 'group', 'unit', 'status', 'note']} labels={['Mã sản phẩm', 'Tên sản phẩm', 'Nhóm', 'Đơn vị', 'Trạng thái', 'Ghi chú']} data={data} setData={setData} permissions={userPermissions} />,
    'master-suppliers': <MasterCatalogPage title="Danh mục nhà cung cấp" storageKey="supplierCatalog" fields={['code', 'name', 'phone', 'address', 'status', 'note']} labels={['Mã NCC', 'Tên NCC', 'Điện thoại', 'Địa chỉ', 'Trạng thái', 'Ghi chú']} data={data} setData={setData} permissions={userPermissions} />,
    'master-customers': <MasterCatalogPage title="Danh mục khách hàng" storageKey="customerCatalog" fields={['code', 'name', 'phone', 'address', 'status', 'note']} labels={['Mã KH', 'Tên KH', 'Điện thoại', 'Địa chỉ', 'Trạng thái', 'Ghi chú']} data={data} setData={setData} permissions={userPermissions} />,
    'master-employees': <MasterCatalogPage title="Danh sách nhân viên" storageKey="employeeCatalog" fields={['code', 'name', 'productionTeam', 'title', 'operationRole', 'status', 'qrEmployee']} labels={['Mã NV', 'Họ tên', 'Tổ sản xuất', 'Chức danh', 'Vai trò vận hành', 'Trạng thái', 'QR nhân viên']} data={data} setData={setData} permissions={userPermissions} />,
    'master-teams': <MasterCatalogPage title="Danh mục tổ sản xuất" storageKey="teamCatalog" permissionKey="team" fields={['code', 'name', 'leader', 'note', 'status']} labels={['Mã tổ', 'Tên tổ', 'Tổ trưởng', 'Ghi chú', 'Trạng thái']} data={data} setData={setData} permissions={userPermissions} />,
    'master-shifts': <MasterCatalogPage title="Danh mục ca làm việc" storageKey="shiftCatalog" permissionKey="shift" fields={['code', 'name', 'startTime', 'endTime', 'note', 'status']} labels={['Mã ca', 'Tên ca', 'Giờ bắt đầu', 'Giờ kết thúc', 'Ghi chú', 'Trạng thái']} data={data} setData={setData} permissions={userPermissions} />,
    'admin-machines': <MixingMachineCatalogPage data={data} setData={setData} user={user} permissions={userPermissions} />,
  }

  return (
    <div className="app-shell">
      {mobileMenuOpen && <div className="mobile-drawer-overlay" onClick={() => setMobileMenuOpen(false)} />}
      <Sidebar selected={page} onChange={changePage} navItems={navItems} className="desktop-sidebar" />
      <Sidebar
        selected={page}
        onChange={changePage}
        navItems={navItems}
        className={`mobile-sidebar ${mobileMenuOpen ? 'mobile-open' : ''}`}
        variant="mobile"
        user={user}
        onClose={() => setMobileMenuOpen(false)}
        onLogout={logout}
      />
      <main className="main-content">
        <header className="mobile-header">
          <button type="button" className="mobile-menu-button" onClick={() => setMobileMenuOpen(true)} aria-label="Mở menu">☰</button>
          <div className="mobile-brand">
            <img src="/logo-sonhoabinh.png" alt="Sơn Hòa Bình" className="mobile-brand-logo" />
            <div className="mobile-brand-text">
              <strong>SƠN HÒA BÌNH</strong>
              <span>{title}</span>
            </div>
          </div>
          <div className="mobile-user">
            <span>{user.fullName || user.username}</span>
            <button className="secondary-button" onClick={logout}>Đăng xuất</button>
          </div>
        </header>
        <TopBar title={title} subtitle={subtitle} user={user} onLogout={logout} />
        {pages[page] || pages.dashboard}
      </main>
    </div>
  )
}

export default App




