import rawCustomerCatalog from '../../danh_muc_khach_hang.json'

export const customerCatalog = rawCustomerCatalog

export function normalizeText(value) {
  return String(value || '')
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/đ/g, 'd')
    .replace(/Đ/g, 'd')
    .replace(/\s+/g, ' ')
    .trim()
}

export function filterCustomerCatalog(customers = [], customerSearch = '') {
  const keyword = normalizeText(customerSearch)
  const source = Array.isArray(customers) ? customers : []
  if (!keyword) return source.slice(0, 30)

  return source
    .filter((customer) => normalizeText([
      customer.customerCode,
      customer.code,
      customer.customerName,
      customer.name,
      customer.channelCode,
      customer.province,
    ].filter(Boolean).join(' ')).includes(keyword))
    .slice(0, 50)
}

export default customerCatalog
