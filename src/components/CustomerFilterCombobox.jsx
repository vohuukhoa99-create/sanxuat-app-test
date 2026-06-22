import { useEffect, useMemo, useState } from 'react'
import { filterCustomerCatalog } from '../data/customerCatalog.js'

export function CustomerFilterCombobox({
  options = [],
  value = '',
  emptyValue = '',
  allLabel = 'Tất cả khách hàng',
  placeholder = 'Gõ tên khách hàng',
  onChange,
}) {
  const [open, setOpen] = useState(false)
  const [search, setSearch] = useState('')
  const customerOptions = useMemo(() => options.map((name, index) => ({
    id: `customer-filter-${index}`,
    customerName: name,
  })), [options])
  const filteredCustomers = useMemo(() => filterCustomerCatalog(customerOptions, search), [customerOptions, search])

  useEffect(() => {
    setSearch(value === emptyValue ? '' : value)
  }, [emptyValue, value])

  const selectValue = (nextValue) => {
    onChange(nextValue)
    setSearch(nextValue === emptyValue ? '' : nextValue)
    setOpen(false)
  }

  return (
    <div className="customer-combobox">
      <input
        value={search}
        placeholder={allLabel || placeholder}
        onFocus={() => setOpen(true)}
        onChange={(event) => {
          setSearch(event.target.value)
          setOpen(true)
        }}
        onBlur={() => window.setTimeout(() => setOpen(false), 120)}
        aria-autocomplete="list"
        aria-expanded={open}
      />
      {open && (
        <div className="customer-combobox-menu">
          <button type="button" onMouseDown={(event) => event.preventDefault()} onClick={() => selectValue(emptyValue)}>
            <span>{allLabel}</span>
          </button>
          {filteredCustomers.length ? filteredCustomers.map((customer) => (
            <button type="button" key={customer.id} onMouseDown={(event) => event.preventDefault()} onClick={() => selectValue(customer.customerName)}>
              <span>{customer.customerName}</span>
            </button>
          )) : <div className="customer-combobox-empty">Không có khách hàng phù hợp</div>}
        </div>
      )}
    </div>
  )
}
