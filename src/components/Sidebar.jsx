import { useEffect, useMemo, useState } from 'react'
import { defaultNavItems } from '../data/navigation.js'

export function Sidebar({
  selected,
  onChange,
  navItems = defaultNavItems,
  className = '',
  variant = 'desktop',
  user,
  onClose,
  onLogout,
}) {
  const isMobile = variant === 'mobile'
  const childrenByParent = useMemo(() => navItems.reduce((acc, item) => {
    if (!item.parentId) return acc
    acc[item.parentId] = [...(acc[item.parentId] || []), item]
    return acc
  }, {}), [navItems])
  const parentByChild = useMemo(() => navItems.reduce((acc, item) => {
    if (item.parentId) acc[item.id] = item.parentId
    return acc
  }, {}), [navItems])
  const rootItems = useMemo(() => navItems.filter((item) => !item.parentId), [navItems])
  const [openGroups, setOpenGroups] = useState(() => {
    const activeParent = parentByChild[selected]
    return activeParent ? { [activeParent]: true } : {}
  })

  useEffect(() => {
    const activeParent = parentByChild[selected]
    if (activeParent) setOpenGroups((current) => ({ ...current, [activeParent]: true }))
  }, [parentByChild, selected])

  const handleItemClick = (item) => {
    const children = childrenByParent[item.id] || []
    if (children.length || item.type === 'group') {
      setOpenGroups((current) => ({ ...current, [item.id]: !current[item.id] }))
      return
    }
    onChange(item.id)
  }

  return (
    <aside className={`sidebar ${className}`}>
      {isMobile ? (
        <>
          <div className="mobile-sidebar-header mobile-drawer-header">
            <img
              src="/logo-sonhoabinh.png"
              alt="Sơn Hòa Bình"
              className="sidebar-brand-logo mobile-sidebar-logo"
            />
            <div className="mobile-sidebar-brand mobile-title-box">
              <h2 className="mobile-drawer-title">SƠN HÒA BÌNH</h2>
              <p className="mobile-drawer-subtitle">Hệ thống quản lý sản xuất V3</p>
            </div>
            <button type="button" className="mobile-sidebar-close mobile-drawer-close" onClick={onClose} aria-label="Đóng menu">×</button>
          </div>
          <div className="mobile-user-box">
            <div className="mobile-user-name">{user?.fullName || user?.username}</div>
            <div className="mobile-user-role">{user?.role}</div>
            <button type="button" className="mobile-logout-button" onClick={onLogout}>Đăng xuất</button>
          </div>
        </>
      ) : (
        <div className="sidebar-brand">
          <img
            src="/logo-sonhoabinh.png"
            alt="Sơn Hòa Bình"
            className="sidebar-brand-logo"
          />
          <div className="sidebar-brand-text">
            <span className="sidebar-brand-title">SƠN HÒA BÌNH</span>
            <span className="sidebar-brand-subtitle">Hệ thống quản lý sản xuất</span>
          </div>
        </div>
      )}

      <nav className={isMobile ? 'sidebar-nav mobile-nav-list' : 'sidebar-nav'}>
        {rootItems.map((item) => {
          const children = childrenByParent[item.id] || []
          const childActive = children.some((child) => child.id === selected)
          const expanded = openGroups[item.id] || childActive
          return (
            <div className="nav-group" key={item.id}>
              <button
                type="button"
                className={`${isMobile ? 'mobile-nav-item' : 'nav-item'} nav-item ${children.length ? 'nav-parent' : ''} ${selected === item.id || childActive ? 'active' : ''}`}
                onClick={() => handleItemClick(item)}
                aria-expanded={children.length ? expanded : undefined}
              >
                <span>{item.label}</span>
                {children.length > 0 && <span className="nav-caret">{expanded ? '⌃' : '⌄'}</span>}
              </button>
              {children.length > 0 && expanded && (
                <div className="nav-submenu">
                  {children.map((child) => (
                    <button
                      type="button"
                      key={child.id}
                      className={`${isMobile ? 'mobile-nav-item' : 'nav-item'} nav-item nav-subitem ${selected === child.id ? 'active' : ''}`}
                      onClick={() => onChange(child.id)}
                    >
                      {child.label}
                    </button>
                  ))}
                </div>
              )}
            </div>
          )
        })}
      </nav>
    </aside>
  )
}
