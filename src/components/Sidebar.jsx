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

  return (
    <aside className={`sidebar ${className}`}>
      {isMobile ? (
        <>
          <div className="mobile-drawer-header">
            <div className="mobile-drawer-brand">
              <span className="mobile-logo-box">HB</span>
              <div>
                <div className="mobile-drawer-title">Sơn Hòa Bình</div>
                <div className="mobile-drawer-subtitle">Hệ thống quản lý sản xuất V3</div>
              </div>
            </div>
            <button type="button" className="mobile-drawer-close" onClick={onClose} aria-label="Đóng menu">×</button>
          </div>
          <div className="mobile-user-box">
            <div className="mobile-user-name">{user?.fullName || user?.username}</div>
            <div className="mobile-user-role">{user?.role}</div>
            <button type="button" className="mobile-logout-button" onClick={onLogout}>Đăng xuất</button>
          </div>
        </>
      ) : (
        <div className="sidebar-brand">
          <span>Sơn Hòa Bình</span>
        </div>
      )}

      <nav className={isMobile ? 'sidebar-nav mobile-nav-list' : 'sidebar-nav'}>
        {navItems.map((item) => (
          <button
            key={item.id}
            type="button"
            className={`${isMobile ? 'mobile-nav-item' : 'nav-item'} nav-item ${selected === item.id ? 'active' : ''}`}
            onClick={() => onChange(item.id)}
          >
            {item.label}
          </button>
        ))}
      </nav>
    </aside>
  )
}
