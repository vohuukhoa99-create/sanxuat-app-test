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
