import '../styles/Sidebar.css';
import { useState } from 'react';
import { NavLink, useNavigate } from 'react-router-dom';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function Sidebar() {
  const [isOpen, setIsOpen] = useState(false);
  const navigate = useNavigate();

  // 🔐 Get user + role
  const storedUser = localStorage.getItem('user');
  const user = storedUser ? JSON.parse(storedUser) : null;
  const role = user?.role;

  const toggleSidebar = () => setIsOpen(!isOpen);
  const closeSidebar = () => setIsOpen(false);

  // Logout
  const handleLogout = async () => {
    try {
      await api.post(`${BASE_URL}/logout/`);
    } catch (err) {
      console.error('Logout error:', err);
    } finally {
      localStorage.removeItem('user');
      closeSidebar();
      navigate('/login');
    }
  };

  return (
    <>
      {/* Mobile Toggle */}
      <button
        className="sidebar-toggle"
        onClick={toggleSidebar}
        aria-label="Toggle navigation menu"
      >
        <span className="sidebar-toggle__icon">☰</span>
      </button>

      {/* Overlay */}
      <div
        className={`sidebar-overlay ${isOpen ? 'sidebar-overlay--active' : ''}`}
        onClick={closeSidebar}
      />

      {/* Sidebar */}
      <aside className={`sidebar ${isOpen ? 'sidebar--open' : ''}`}>
        <div className="sidebar__header">
          <h2 className="sidebar__title">Ticketing App</h2>
        </div>

        <nav className="sidebar__nav">
          <ul className="sidebar__menu">
            <div className="sidebar__menu-top">

              {/* 🏠 Dashboard - ALL USERS */}
              <li className="sidebar__menu-item">
                <NavLink
                  to="/dashboard"
                  end
                  className={({ isActive }) =>
                    `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`
                  }
                  onClick={closeSidebar}
                >
                 🏠 Home
                </NavLink>
              </li>

              {/* 🏢 Company Management - SUPER ADMIN */}
              {role === 'superadmin' && (
                <li className="sidebar__menu-item">
                  <NavLink
                    to="/dashboard/companies"
                    className={({ isActive }) =>
                      `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`
                    }
                    onClick={closeSidebar}
                  >
                    🏢 Company Management
                  </NavLink>
                </li>
              )}

              {/* 👥 User Management - SUPER ADMIN */}
              {role === 'superadmin' && (
                <li className="sidebar__menu-item">
                  <NavLink
                    to="/dashboard/users"
                    className={({ isActive }) =>
                      `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`
                    }
                    onClick={closeSidebar}
                  >
                    👥 User Management
                  </NavLink>
                </li>
              )}

              {/* 📝 Ticket Report - BRANCH ADMIN */}
              {role === 'branch_admin' && (
                <li className="sidebar__menu-item">
                  <NavLink
                    to="/dashboard/ticket-report"
                    className={({ isActive }) =>
                      `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`
                    }
                    onClick={closeSidebar}
                  >
                    📝 Ticket Report
                  </NavLink>
                </li>
              )}
            </div>

            {/* 🚪 Logout */}
            <div className="sidebar__menu-bottom">
              <li className="sidebar__menu-item">
                <button
                  onClick={handleLogout}
                  className="sidebar__link sidebar__link--logout"
                  style={{
                    width: '100%',
                    textAlign: 'left',
                    background: 'none',
                    border: 'none',
                    cursor: 'pointer'
                  }}
                >
                  Logout
                </button>
              </li>
            </div>
          </ul>
        </nav>
      </aside>
    </>
  );
}
