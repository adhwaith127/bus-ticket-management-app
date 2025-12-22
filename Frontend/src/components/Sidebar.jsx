import '../styles/Sidebar.css';
import { useState } from 'react';
import { NavLink, useNavigate } from 'react-router-dom'; // Added useNavigate
import api, { BASE_URL } from '../assets/js/axiosConfig'; // Added API imports

export default function Sidebar() {
  const [isOpen, setIsOpen] = useState(false);
  const navigate = useNavigate(); // Initialize navigate

  const toggleSidebar = () => {
    setIsOpen(!isOpen);
  };

  const closeSidebar = () => {
    setIsOpen(false);
  };

  // Logout Logic
  const handleLogout = async () => {
    try {
      await api.post(`${BASE_URL}/logout/`);
      localStorage.removeItem('user');
      closeSidebar();
      navigate('/login');
    } catch (err) {
      console.error('Logout error:', err);
      // Still clear local storage and redirect even if the server call fails
      localStorage.removeItem('user');
      closeSidebar();
      navigate('/login');
    }
  };

  return (
    <>
      {/* Mobile Toggle Button */}
      <button
        className="sidebar-toggle"
        onClick={toggleSidebar}
        aria-label="Toggle navigation menu"
      >
        <span className="sidebar-toggle__icon">☰</span>
      </button>

      {/* Mobile Overlay */}
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
              <li className="sidebar__menu-item">
                <NavLink
                  to="/dashboard"
                  end
                  className={({ isActive }) => 
                    `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`
                  }
                  onClick={closeSidebar}
                >
                  Home
                </NavLink>
              </li>
              
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
            </div>

            {/* Added Logout Section at the Bottom */}
            <div className="sidebar__menu-bottom">
              <li className="sidebar__menu-item">
                <button
                  onClick={handleLogout}
                  className="sidebar__link sidebar__link--logout"
                  style={{ width: '100%', textAlign: 'left', background: 'none', border: 'none', cursor: 'pointer' }}
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