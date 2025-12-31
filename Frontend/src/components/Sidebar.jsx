import '../styles/Sidebar.css';
import { useState } from 'react';
import { NavLink, useNavigate } from 'react-router-dom';
import api, { BASE_URL } from '../assets/js/axiosConfig';
// import logo from '../assets/images/logo_5.jpg';


export default function Sidebar() {
  const [isOpen, setIsOpen] = useState(false);
  const [isReportsOpen, setIsReportsOpen] = useState(false);
  const navigate = useNavigate();

  const storedUser = localStorage.getItem('user');
  const user = storedUser ? JSON.parse(storedUser) : null;
  const role = user?.role;
  const username = user?.username || user?.name || 'User'; 

  const toggleSidebar = () => setIsOpen(!isOpen);
  const closeSidebar = () => setIsOpen(false);
  const toggleReports = () => setIsReportsOpen(!isReportsOpen);

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
      <button className="sidebar-toggle" onClick={toggleSidebar} aria-label="Toggle navigation menu">
        <i className="fa-solid fa-bars sidebar-toggle__icon"></i>
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

              {/* Home */}
              <li className="sidebar__menu-item">
                <NavLink
                  to="/dashboard"
                  end
                  className={({ isActive }) => `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`}
                  onClick={closeSidebar}
                >
                  <i className="fa-solid fa-house sidebar__icon"></i>
                  <span>Home</span>
                </NavLink>
              </li>

              {/* Company Management */}
              {role === 'superadmin' && (
                <li className="sidebar__menu-item">
                  <NavLink
                    to="/dashboard/companies"
                    className={({ isActive }) => `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`}
                    onClick={closeSidebar}
                  >
                    <i className="fa-solid fa-building sidebar__icon"></i>
                    <span>Company Management</span>
                  </NavLink>
                </li>
              )}

              {/* User Management */}
              {role === 'superadmin' && (
                <li className="sidebar__menu-item">
                  <NavLink
                    to="/dashboard/users"
                    className={({ isActive }) => `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`}
                    onClick={closeSidebar}
                  >
                    <i className="fa-solid fa-users-gear sidebar__icon"></i>
                    <span>User Management</span>
                  </NavLink>
                </li>
              )}

              {/* Branch Management */}
              {role === 'branch_admin' && (
                <li className="sidebar__menu-item">
                  <NavLink
                    to="/dashboard/branches"
                    className={({ isActive }) => `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`}
                    onClick={closeSidebar}
                  >
                    <i className="fa-solid fa-code-branch sidebar__icon"></i>
                    <span>Branch Management</span>
                  </NavLink>
                </li>
              )}

              {/* Reports Dropdown */}
              {role === 'branch_admin' && (
                <li className="sidebar__menu-item">
                  <button 
                    className={`sidebar__link sidebar__dropdown-btn ${isReportsOpen ? 'sidebar__link--active-parent' : ''}`}
                    onClick={toggleReports}
                  >
                    <div className="sidebar__link-content">
                      <i className="fa-solid fa-chart-pie sidebar__icon"></i>
                      <span>Reports</span>
                    </div>
                    <i className={`fa-solid fa-chevron-down sidebar__chevron ${isReportsOpen ? 'sidebar__chevron--open' : ''}`}></i>
                  </button>
                  
                  <div className={`sidebar__submenu ${isReportsOpen ? 'sidebar__submenu--open' : ''}`}>
                    <NavLink
                      to="/dashboard/ticket-report"
                      className={({ isActive }) => `sidebar__sublink ${isActive ? 'sidebar__sublink--active' : ''}`}
                      onClick={closeSidebar}
                    >
                      <i className="fa-solid fa-file-lines sidebar__icon sidebar__icon--small"></i>
                      <span>Ticket Report</span>
                    </NavLink>
                    <NavLink
                      to="/dashboard/trip-close-report"
                      className={({ isActive }) => `sidebar__sublink ${isActive ? 'sidebar__sublink--active' : ''}`}
                      onClick={closeSidebar}
                    >
                      <i className="fa-solid fa-clipboard-check sidebar__icon sidebar__icon--small"></i>
                      <span>Trip Close Report</span>
                    </NavLink>
                  </div>
                </li>
              )}

              {/* Settlements */}
              {role === 'branch_admin' && (
                <li className="sidebar__menu-item">
                  <NavLink
                    to="/dashboard/settlements"
                    className={({ isActive }) => `sidebar__link ${isActive ? 'sidebar__link--active' : ''}`}
                    onClick={closeSidebar}
                  >
                    <i className="fa-solid fa-hand-holding-dollar sidebar__icon"></i>
                    <span>Settlements</span>
                  </NavLink>
                </li>
              )}
            </div>

            {/* Footer Section */}
            <div className="sidebar__menu-bottom">
              <div className="sidebar__profile">
                <div className="sidebar__profile-info">
                  <p className="sidebar__username">{username}</p>
                  <p className="sidebar__role">{role?.replace('_', ' ')}</p>
                </div>
              </div>

              <li className="sidebar__menu-item">
                <button onClick={handleLogout} className="sidebar__link sidebar__link--logout">
                  <i className="fa-solid fa-arrow-right-from-bracket sidebar__icon"></i>
                  <span>Logout</span>
                </button>
              </li>

              <div className="sidebar__copyright">
                © Softland India Ltd
                {/* <img src={logo} alt="Company Logo" /> */}
              </div>
            </div>
          </ul>
        </nav>
      </aside>
    </>
  );
}