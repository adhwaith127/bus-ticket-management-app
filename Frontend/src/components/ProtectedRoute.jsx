import { Navigate, Outlet, useLocation } from 'react-router-dom';
import { useEffect, useState } from 'react';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function ProtectedRoute() {
  const [isAuthenticated, setIsAuthenticated] = useState(null);
  const [userRole, setUserRole] = useState(null);
  const [loading, setLoading] = useState(true);
  const location = useLocation();

  useEffect(() => {
    // Verify authentication from backend instead of just localStorage
    verifyAuthFromBackend();
  }, []);

  const verifyAuthFromBackend = async () => {
    try {
      const response = await api.get(`${BASE_URL}/verify-auth/`);
      if (response.data.authenticated) {
        setIsAuthenticated(true);
        setUserRole(response.data.user.role);
        // Update localStorage with verified user data
        localStorage.setItem('user', JSON.stringify(response.data.user));
      } else {
        setIsAuthenticated(false);
        localStorage.removeItem('user');
      }
    } catch (error) {
      console.error('Auth verification failed:', error);
      setIsAuthenticated(false);
      localStorage.removeItem('user');
    } finally {
      setLoading(false);
    }
  };

  // Role-based access control with redirect
  useEffect(() => {
    if (!loading && isAuthenticated && userRole) {
      const path = location.pathname;
      
      // Check role-based restrictions
      if (path.includes('/companies') || path.includes('/users')) {
        if (userRole !== 'superadmin') {
          window.alert('Access Denied: You do not have permission to view this page');
          window.location.href = '/dashboard';
        }
      } else if (path.includes('/ticket-report')) {
        // if (userRole !== 'branch_admin' && userRole !== 'superadmin') {
        if (userRole !== 'branch_admin') {
          window.alert('Access Denied: You do not have permission to view this page');
          window.location.href = '/dashboard';
        }
      }
    }
  }, [loading, isAuthenticated, userRole, location.pathname]);

  if (loading) {
    return (
      <div style={{ 
        display: 'flex', 
        justifyContent: 'center', 
        alignItems: 'center', 
        height: '100vh',
        fontSize: '18px'
      }}>
        Loading...
      </div>
    );
  }

  if (!isAuthenticated) {
    return <Navigate to="/login" replace />;
  }

  return <Outlet />;
}