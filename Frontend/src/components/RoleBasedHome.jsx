import AdminHome from '../pages/AdminHome';
import BranchDashboard from '../pages/BranchDashboard';
import UserHome from '../pages/UserHome';

export default function RoleBasedHome() {
  // Get user from localStorage
  const storedUser = localStorage.getItem('user');
  const user = storedUser ? JSON.parse(storedUser) : null;
  const role = user?.role;

  // Render appropriate home based on role
  if (role === 'superadmin') {
    return <AdminHome />;
  } else if (role === 'branch_admin') {
    return <BranchDashboard />;
  } else if (role === 'user') {
    return <UserHome />
  }

  // Fallback (shouldn't reach here due to ProtectedRoute)
  return (
    <div style={{ padding: '20px' }}>
      <h2>Unknown Role</h2>
      <p>Please contact administrator.</p>
    </div>
  );
}