import AdminHome from '../pages/dashboards/AdminHome';
import CompanyDashboard from '../pages/dashboards/CompanyDashboard';
import UserHome from '../pages/dashboards/UserHome';
import DealerDashboard from '../pages/dashboards/DealerDashboard';
import ExecutiveDashboard from '../pages/dashboards/ExecutiveDashboard';

export default function RoleBasedHome() {
  const storedUser = localStorage.getItem('user');
  const user = storedUser ? JSON.parse(storedUser) : null;
  const role = user?.role;

  if (role === 'superadmin') return <AdminHome />;
  if (role === 'company_admin') return <CompanyDashboard />;
  if (role === 'user') return <UserHome />;
  if (role === 'dealer_user') return <DealerDashboard />;
  if (role === 'executive_user') return <ExecutiveDashboard />;

  return (
    <div className="p-6">
      <h1 className="text-lg font-semibold">Unknown Role</h1>
      <p className="text-slate-600">Contact admin.</p>
    </div>
  );
}
