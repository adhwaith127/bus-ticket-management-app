import { Outlet } from 'react-router-dom';
import Sidebar from '../components/Sidebar';
import '../styles/Dashboard.css';

const Dashboard = () => {
  return (
    <div className="dashboard">
      <Sidebar />
      <main className="dashboard__content">
        <Outlet />
      </main>
    </div>
  );
};

export default Dashboard;