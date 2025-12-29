import { useEffect, useState } from "react";
import api, { BASE_URL } from "../assets/js/axiosConfig";
import "../styles/BranchDashboard.css";

export default function BranchDashboard() {
  const storedUser = localStorage.getItem("user") ? JSON.parse(localStorage.getItem("user")) : null;
  const username = storedUser?.username || "User";

  const [metrics, setMetrics] = useState({
    daily_collection: null,
    bus_count: null,
    total_cash: null,
    total_upi: null
  });

  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetchDashboardData();
  }, []);

  const fetchDashboardData = async () => {
    setLoading(true);
    try {
      // Placeholder call
      // Replace with: const res = await api.get(`${BASE_URL}/branch-dashboard/`);
      await new Promise((resolve) => setTimeout(resolve, 600)); // mimic latency
      const res = { data: null };

      if (!res?.data) {
        // simulate empty
        setMetrics({ daily_collection: null, bus_count: null, total_cash: null, total_upi: null });
      }
    } catch (err) {
      console.error("Dashboard fetch error", err);
    } finally {
      setLoading(false);
    }
  };

  const formatValue = (value) => (value === null || value === undefined ? "--" : value);

  return (
    <div className="branch-dashboard-page">
      <div className="branch-dashboard-header">
        <h1>Welcome, {username}</h1>
        <p>Your branch dashboard overview</p>
      </div>

      <div className="branch-dashboard-grid">
        <div className="branch-card">
          <h3>Daily Collection</h3>
          <span className="value">{loading ? "..." : formatValue(metrics.daily_collection)}</span>
        </div>

        <div className="branch-card">
          <h3>Total Buses Contributed</h3>
          <span className="value">{loading ? "..." : formatValue(metrics.bus_count)}</span>
        </div>

        <div className="branch-card">
          <h3>Total Cash Payment</h3>
          <span className="value">{loading ? "..." : formatValue(metrics.total_cash)}</span>
        </div>

        <div className="branch-card">
          <h3>Total UPI Payment</h3>
          <span className="value">{loading ? "..." : formatValue(metrics.total_upi)}</span>
        </div>
      </div>
    </div>
  );
}
