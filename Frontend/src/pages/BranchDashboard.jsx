import { useEffect, useState } from "react";
import "../styles/BranchDashboard.css";
// import api, { BASE_URL } from "../assets/js/axiosConfig";
import MetricCard from '../components/MetricCard'

export default function BranchDashboard() {
  const storedUser = localStorage.getItem("user") ? JSON.parse(localStorage.getItem("user")) : null;
  const username = storedUser?.username || "User";

  // ========== SECTION 1: STATE STRUCTURE ==========
  // State for metrics organized by sections
  const [metrics, setMetrics] = useState({
    collections: {
      daily_cash: null,        // Renamed from 'cash'
      daily_upi: null,          // Renamed from 'upi'
      pending: null,
      monthly_total: null       // NEW: Monthly collection up to today
    },
    operations: {
      buses_active: null,
      buses_total: null,        // NEW: Total buses available
      trips_completed: null,
      trips_scheduled: null,    // NEW: Total trips scheduled
      total_passengers: null,
      routes_active: null,
      routes_total: null        // NEW: Total routes available
    },
    settlements: {
      total_transactions: null,
      verified: null,
      pending_verification: null,
      failed: null
    }
  });

  // Date filter state - defaults to today
  const [selectedDate, setSelectedDate] = useState(new Date().toISOString().split('T')[0]);
  const [loading, setLoading] = useState(true);

  // Fetch data when component mounts or date changes
  useEffect(() => {
    fetchDashboardData();
  }, [selectedDate]);

  // ========== SECTION 2: DATA FETCHING ==========
  const fetchDashboardData = async () => {
    setLoading(true);
    try {
      // TODO: Replace with actual API call
      // const response = await api.get(`${BASE_URL}/branch-dashboard/?date=${selectedDate}`);
      // setMetrics(response.data);
      
      // Simulating API latency
      await new Promise((resolve) => setTimeout(resolve, 800));
      
      // Placeholder data - this structure matches what your API should return
      // Monthly total should be calculated by backend: sum of all daily_total from month start to selectedDate
      setMetrics({
        collections: {
          daily_cash: 28150.00,      // Today's cash collection
          daily_upi: 17130.50,        // Today's UPI collection
          pending: 3200.00,
          monthly_total: 487650.75    // Sum from 1st of month to selected date
        },
        operations: {
          buses_active: 12,
          buses_total: 18,            // Total buses in fleet
          trips_completed: 48,
          trips_scheduled: 65,        // Total trips planned for the day
          total_passengers: 1240,
          routes_active: 8,
          routes_total: 12            // Total routes available
        },
        settlements: {
          total_transactions: 156,
          verified: 142,
          pending_verification: 12,
          failed: 2
        }
      });
    } catch (err) {
      console.error("Dashboard fetch error", err);
    } finally {
      setLoading(false);
    }
  };

  // ========== SECTION 3: FORMATTING FUNCTIONS ==========
  // Utility function to format currency values
  const formatCurrency = (value) => {
    if (value === null || value === undefined) return "--";
    return `₹${value.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  };

  // Utility function to format number values
  const formatNumber = (value) => {
    if (value === null || value === undefined) return "--";
    return value.toLocaleString('en-IN');
  };

  // NEW: Utility function to format active/total display
  const formatActiveTotal = (active, total) => {
    if (active === null || active === undefined || total === null || total === undefined) return "--";
    return `${formatNumber(active)}/${formatNumber(total)}`;
  };

  // ========== SECTION 4: JSX RENDERING ==========
  return (
    <div className="branch-dashboard-page">
      {/* Header Section with Date Filter */}
      <div className="dashboard-header">
        <div className="header-text">
          <h1>Branch Dashboard</h1>
          <p>Welcome back, {username}</p>
        </div>
        
        <div className="date-filter">
          <i className="fas fa-calendar-alt"></i>
          <input
            type="date"
            value={selectedDate}
            onChange={(e) => setSelectedDate(e.target.value)}
            className="date-input"
            max={new Date().toISOString().split('T')[0]}
          />
        </div>
      </div>

      {/* SECTION 1: Collection Metrics - UPDATED */}
      <div className="section-header">
        <i className="fas fa-wallet"></i>
        <h2>Collection Overview</h2>
      </div>
      <div className="metrics-grid">
        <MetricCard
          title="Total Daily Collection"
          value={loading ? "..." : formatCurrency(
            (metrics.collections.daily_cash || 0) + (metrics.collections.daily_upi || 0)
          )}
          iconClass="fas fa-rupee-sign"
          color="#3b82f6"
          loading={loading}
        />
        <MetricCard
          title="Daily Cash Collection"
          value={loading ? "..." : formatCurrency(metrics.collections.daily_cash)}
          iconClass="fas fa-money-bill-wave"
          color="#10b981"
          loading={loading}
        />
        <MetricCard
          title="Daily UPI Collection"
          value={loading ? "..." : formatCurrency(metrics.collections.daily_upi)}
          iconClass="fas fa-credit-card"
          color="#8b5cf6"
          loading={loading}
        />
        <MetricCard
          title="Monthly Collection"
          value={loading ? "..." : formatCurrency(metrics.collections.monthly_total)}
          iconClass="fas fa-chart-line"
          color="#f59e0b"
          loading={loading}
        />
      </div>

      {/* SECTION 2: Operations Metrics - UPDATED with Active/Total Format */}
      <div className="section-header">
        <i className="fas fa-bus"></i>
        <h2>Operations Overview</h2>
      </div>
      <div className="metrics-grid">
        <MetricCard
          title="Buses (Active/Total)"
          value={loading ? "..." : formatActiveTotal(metrics.operations.buses_active, metrics.operations.buses_total)}
          iconClass="fas fa-bus-alt"
          color="#14b8a6"
          loading={loading}
        />
        <MetricCard
          title="Trips (Completed/Scheduled)"
          value={loading ? "..." : formatActiveTotal(metrics.operations.trips_completed, metrics.operations.trips_scheduled)}
          iconClass="fas fa-route"
          color="#22c55e"
          loading={loading}
        />
        <MetricCard
          title="Routes (Active/Total)"
          value={loading ? "..." : formatActiveTotal(metrics.operations.routes_active, metrics.operations.routes_total)}
          iconClass="fas fa-map-marked-alt"
          color="#a855f7"
          loading={loading}
        />
        <MetricCard
          title="Total Passengers"
          value={loading ? "..." : formatNumber(metrics.operations.total_passengers)}
          iconClass="fas fa-users"
          color="#3b82f6"
          loading={loading}
        />
      </div>

      {/* SECTION 3: Settlement & Reconciliation */}
      <div className="section-header">
        <i className="fas fa-file-invoice-dollar"></i>
        <h2>Settlements & Reconciliation</h2>
      </div>
      <div className="metrics-grid">
        <MetricCard
          title="Total Transactions"
          value={loading ? "..." : formatNumber(metrics.settlements.total_transactions)}
          iconClass="fas fa-receipt"
          color="#475569"
          loading={loading}
        />
        <MetricCard
          title="Verified Settlements"
          value={loading ? "..." : formatNumber(metrics.settlements.verified)}
          iconClass="fas fa-check-circle"
          color="#22c55e"
          loading={loading}
        />
        <MetricCard
          title="Pending Verification"
          value={loading ? "..." : formatNumber(metrics.settlements.pending_verification)}
          iconClass="fas fa-exclamation-triangle"
          color="#f59e0b"
          loading={loading}
        />
        <MetricCard
          title="Failed/Disputed"
          value={loading ? "..." : formatNumber(metrics.settlements.failed)}
          iconClass="fas fa-times-circle"
          color="#ef4444"
          loading={loading}
        />
      </div>
    </div>
  );
}