import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function TicketReport() {
  // ===== SECTION 1: STATE MANAGEMENT =====
  const [transactions, setTransactions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isRefreshing, setIsRefreshing] = useState(false);
  const [error, setError] = useState(null);
  
  const [filters, setFilters] = useState({
    startDate: '',
    endDate: '',
    deviceId: 'ALL',
    branchCode: 'ALL',
    paymentMode: 'ALL'
  });
  
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(10);
  
  // Modal states
  const [showModal, setShowModal] = useState(false);
  const [selectedTransaction, setSelectedTransaction] = useState(null);

  // ===== SECTION 2: DATE & API LOGIC =====
  const getTodayDate = () => {
    const today = new Date();
    return today.toISOString().split('T')[0]; // YYYY-MM-DD format
  };

  // Initialize with today's date
  useEffect(() => {
    const today = getTodayDate();
    setFilters(prev => ({
      ...prev,
      startDate: today,
      endDate: today
    }));
  }, []);

  // Trigger API call when dates change
  useEffect(() => {
    if (filters.startDate && filters.endDate) {
      fetchTransactions(filters.startDate, filters.endDate);
    }
  }, [filters.startDate, filters.endDate]);

  const fetchTransactions = async (startDate, endDate) => {
    try {
      const isInitialLoad = transactions.length === 0;
      if (isInitialLoad) {
        setLoading(true);
      } else {
        setIsRefreshing(true);
      }
      
      const response = await api.get(
        `${BASE_URL}/get_all_transaction_data?from_date=${startDate}&to_date=${endDate}`
      );
      
      if (response.data.message === 'success') {
        setTransactions(response.data.data);
        setError(null);
      } else {
        setError('Failed to fetch transactions');
      }
    } catch (err) {
      if (err.response) {
        setError(`Server Error: ${err.response.status} - ${err.response.data?.message}`);
      } else if (err.request) {
        setError('No response from server.');
      } else {
        setError('Error: ' + err.message);
      }
    } finally {
      setLoading(false);
      setIsRefreshing(false);
    }
  };

  // Placeholder for future branch name fetch
  // const fetchBranches = async () => {
  //   try {
  //     const response = await api.get(`${BASE_URL}/branches/`);
  //     // Map branch codes to names: { 'BR001': 'Branch Name', ... }
  //   } catch (err) {
  //     console.error('Failed to fetch branches:', err);
  //   }
  // };

  // ===== SECTION 3: DYNAMIC DROPDOWN OPTIONS =====
  const getUniqueOptions = () => {
    const deviceIds = [...new Set(transactions.map(t => t.device_id).filter(Boolean))].sort();
    const branchCodes = [...new Set(transactions.map(t => t.branch_code).filter(Boolean))].sort();
    const paymentModes = [...new Set(transactions.map(t => t.payment_mode_display).filter(Boolean))].sort();
    
    return { deviceIds, branchCodes, paymentModes };
  };

  const { deviceIds, branchCodes, paymentModes } = getUniqueOptions();

  // ===== SECTION 4: FILTER LOGIC =====
  const getFilteredData = () =>
    transactions.filter(t => {
      // Device ID filter
      if (filters.deviceId && filters.deviceId !== 'ALL') {
        if (t.device_id !== filters.deviceId) return false;
      }

      // Branch Code filter
      if (filters.branchCode && filters.branchCode !== 'ALL') {
        if (t.branch_code !== filters.branchCode) return false;
      }

      // Payment Mode filter
      if (filters.paymentMode && filters.paymentMode !== 'ALL') {
        if (t.payment_mode_display !== filters.paymentMode) return false;
      }

      return true;
    });

  const filteredData = getFilteredData();

  // ===== SECTION 6: PAGINATION =====
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);
  const currentData = filteredData.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  // ===== SECTION 5: SUMMARY CALCULATIONS =====
  const calculateSummary = (data) => {
    const totalTickets = data.reduce((sum, t) => sum + (t.total_tickets || 0), 0);
    const totalAmount = data.reduce((sum, t) => sum + parseFloat(t.ticket_amount || 0), 0);
    const upiCount = data.filter(t => t.payment_mode_display === 'UPI').length;
    const cashCount = data.filter(t => t.payment_mode_display === 'Cash').length;
    
    return { totalTickets, totalAmount, upiCount, cashCount };
  };

  // Calculate summary based on CURRENT PAGE data (what's visible in table)
  const summary = calculateSummary(currentData);

  const changePage = (p) => setCurrentPage(p);

  const clearFilters = () => {
    const today = getTodayDate();
    setFilters({
      startDate: today,
      endDate: today,
      deviceId: 'ALL',
      branchCode: 'ALL',
      paymentMode: 'ALL'
    });
    setCurrentPage(1);
  };

  // ===== SECTION 7: EXPORT LOGIC =====
  const exportToExcel = () => {
    const exportData = filteredData.map(t => ({
      DeviceID: t.device_id,
      TripNumber: t.trip_number,
      TicketNumber: t.ticket_number,
      Date: t.formatted_ticket_date,
      Time: t.ticket_time,
      BranchCode: t.branch_code,
      TotalTickets: t.total_tickets,
      TicketAmount: t.ticket_amount,
      PaymentMode: t.payment_mode_display,
      TicketType: t.ticket_type_display,
      FromStage: t.from_stage,
      ToStage: t.to_stage,
      FullCount: t.full_count,
      HalfCount: t.half_count,
      STCount: t.st_count,
      PhyCount: t.phy_count,
      LuggCount: t.lugg_count,
      LuggAmount: t.lugg_amount,
      AdjustAmount: t.adjust_amount,
      WarrantAmount: t.warrant_amount,
      RefundAmount: t.refund_amount,
      LadiesCount: t.ladies_count,
      SeniorCount: t.senior_count,
      TransactionID: t.transaction_id,
      Reference: t.reference_number,
      PassID: t.pass_id,
      RefundStatus: t.refund_status,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exportData), "Transactions");
    XLSX.writeFile(wb, `ticket_report_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  // ===== SECTION 8: MODAL HANDLERS =====
  const openModal = (transaction) => {
    setSelectedTransaction(transaction);
    setShowModal(true);
  };

  const closeModal = () => {
    setShowModal(false);
    setSelectedTransaction(null);
  };

  // ===== LOADING / ERROR STATES =====
  if (loading) {
    return (
      <div className="flex items-center justify-center h-screen text-slate-500 text-lg">
        Loading transaction data...
      </div>
    );
  }

  if (error) {
    return (
      <div className="flex items-center justify-center h-screen text-red-600 font-medium">
        {error}
      </div>
    );
  }

  // ===== SECTION 9: UI RENDER =====
  return (
    <div className="p-6 md:p-10 min-h-screen bg-slate-50">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-6 gap-4">
        <div>
          <h1 className="text-3xl font-bold text-slate-800">Ticket Transaction Reports</h1>
          <p className="text-slate-500 mt-1">View and manage daily ticket transactions</p>
        </div>
        <button
          onClick={exportToExcel}
          className="bg-slate-800 hover:bg-slate-700 text-white px-5 py-2.5 rounded-xl shadow-lg transition"
        >
          Download Report
        </button>
      </div>

      {/* Summary Cards */}
      <div className="grid grid-cols-1 md:grid-cols-4 gap-4 mb-6">
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm">
          <div className="text-slate-500 text-sm font-medium">Total Tickets</div>
          <div className="text-2xl font-bold text-slate-800 mt-1">{summary.totalTickets}</div>
        </div>
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm">
          <div className="text-slate-500 text-sm font-medium">Total Amount</div>
          <div className="text-2xl font-bold text-slate-800 mt-1">₹{summary.totalAmount.toFixed(2)}</div>
        </div>
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm">
          <div className="text-slate-500 text-sm font-medium">UPI Payments</div>
          <div className="text-2xl font-bold text-slate-800 mt-1">{summary.upiCount}</div>
        </div>
        <div className="bg-white rounded-xl p-5 border border-slate-200 shadow-sm">
          <div className="text-slate-500 text-sm font-medium">Cash Payments</div>
          <div className="text-2xl font-bold text-slate-800 mt-1">{summary.cashCount}</div>
        </div>
      </div>

      {/* Filters */}
      <div className="bg-white border border-slate-200 rounded-xl shadow-sm p-4 md:p-6 mb-6">
        <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
          {/* Start Date */}
          <div className="flex flex-col">
            <label className="text-xs font-medium text-slate-500 mb-1">Start Date</label>
            <input
              type="date"
              value={filters.startDate}
              onChange={(e) => setFilters({ ...filters, startDate: e.target.value })}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500"
            />
          </div>

          {/* End Date */}
          <div className="flex flex-col">
            <label className="text-xs font-medium text-slate-500 mb-1">End Date</label>
            <input
              type="date"
              value={filters.endDate}
              onChange={(e) => setFilters({ ...filters, endDate: e.target.value })}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500"
            />
          </div>

          {/* Device ID Dropdown */}
          <div className="flex flex-col">
            <label className="text-xs font-medium text-slate-500 mb-1">Device ID</label>
            <select
              value={filters.deviceId}
              onChange={(e) => setFilters({ ...filters, deviceId: e.target.value })}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500"
            >
              <option value="ALL">ALL</option>
              {deviceIds.map(id => (
                <option key={id} value={id}>{id}</option>
              ))}
            </select>
          </div>

          {/* Branch Code Dropdown */}
          <div className="flex flex-col">
            <label className="text-xs font-medium text-slate-500 mb-1">Branch Code</label>
            <select
              value={filters.branchCode}
              onChange={(e) => setFilters({ ...filters, branchCode: e.target.value })}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500"
            >
              <option value="ALL">ALL</option>
              {branchCodes.map(code => (
                <option key={code} value={code}>{code}</option>
              ))}
            </select>
          </div>

          {/* Payment Mode Dropdown */}
          <div className="flex flex-col">
            <label className="text-xs font-medium text-slate-500 mb-1">Payment Mode</label>
            <select
              value={filters.paymentMode}
              onChange={(e) => setFilters({ ...filters, paymentMode: e.target.value })}
              className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500"
            >
              <option value="ALL">ALL</option>
              {paymentModes.map(mode => (
                <option key={mode} value={mode}>{mode}</option>
              ))}
            </select>
          </div>
        </div>

        <div className="flex justify-end mt-4">
          <button
            onClick={clearFilters}
            className="border border-slate-300 px-4 py-2 rounded-lg text-sm text-slate-700 hover:bg-slate-100"
          >
            Clear Filters
          </button>
        </div>
      </div>

      {/* Summary Info */}
      <div className="text-sm text-slate-500 mb-3">
        Showing {Math.min(currentPage * itemsPerPage, filteredData.length)} of {filteredData.length} transactions
      </div>

      {/* Table */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm text-left border-collapse">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 text-xs uppercase tracking-wide">
              <tr>
                <th className="px-4 py-3 font-semibold">Device ID</th>
                <th className="px-4 py-3 font-semibold">Trip No</th>
                <th className="px-4 py-3 font-semibold">Ticket No</th>
                <th className="px-4 py-3 font-semibold">Date</th>
                <th className="px-4 py-3 font-semibold">Time</th>
                <th className="px-4 py-3 font-semibold">Branch</th>
                <th className="px-4 py-3 font-semibold">Total Tickets</th>
                <th className="px-4 py-3 font-semibold">Amount</th>
                <th className="px-4 py-3 font-semibold">Payment</th>
                <th className="px-4 py-3 font-semibold">Info</th>
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100">
              {isRefreshing ? (
                <tr>
                  <td colSpan={10} className="px-4 py-8 text-center text-slate-500">
                    <div className="flex items-center justify-center gap-2">
                      <div className="animate-spin rounded-full h-5 w-5 border-b-2 border-slate-500"></div>
                      Loading...
                    </div>
                  </td>
                </tr>
              ) : currentData.length ? (
                currentData.map((t) => (
                  <tr key={t.id} className="hover:bg-slate-50 transition">
                    <td className="px-4 py-3">{t.device_id}</td>
                    <td className="px-4 py-3">{t.trip_number}</td>
                    <td className="px-4 py-3">{t.ticket_number || "-"}</td>
                    <td className="px-4 py-3">{t.formatted_ticket_date}</td>
                    <td className="px-4 py-3">{t.ticket_time}</td>
                    <td className="px-4 py-3">{t.branch_code || "-"}</td>
                    <td className="px-4 py-3">{t.total_tickets}</td>
                    <td className="px-4 py-3">₹{t.ticket_amount}</td>
                    <td className="px-4 py-3">
                      <span className={`px-2 py-1 rounded text-xs ${
                        t.payment_mode_display === 'UPI' 
                          ? 'bg-blue-100 text-blue-700' 
                          : 'bg-green-100 text-green-700'
                      }`}>
                        {t.payment_mode_display}
                      </span>
                    </td>
                    <td className="px-4 py-3">
                      <button
                        onClick={() => openModal(t)}
                        className="text-slate-600 hover:text-slate-900 transition"
                        title="View Details"
                      >
                        <svg className="w-5 h-5" fill="currentColor" viewBox="0 0 20 20">
                          <path d="M18 10a8 8 0 11-16 0 8 8 0 0116 0zm-7-4a1 1 0 11-2 0 1 1 0 012 0zM9 9a1 1 0 000 2v3a1 1 0 001 1h1a1 1 0 100-2v-3a1 1 0 00-1-1H9z" />
                        </svg>
                      </button>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={10} className="px-4 py-6 text-center text-slate-500">
                    No transaction data found
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* Pagination */}
      {totalPages > 1 && (
        <div className="flex items-center justify-center space-x-2 mt-6">
          <button
            onClick={() => changePage(currentPage - 1)}
            disabled={currentPage === 1}
            className="px-3 py-1.5 rounded-lg border disabled:opacity-30"
          >
            Prev
          </button>

          {[...Array(totalPages)].map((_, i) => {
            const n = i + 1;
            return (
              <button
                key={n}
                onClick={() => changePage(n)}
                className={`px-3 py-1.5 rounded-lg border ${
                  currentPage === n ? "bg-slate-800 text-white border-slate-800" : ""
                }`}
              >
                {n}
              </button>
            );
          })}

          <button
            onClick={() => changePage(currentPage + 1)}
            disabled={currentPage === totalPages}
            className="px-3 py-1.5 rounded-lg border disabled:opacity-30"
          >
            Next
          </button>
        </div>
      )}

      {/* Modal */}
      {showModal && selectedTransaction && (
        <div className="fixed inset-0 bg-black bg-opacity-30 flex items-center justify-center z-50 p-4">
          <div className="bg-white rounded-xl shadow-2xl max-w-2xl w-full max-h-[80vh] overflow-y-auto">
            {/* Modal Header */}
            <div className="sticky top-0 bg-white border-b border-slate-200 px-6 py-4 flex items-center justify-between">
              <h2 className="text-xl font-bold text-slate-800">Transaction Details</h2>
              <button
                onClick={closeModal}
                className="text-slate-400 hover:text-slate-600 transition"
              >
                <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M6 18L18 6M6 6l12 12" />
                </svg>
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 space-y-4">
              {/* Ticket Type */}
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <div className="text-xs text-slate-500 font-medium">Ticket Type</div>
                  <div className="text-sm text-slate-800 mt-1">{selectedTransaction.ticket_type_display}</div>
                </div>
                <div>
                  <div className="text-xs text-slate-500 font-medium">Request Type</div>
                  <div className="text-sm text-slate-800 mt-1">{selectedTransaction.request_type || "-"}</div>
                </div>
              </div>

              {/* Stages */}
              <div className="grid grid-cols-2 gap-4">
                <div>
                  <div className="text-xs text-slate-500 font-medium">From Stage</div>
                  <div className="text-sm text-slate-800 mt-1">{selectedTransaction.from_stage}</div>
                </div>
                <div>
                  <div className="text-xs text-slate-500 font-medium">To Stage</div>
                  <div className="text-sm text-slate-800 mt-1">{selectedTransaction.to_stage}</div>
                </div>
              </div>

              {/* Passenger Counts */}
              <div className="border-t border-slate-200 pt-4">
                <h3 className="font-semibold text-slate-700 mb-3">Passenger Counts</h3>
                <div className="grid grid-cols-3 gap-4">
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Full</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.full_count}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Half</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.half_count}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Student</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.st_count}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Physical</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.phy_count}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Luggage</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.lugg_count}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Ladies</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.ladies_count}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Senior</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.senior_count}</div>
                  </div>
                </div>
              </div>

              {/* Amounts */}
              <div className="border-t border-slate-200 pt-4">
                <h3 className="font-semibold text-slate-700 mb-3">Amount Details</h3>
                <div className="grid grid-cols-2 gap-4">
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Luggage Amount</div>
                    <div className="text-sm text-slate-800 mt-1">₹{selectedTransaction.lugg_amount}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Adjust Amount</div>
                    <div className="text-sm text-slate-800 mt-1">₹{selectedTransaction.adjust_amount}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Warrant Amount</div>
                    <div className="text-sm text-slate-800 mt-1">₹{selectedTransaction.warrant_amount}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Refund Amount</div>
                    <div className="text-sm text-slate-800 mt-1">₹{selectedTransaction.refund_amount}</div>
                  </div>
                </div>
              </div>

              {/* References */}
              <div className="border-t border-slate-200 pt-4">
                <h3 className="font-semibold text-slate-700 mb-3">Reference Information</h3>
                <div className="space-y-3">
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Transaction ID</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.transaction_id || "-"}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Reference Number</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.reference_number || "-"}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Pass ID</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.pass_id || "-"}</div>
                  </div>
                  <div>
                    <div className="text-xs text-slate-500 font-medium">Refund Status</div>
                    <div className="text-sm text-slate-800 mt-1">{selectedTransaction.refund_status || "-"}</div>
                  </div>
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="border-t border-slate-200 px-6 py-4 flex justify-end">
              <button
                onClick={closeModal}
                className="bg-slate-800 hover:bg-slate-700 text-white px-4 py-2 rounded-lg transition"
              >
                Close
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}