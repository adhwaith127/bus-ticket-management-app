import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function TicketReport() {
  // ===== STATE =====
  const [transactions, setTransactions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  
  const [filters, setFilters] = useState({
    startDate: '',
    endDate: '',
    deviceId: '',
    companyCode: '',
    ticketStatus: ''
  });
  
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(10);

  // ===== FETCH =====
  useEffect(() => {
    fetchTransactions();
  }, []);

  const fetchTransactions = async () => {
    try {
      setLoading(true);
      const response = await api.get(`${BASE_URL}/get_all_transaction_data`);
      if (response.data.message === 'success') {
        setTransactions(response.data.data);
      } else setError('Failed to fetch transactions');
    } catch (err) {
      if (err.response) setError(`Server Error: ${err.response.status} - ${err.response.data?.message}`);
      else if (err.request) setError('No response from server.');
      else setError('Error: ' + err.message);
    } finally {
      setLoading(false);
    }
  };

  // ===== FILTER =====
  const getFilteredData = () =>
    transactions.filter(t => {
      if (filters.startDate && t.ticket_date)
        if (new Date(t.ticket_date) < new Date(filters.startDate)) return false;

      if (filters.endDate && t.ticket_date)
        if (new Date(t.ticket_date) > new Date(filters.endDate)) return false;

      if (filters.deviceId && t.device_id)
        if (!t.device_id.toLowerCase().includes(filters.deviceId.toLowerCase())) return false;

      if (filters.companyCode && t.company_code)
        if (!t.company_code.toLowerCase().includes(filters.companyCode.toLowerCase())) return false;

      if (filters.ticketStatus && t.ticket_status)
        if (!t.ticket_status.toLowerCase().includes(filters.ticketStatus.toLowerCase())) return false;

      return true;
    });

  const filteredData = getFilteredData();
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);
  const currentData = filteredData.slice((currentPage - 1) * itemsPerPage, currentPage * itemsPerPage);

  const changePage = (p) => setCurrentPage(p);
  const clearFilters = () => {
    setFilters({ startDate: '', endDate: '', deviceId: '', companyCode: '', ticketStatus: '' });
    setCurrentPage(1);
  };

  // ===== EXPORT =====
  const exportToExcel = () => {
    const exportData = filteredData.map(t => ({
      ID: t.id,
      RequestType: t.request_type,
      DeviceID: t.device_id,
      TripNumber: t.trip_number,
      TicketNumber: t.ticket_number,
      TicketDate: t.ticket_date,
      TicketTime: t.ticket_time,
      FromStage: t.from_stage,
      ToStage: t.to_stage,
      FullCount: t.full_count,
      HalfCount: t.half_count,
      ST: t.st_count,
      Phy: t.phy_count,
      Lugg: t.lugg_count,
      TicketAmount: t.ticket_amount,
      LuggAmount: t.lugg_amount,
      TicketType: t.ticket_type,
      Status: t.ticket_status,
      Reference: t.reference_number,
      TxnID: t.transaction_id,
      Company: t.company_code,
      Created: t.created_at,
    }));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(exportData), "Transactions");
    XLSX.writeFile(wb, `ticket_report_${new Date().toISOString().slice(0,10)}.xlsx`);
  };

  // ===== LOADING / ERROR =====
  if (loading)
    return (
      <div className="flex items-center justify-center h-screen text-slate-500 text-lg">
        Loading transaction data...
      </div>
    );

  if (error)
    return (
      <div className="flex items-center justify-center h-screen text-red-600 font-medium">
        {error}
      </div>
    );

  // ===== UI =====
  return (
    <div className="p-6 md:p-10 min-h-screen bg-slate-50 animate-fade-in">
      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
        <div>
          <h1 className="text-3xl font-bold text-slate-800 tracking-tight">Ticket Transaction Reports</h1>
          <p className="text-slate-500 mt-1">View and manage daily ticket transactions</p>
        </div>
        <button
          onClick={exportToExcel}
          className="bg-slate-800 hover:bg-slate-700 text-white px-5 py-2.5 rounded-xl shadow-lg transition"
        >
          Download Report
        </button>
      </div>

      {/* Filters */}
      <div className="bg-white border border-slate-200 rounded-xl shadow-sm p-4 md:p-6 mb-6">
        <div className="grid grid-cols-1 md:grid-cols-5 gap-4">
          {[
            { label: "Start Date", type: "date", key: "startDate" },
            { label: "End Date", type: "date", key: "endDate" },
            { label: "Device ID", placeholder: "Search device...", key: "deviceId" },
            { label: "Company Code", placeholder: "Search company...", key: "companyCode" },
            { label: "Ticket Status", placeholder: "Search status...", key: "ticketStatus" }
          ].map((f) => (
            <div key={f.key} className="flex flex-col">
              <label className="text-xs font-medium text-slate-500">{f.label}</label>
              <input
                type={f.type || "text"}
                placeholder={f.placeholder}
                value={filters[f.key]}
                onChange={(e) => setFilters({ ...filters, [f.key]: e.target.value })}
                className="border border-slate-300 rounded-lg px-3 py-2 text-sm focus:ring-2 focus:ring-slate-500"
              />
            </div>
          ))}
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

      {/* Summary */}
      <div className="text-sm text-slate-500 mb-3">
        Showing {Math.min(currentPage * itemsPerPage, filteredData.length)} of {filteredData.length} transactions
      </div>

      {/* Table */}
      <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-sm text-left border-collapse">
            <thead className="bg-slate-50 border-b border-slate-200 text-slate-600 text-xs uppercase tracking-wide">
              <tr>
                {[
                  "Ticket #","Device","Trip","Date","Time","From","To",
                  "Type","Status","Full","Half","ST","Phy","Lugg",
                  "Ticket Amt","Lugg Amt","Company","Reference","Txn ID"
                ].map((h) => (
                  <th key={h} className="px-4 py-3 font-semibold">{h}</th>
                ))}
              </tr>
            </thead>

            <tbody className="divide-y divide-slate-100">
              {currentData.length ? (
                currentData.map((t) => (
                  <tr key={t.id} className="hover:bg-slate-50 transition">
                    <td className="px-4 py-3">{t.ticket_number || "-"}</td>
                    <td className="px-4 py-3">{t.device_id}</td>
                    <td className="px-4 py-3">{t.trip_number}</td>
                    <td className="px-4 py-3">{t.ticket_date}</td>
                    <td className="px-4 py-3">{t.ticket_time}</td>
                    <td className="px-4 py-3">{t.from_stage}</td>
                    <td className="px-4 py-3">{t.to_stage}</td>
                    <td className="px-4 py-3">{t.ticket_type}</td>
                    <td className="px-4 py-3">{t.ticket_status}</td>
                    <td className="px-4 py-3">{t.full_count}</td>
                    <td className="px-4 py-3">{t.half_count}</td>
                    <td className="px-4 py-3">{t.st_count}</td>
                    <td className="px-4 py-3">{t.phy_count}</td>
                    <td className="px-4 py-3">{t.lugg_count}</td>
                    <td className="px-4 py-3">₹{t.ticket_amount}</td>
                    <td className="px-4 py-3">₹{t.lugg_amount}</td>
                    <td className="px-4 py-3">{t.company_code}</td>
                    <td className="px-4 py-3">{t.reference_number}</td>
                    <td className="px-4 py-3">{t.transaction_id}</td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={20} className="px-4 py-6 text-center text-slate-500">
                    No trip data found
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
    </div>
  );
}
