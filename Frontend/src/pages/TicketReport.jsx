import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import '../styles/TicketReport.css';
// UPDATED: Import api from axiosConfig for authentication
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function TicketReport() {
  // ==================== STATE MANAGEMENT SECTION ====================
  const [transactions, setTransactions] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  
  // Filter states
  const [filters, setFilters] = useState({
    startDate: '',
    endDate: '',
    deviceId: '',
    companyCode: '',
    ticketStatus: ''
  });
  
  // Pagination states
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(10);

  // ==================== DATA FETCHING SECTION ====================
  useEffect(() => {
    fetchTransactions();
  }, []);

  const fetchTransactions = async () => {
    try {
      setLoading(true);
      // UPDATED: Use api instance with authentication cookies
      const response = await api.get(`${BASE_URL}/get_all_transaction_data`);
      
      if (response.data.message === 'success') {
        setTransactions(response.data.data);
      } else {
        setError('Failed to fetch transactions');
      }
    } catch (err) {
      if (err.response) {
        setError(`Server Error: ${err.response.status} - ${err.response.data.message || 'Unknown error'}`);
      } else if (err.request) {
        setError('No response from server. Please check your connection.');
      } else {
        setError('Error setting up request: ' + err.message);
      }
    } finally {
      setLoading(false);
    }
  };

  // ==================== FILTER LOGIC SECTION ====================
  const getFilteredData = () => {
    return transactions.filter(transaction => {
      // Date range filter
      if (filters.startDate && transaction.ticket_date) {
        if (new Date(transaction.ticket_date) < new Date(filters.startDate)) {
          return false;
        }
      }
      if (filters.endDate && transaction.ticket_date) {
        if (new Date(transaction.ticket_date) > new Date(filters.endDate)) {
          return false;
        }
      }
      
      // Device ID filter
      if (filters.deviceId && transaction.device_id) {
        if (!transaction.device_id.toLowerCase().includes(filters.deviceId.toLowerCase())) {
          return false;
        }
      }
      
      // Company code filter
      if (filters.companyCode && transaction.company_code) {
        if (!transaction.company_code.toLowerCase().includes(filters.companyCode.toLowerCase())) {
          return false;
        }
      }
      
      // Ticket status filter
      if (filters.ticketStatus && transaction.ticket_status) {
        if (!transaction.ticket_status.toLowerCase().includes(filters.ticketStatus.toLowerCase())) {
          return false;
        }
      }
      
      return true;
    });
  };

  // ==================== PAGINATION LOGIC SECTION ====================
  const filteredData = getFilteredData();
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentData = filteredData.slice(startIndex, endIndex);

  const handlePageChange = (pageNumber) => {
    setCurrentPage(pageNumber);
  };

  const handleFilterChange = (filterName, value) => {
    setFilters(prev => ({ ...prev, [filterName]: value }));
    setCurrentPage(1); // Reset to first page when filters change
  };

  const clearFilters = () => {
    setFilters({
      startDate: '',
      endDate: '',
      deviceId: '',
      companyCode: '',
      ticketStatus: ''
    });
    setCurrentPage(1);
  };

  // ==================== EXCEL EXPORT SECTION ====================
  const exportToExcel = () => {
    // Prepare data for export (using filtered data)
    const exportData = filteredData.map(transaction => ({
      'ID': transaction.id,
      'Request Type': transaction.request_type || '',
      'Device ID': transaction.device_id || '',
      'Trip Number': transaction.trip_number || '',
      'Ticket Number': transaction.ticket_number || '',
      'Ticket Date': transaction.ticket_date || '',
      'Ticket Time': transaction.ticket_time || '',
      'From Stage': transaction.from_stage || 0,
      'To Stage': transaction.to_stage || 0,
      'Full Count': transaction.full_count || 0,
      'Half Count': transaction.half_count || 0,
      'ST Count': transaction.st_count || 0,
      'Phy Count': transaction.phy_count || 0,
      'Lugg Count': transaction.lugg_count || 0,
      'Ticket Amount': transaction.ticket_amount || 0,
      'Lugg Amount': transaction.lugg_amount || 0,
      'Ticket Type': transaction.ticket_type || '',
      'Ticket Status': transaction.ticket_status || '',
      'Reference Number': transaction.reference_number || '',
      'Transaction ID': transaction.transaction_id || '',
      'Company Code': transaction.company_code || '',
      'Created At': transaction.created_at || ''
    }));

    // Create worksheet
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    
    // Create workbook
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Transactions');
    
    // Download file
    const fileName = `ticket_report_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  // ==================== UI RENDERING SECTION ====================
  if (loading) {
    return (
      <div className="ticketReport">
        <div className="ticketReport__loading">Loading transaction data...</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="ticketReport">
        <div className="ticketReport__error">{error}</div>
      </div>
    );
  }

  return (
    <div className="ticketReport">
      {/* Header Section */}
      <div className="ticketReport__header">
        <div>
          <h1 className="ticketReport__title">Ticket Transaction Reports</h1>
          <p className="ticketReport__subtitle">View and manage daily ticket transactions</p>
        </div>
        <button className="ticketReport__button ticketReport__button--primary"onClick={exportToExcel}>
          Download Report
        </button>
      </div>

      {/* Filters Section */}
      <div className="ticketReport__filters">
        <div className="ticketReport__filterGroup">
          <label className="ticketReport__filterLabel">Start Date:</label>
          <input
            type="date"
            className="ticketReport__filterInput"
            value={filters.startDate}
            onChange={(e) => handleFilterChange('startDate', e.target.value)}
          />
        </div>

        <div className="ticketReport__filterGroup">
          <label className="ticketReport__filterLabel">End Date:</label>
          <input
            type="date"
            className="ticketReport__filterInput"
            value={filters.endDate}
            onChange={(e) => handleFilterChange('endDate', e.target.value)}
          />
        </div>

        <div className="ticketReport__filterGroup">
          <label className="ticketReport__filterLabel">Device ID:</label>
          <input
            type="text"
            className="ticketReport__filterInput"
            placeholder="Search device..."
            value={filters.deviceId}
            onChange={(e) => handleFilterChange('deviceId', e.target.value)}
          />
        </div>

        <div className="ticketReport__filterGroup">
          <label className="ticketReport__filterLabel">Company Code:</label>
          <input
            type="text"
            className="ticketReport__filterInput"
            placeholder="Search company..."
            value={filters.companyCode}
            onChange={(e) => handleFilterChange('companyCode', e.target.value)}
          />
        </div>

        <div className="ticketReport__filterGroup">
          <label className="ticketReport__filterLabel">Ticket Status:</label>
          <input
            type="text"
            className="ticketReport__filterInput"
            placeholder="Search status..."
            value={filters.ticketStatus}
            onChange={(e) => handleFilterChange('ticketStatus', e.target.value)}
          />
        </div>

        <button 
          className="ticketReport__button ticketReport__button--secondary"
          onClick={clearFilters}
        >
          Clear Filters
        </button>
      </div>

      {/* Summary Section */}
      <div className="ticketReport__summary">
        <span className="ticketReport__summaryText">
          Showing {Math.min(endIndex, filteredData.length)} of {filteredData.length} transactions
        </span>
      </div>

      {/* Table Section */}
      <div className="ticketReport__tableWrapper">
        <table className="ticketReport__table">
          <thead className="ticketReport__thead">
            <tr className="ticketReport__row ticketReport__row--header">
              <th className="ticketReport__th">Ticket Number</th>
              <th className="ticketReport__th">Device ID</th>
              <th className="ticketReport__th">Trip Number</th>
              <th className="ticketReport__th">Date</th>
              <th className="ticketReport__th">Time</th>
              <th className="ticketReport__th">From Stage</th>
              <th className="ticketReport__th">To Stage</th>
              <th className="ticketReport__th">Ticket Type</th>
              <th className="ticketReport__th">Status</th>
              <th className="ticketReport__th">Full</th>
              <th className="ticketReport__th">Half</th>
              <th className="ticketReport__th">ST</th>
              <th className="ticketReport__th">Phy</th>
              <th className="ticketReport__th">Lugg</th>
              <th className="ticketReport__th">Ticket Amt</th>
              <th className="ticketReport__th">Lugg Amt</th>
              <th className="ticketReport__th">Company</th>
              <th className="ticketReport__th">Reference</th>
              <th className="ticketReport__th">Transaction ID</th>
            </tr>
          </thead>
          <tbody className="ticketReport__tbody">
            {currentData.length > 0?(
              currentData.map((transaction, index) => (
                <tr 
                  key={transaction.id} 
                  className={`ticketReport__row ${index % 2 === 0 ? 'ticketReport__row--even' : 'ticketReport__row--odd'}`}
                >
                  <td className="ticketReport__td">{transaction.ticket_number || '-'}</td>
                  <td className="ticketReport__td">{transaction.device_id || '-'}</td>
                  <td className="ticketReport__td">{transaction.trip_number || '-'}</td>
                  <td className="ticketReport__td">{transaction.ticket_date || '-'}</td>
                  <td className="ticketReport__td">{transaction.ticket_time || '-'}</td>
                  <td className="ticketReport__td">{transaction.from_stage || 0}</td>
                  <td className="ticketReport__td">{transaction.to_stage || 0}</td>
                  <td className="ticketReport__td">{transaction.ticket_type || '-'}</td>
                  <td className="ticketReport__td">{transaction.ticket_status || '-'}</td>
                  <td className="ticketReport__td">{transaction.full_count || 0}</td>
                  <td className="ticketReport__td">{transaction.half_count || 0}</td>
                  <td className="ticketReport__td">{transaction.st_count || 0}</td>
                  <td className="ticketReport__td">{transaction.phy_count || 0}</td>
                  <td className="ticketReport__td">{transaction.lugg_count || 0}</td>
                  <td className="ticketReport__td">₹{transaction.ticket_amount || 0}</td>
                  <td className="ticketReport__td">₹{transaction.lugg_amount || 0}</td>
                  <td className="ticketReport__td">{transaction.company_code || '-'}</td>
                  <td className="ticketReport__td">{transaction.reference_number || '-'}</td>
                  <td className="ticketReport__td">{transaction.transaction_id || '-'}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="11" className="ticketReport__empty">No trip data found.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* Pagination Section */}
      {totalPages > 1 && (
        <div className="ticketReport__pagination">
          <button
            className="ticketReport__paginationButton"
            onClick={() => handlePageChange(currentPage - 1)}
            disabled={currentPage === 1}
          >
            Previous
          </button>

          {[...Array(totalPages)].map((_, index) => {
            const pageNumber = index + 1;
            // Show first page, last page, current page, and pages around current
            if (
              pageNumber === 1 ||
              pageNumber === totalPages ||
              (pageNumber >= currentPage - 1 && pageNumber <= currentPage + 1)
            ) {
              return (
                <button
                  key={pageNumber}
                  className={`ticketReport__paginationButton ${
                    currentPage === pageNumber ? 'ticketReport__paginationButton--active' : ''
                  }`}
                  onClick={() => handlePageChange(pageNumber)}
                >
                  {pageNumber}
                </button>
              );
            } else if (
              pageNumber === currentPage - 2 ||
              pageNumber === currentPage + 2
            ) {
              return <span key={pageNumber} className="ticketReport__paginationEllipsis">...</span>;
            }
            return null;
          })}

          <button
            className="ticketReport__paginationButton"
            onClick={() => handlePageChange(currentPage + 1)}
            disabled={currentPage === totalPages}
          >
            Next
          </button>
        </div>
      )}
    </div>
  );
}