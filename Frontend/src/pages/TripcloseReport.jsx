import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import '../styles/TripcloseReport.css';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function TripcloseReport() {
  // ==================== STATE MANAGEMENT ====================
  const [tripData, setTripData] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(10);

  // Filters
  const [filters, setFilters] = useState({
    startDate: '',
    endDate: '',
    palmtecId: '', // Maps to device_id concept
    routeCode: '',
    tripNo: ''
  });

  // ==================== DATA FETCHING ====================
  useEffect(() => {
    fetchTripData();
  }, []);

  const fetchTripData = async () => {
    try {
      setLoading(true);
      const response = await api.get(`${BASE_URL}/get_all_trip_close_data`);
      
      if (response.data.message === 'success') {
        setTripData(response.data.data);
      } else {
        // Fallback for different response structures
        setTripData(response.data.data || []);
      }
    } catch (err) {
      console.error("Fetch Error:", err);
      if (err.response) {
        setError(`Server Error: ${err.response.status} - ${err.response.data.message || 'Unknown error'}`);
      } else if (err.request) {
        setError('No response from server. Check connection.');
      } else {
        setError('Error setting up request: ' + err.message);
      }
    } finally {
      setLoading(false);
    }
  };

  // ==================== FILTER LOGIC ====================
  const getFilteredData = () => {
    return tripData.filter(item => {
      // 1. Date Range Filter (using start_datetime)
      if (filters.startDate) {
        const itemDate = new Date(item.start_datetime);
        const filterStart = new Date(filters.startDate);
        if (itemDate < filterStart) return false;
      }
      if (filters.endDate) {
        const itemDate = new Date(item.start_datetime);
        const filterEnd = new Date(filters.endDate);
        filterEnd.setHours(23, 59, 59, 999); // Include full end day
        if (itemDate > filterEnd) return false;
      }

      // 2. Text Filters
      if (filters.palmtecId && item.palmtec_id) {
        if (!item.palmtec_id.toLowerCase().includes(filters.palmtecId.toLowerCase())) return false;
      }
      if (filters.routeCode && item.route_code) {
        if (!item.route_code.toLowerCase().includes(filters.routeCode.toLowerCase())) return false;
      }
      if (filters.tripNo && item.trip_no) {
        if (!String(item.trip_no).includes(filters.tripNo)) return false;
      }

      return true;
    });
  };

  // ==================== PAGINATION ====================
  const filteredData = getFilteredData();
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);
  const startIndex = (currentPage - 1) * itemsPerPage;
  const endIndex = startIndex + itemsPerPage;
  const currentData = filteredData.slice(startIndex, endIndex);

  const handlePageChange = (page) => setCurrentPage(page);

  const handleFilterChange = (key, value) => {
    setFilters(prev => ({ ...prev, [key]: value }));
    setCurrentPage(1);
  };

  const clearFilters = () => {
    setFilters({ startDate: '', endDate: '', palmtecId: '', routeCode: '', tripNo: '' });
    setCurrentPage(1);
  };

  // ==================== EXCEL EXPORT ====================
  const exportToExcel = () => {
    const exportData = filteredData.map(item => {
      // Calculate Cash Collection derived field
      const totalColl = parseFloat(item.total_collection || 0);
      const upiColl = parseFloat(item.upi_ticket_amount || 0);
      const cashColl = totalColl - upiColl;

      return {
        'ID': item.id,
        'Device ID': item.palmtec_id,
        'Route': item.route_code,
        'Trip No': item.trip_no,
        'Schedule': item.schedule,
        'Direction': item.up_down_trip,
        'Start Date': new Date(item.start_datetime).toLocaleDateString(),
        'Start Time': new Date(item.start_datetime).toLocaleTimeString(),
        'End Time': item.end_datetime ? new Date(item.end_datetime).toLocaleTimeString() : '-',
        
        // Ticket Numbers
        'Start Ticket': item.start_ticket_no,
        'End Ticket': item.end_ticket_no,
        'Total Tickets': item.total_tickets_issued, // From Serializer

        // Passenger Counts
        'Full Pax': item.full_count,
        'Half Pax': item.half_count,
        'Student Pax': item.st1_count,
        'Ladies Pax': item.ladies_count,
        'Senior Pax': item.senior_count,
        'Phy Handicap': item.physical_count,
        'Luggage Count': item.luggage_count,
        'Pass Holders': item.pass_count,
        'Total Pax': item.total_passengers, // From Serializer

        // Financials
        'Full Coll': item.full_collection,
        'Half Coll': item.half_collection,
        'Luggage Coll': item.luggage_collection,
        'UPI Amount': item.upi_ticket_amount,
        'Cash Amount': cashColl.toFixed(2),
        'Expenses': item.expense_amount,
        'Total Collection': item.total_collection
      };
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'TripData');
    const fileName = `trip_report_${new Date().toISOString().split('T')[0]}.xlsx`;
    XLSX.writeFile(workbook, fileName);
  };

  // ==================== RENDER ====================
  if (loading) return <div className="tripReport"><div className="tripReport__loading">Loading Trip Data...</div></div>;
  if (error) return <div className="tripReport"><div className="tripReport__error">{error}</div></div>;

  return (
    <div className="tripReport">
      {/* --- HEADER --- */}
      <div className="tripReport__header">
        <div>
          <h1 className="tripReport__title">Trip Close Reports</h1>
          <p className="tripReport__subtitle">View and manage daily trip closures</p>
        </div>
        <button className="tripReport__btn tripReport__btn--primary" onClick={exportToExcel}>
          Download Report
        </button>
      </div>

      {/* --- FILTERS --- */}
      <div className="tripReport__filters">
        <div className="tripReport__filterGroup">
          <label>Start Date</label>
          <input 
            type="date" 
            value={filters.startDate}
            onChange={(e) => handleFilterChange('startDate', e.target.value)}
          />
        </div>
        <div className="tripReport__filterGroup">
          <label>End Date</label>
          <input 
            type="date" 
            value={filters.endDate}
            onChange={(e) => handleFilterChange('endDate', e.target.value)}
          />
        </div>
        <div className="tripReport__filterGroup">
          <label>Device ID</label>
          <input 
            type="text" 
            placeholder="Search Palmtec ID..."
            value={filters.palmtecId}
            onChange={(e) => handleFilterChange('palmtecId', e.target.value)}
          />
        </div>
        <div className="tripReport__filterGroup">
          <label>Route</label>
          <input 
            type="text" 
            placeholder="Search Route..."
            value={filters.routeCode}
            onChange={(e) => handleFilterChange('routeCode', e.target.value)}
          />
        </div>
        <div className="tripReport__filterGroup">
          <label>Trip No</label>
          <input 
            type="number" 
            placeholder="#"
            value={filters.tripNo}
            onChange={(e) => handleFilterChange('tripNo', e.target.value)}
          />
        </div>
        <div className="tripReport__filterActions">
          <button className="tripReport__btn tripReport__btn--secondary" onClick={clearFilters}>
            Clear
          </button>
        </div>
      </div>

      {/* --- SUMMARY --- */}
      <div className="tripReport__summary">
        Showing <b>{currentData.length}</b> of <b>{filteredData.length}</b> records
      </div>

      {/* --- TABLE --- */}
      <div className="tripReport__tableWrapper">
        <table className="tripReport__table">
          <thead>
            <tr>
              <th>Date</th>
              <th>Device ID</th>
              <th>Route</th>
              <th>Trip</th>
              <th>Sched</th>
              <th>Direction</th>
              <th className="text-right">Total Pax</th>
              <th className="text-right">Tickets</th>
              <th className="text-right">UPI Amt</th>
              <th className="text-right">Expense</th>
              <th className="text-right">Total Coll.</th>
            </tr>
          </thead>
          <tbody>
            {currentData.length > 0 ? (
              currentData.map((item) => (
                <tr key={item.id}>
                  <td>
                    <div className="tripReport__dateTime">
                      <span>{new Date(item.start_datetime).toLocaleDateString()}</span>
                      <small>{new Date(item.start_datetime).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</small>
                    </div>
                  </td>
                  <td><span className="tripReport__badge">{item.palmtec_id}</span></td>
                  <td>{item.route_code || '-'}</td>
                  <td>{item.trip_no}</td>
                  <td>{item.schedule}</td>
                  <td>{item.up_down_trip}</td>
                  
                  {/* Counts from Serializer */}
                  <td className="text-right">{item.total_passengers}</td>
                  <td className="text-right">{item.total_tickets_issued}</td>
                  
                  {/* Financials */}
                  <td className="text-right">₹{item.upi_ticket_amount}</td>
                  <td className="text-right">₹{item.expense_amount}</td>
                  <td className="text-right tripReport__amount">₹{item.total_collection}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan="11" className="tripReport__empty">No trip data found.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>

      {/* --- PAGINATION --- */}
      {totalPages > 1 && (
        <div className="tripReport__pagination">
          <button 
            disabled={currentPage === 1}
            onClick={() => handlePageChange(currentPage - 1)}
          >
            Previous
          </button>
          <span>Page {currentPage} of {totalPages}</span>
          <button 
            disabled={currentPage === totalPages}
            onClick={() => handlePageChange(currentPage + 1)}
          >
            Next
          </button>
        </div>
      )}
    </div>
  );
}