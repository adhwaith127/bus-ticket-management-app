import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import SearchBar from '../components/SearchBar';
import { useFilteredList } from '../assets/js/useFilteredList';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function CrewAssignmentListing() {

  // ── Section 1: State Management ──────────────────────────────────────────────
  const [assignments, setAssignments] = useState([]);
  const [loading, setLoading]         = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode]     = useState('create');
  const [submitting, setSubmitting]   = useState(false);
  const [editingItem, setEditingItem] = useState(null);

  // Dropdown data fetched once on mount
  const [drivers, setDrivers]       = useState([]);
  const [conductors, setConductors] = useState([]);
  const [cleaners, setCleaners]     = useState([]);
  const [vehicles, setVehicles]     = useState([]);

  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const emptyForm = { driver: '', conductor: '', cleaner: '', vehicle: '' };
  const [formData, setFormData] = useState(emptyForm);

  // ── Section 2a: Search & Filter Logic ────────────────────────────────────────────
  const { filteredItems, searchTerm, setSearchTerm, resetSearch } = useFilteredList(
    assignments,
    ['driver', 'conductor', 'cleaner', 'vehicle']
  );

  // ── Section 2b: Data Fetching ─────────────────────────────────────────────────
  useEffect(() => {
    fetchAssignments();
    fetchDropdowns();
  }, []);

  const fetchAssignments = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/crew-assignments/`);
      setAssignments(res.data?.data || []);
      setCurrentPage(1);
    } catch (err) {
      console.error('Error fetching crew assignments:', err);
      setAssignments([]);
    } finally {
      setLoading(false);
    }
  };

  // Parallel fetch of all dropdown options for better performance
  const fetchDropdowns = async () => {
    try {
      const [driversRes, conductorsRes, cleanersRes, vehiclesRes] = await Promise.all([
        api.get(`${BASE_URL}/masterdata/dropdowns/employees/`, { params: { type: 'DRIVER' } }),
        api.get(`${BASE_URL}/masterdata/dropdowns/employees/`, { params: { type: 'CONDUCTOR' } }),
        api.get(`${BASE_URL}/masterdata/dropdowns/employees/`, { params: { type: 'CLEANER' } }),
        api.get(`${BASE_URL}/masterdata/dropdowns/vehicles/`),
      ]);
      setDrivers(driversRes.data?.data || []);
      setConductors(conductorsRes.data?.data || []);
      setCleaners(cleanersRes.data?.data || []);
      setVehicles(vehiclesRes.data?.data || []);
    } catch (err) {
      console.error('Error fetching dropdowns:', err);
    }
  };

  const handleSubmit = async () => {
    setSubmitting(true);
    try {
      // Only send non-empty optional fields
      const payload = {
        driver:  formData.driver,
        vehicle: formData.vehicle,
        ...(formData.conductor && { conductor: formData.conductor }),
        ...(formData.cleaner   && { cleaner:   formData.cleaner }),
      };

      let response;
      if (modalMode === 'edit') {
        response = await api.put(`${BASE_URL}/masterdata/crew-assignments/update/${editingItem.id}/`, payload);
      } else {
        response = await api.post(`${BASE_URL}/masterdata/crew-assignments/create/`, payload);
      }
      if (response?.status === 200 || response?.status === 201) {
        window.alert(response.data.message || 'Success');
        setIsModalOpen(false);
        setFormData(emptyForm);
        fetchAssignments();
      }
    } catch (err) {
      if (!err.response) return window.alert('Server unreachable. Try later.');
      const { data } = err.response;
      const firstError = data.errors
        ? Object.values(data.errors)[0][0]
        : (data.error || data.message);
      window.alert(firstError || 'Validation failed');
    } finally {
      setSubmitting(false);
    }
  };

  // ── Section 3b: Pagination Logic ──────────────────────────────────────────────
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentItems = filteredItems.slice(indexOfFirstItem, indexOfLastItem);
  const totalPages = Math.ceil(filteredItems.length / itemsPerPage);

  const getPageNumbers = () => {
    let startPage = Math.max(1, currentPage - 1);
    let endPage = Math.min(totalPages, startPage + 2);
    
    if (endPage - startPage < 2) {
      startPage = Math.max(1, endPage - 2);
    }
    
    const pages = [];
    for (let i = startPage; i <= endPage; i++) {
      pages.push(i);
    }
    return pages;
  };

  // ── Section 4: Modal Helpers ─────────────────────────────────────────────────
  const openCreateModal = () => {
    setFormData(emptyForm);
    setEditingItem(null);
    setModalMode('create');
    setIsModalOpen(true);
  };

  const openViewModal = (item) => {
    setFormData({
      driver:    item.driver    || '',
      conductor: item.conductor || '',
      cleaner:   item.cleaner   || '',
      vehicle:   item.vehicle   || '',
    });
    setEditingItem(item);
    setModalMode('view');
    setIsModalOpen(true);
  };

  const openEditModal = (item) => {
    setFormData({
      driver:    item.driver    || '',
      conductor: item.conductor || '',
      cleaner:   item.cleaner   || '',
      vehicle:   item.vehicle   || '',
    });
    setEditingItem(item);
    setModalMode('edit');
    setIsModalOpen(true);
  };

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const isReadOnly    = modalMode === 'view';
  const getModalTitle = () => ({ view: 'Crew Assignment Details', edit: 'Edit Crew Assignment', create: 'Create Crew Assignment' }[modalMode]);

  // Renders dropdown in edit/create mode, text field in view mode
  const renderDropdown = (name, label, options, required = false) => (
    <div className="space-y-2">
      <label className="text-sm font-semibold text-slate-700">{label}{required ? ' *' : ' (optional)'}</label>
      {isReadOnly ? (
        <input type="text"
          value={options.find(o => String(o.id) === String(formData[name]))?.employee_name
              || options.find(o => String(o.id) === String(formData[name]))?.bus_reg_num
              || '—'}
          readOnly 
          className="w-full px-4 py-2.5 border border-slate-300 rounded-xl bg-slate-50 text-slate-600"
        />
      ) : (
        <select name={name} value={formData[name]} onChange={handleInputChange}
          className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent bg-white transition-all">
          <option value="">-- {required ? 'Select' : 'None'} --</option>
          {options.map(o => (
            <option key={o.id} value={o.id}>
              {o.employee_name
                ? `${o.employee_name} (${o.employee_code})`
                : o.bus_reg_num}
            </option>
          ))}
        </select>
      )}
    </div>
  );

  // ── Section 5: Render ────────────────────────────────────────────────────────
  return (
    <div className="p-6 md:p-10 min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">

      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
        <div>
          <h1 className="text-3xl font-bold bg-gradient-to-r from-slate-800 to-slate-600 bg-clip-text text-transparent tracking-tight">
            Crew Assignments
          </h1>
          <p className="text-slate-600 mt-1.5">Assign drivers, conductors and cleaners to vehicles</p>
        </div>
        <button onClick={openCreateModal} className="flex items-center justify-center bg-gradient-to-r from-slate-800 to-slate-700 hover:from-slate-700 hover:to-slate-600 text-white px-6 py-3 rounded-xl transition-all duration-200 shadow-lg hover:shadow-xl transform hover:-translate-y-0.5">
          <span className="mr-2 text-lg">+</span>
          <span className="font-medium">Create Assignment</span>
        </button>
      </div>
      {/* Search Bar */}
      <SearchBar
        searchTerm={searchTerm}
        onSearchChange={setSearchTerm}
        onReset={resetSearch}
        placeholder="Search by driver, conductor, cleaner, or vehicle..."
      />
      {/* Enhanced Table */}
      <div className="bg-white rounded-2xl shadow-lg border border-slate-200/60 overflow-hidden backdrop-blur-sm">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead>
              <tr className="bg-gradient-to-r from-slate-50 to-slate-100/50 border-b border-slate-200">
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">ID</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Vehicle</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Driver</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Conductor</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Cleaner</th>
                <th className="px-6 py-4 text-right text-xs font-semibold text-slate-600 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr>
                  <td colSpan="6" className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center">
                      <div className="animate-spin rounded-full h-10 w-10 border-b-2 border-slate-800"></div>
                      <p className="text-slate-500 mt-3">Loading assignments...</p>
                    </div>
                  </td>
                </tr>
              ) : currentItems.length === 0 ? (
                <tr>
                  <td colSpan="6" className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center">
                      <div className="rounded-full bg-slate-100 p-3 mb-3">
                        <svg className="w-6 h-6 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z" />
                        </svg>
                      </div>
                      <p className="text-slate-500 font-medium">No crew assignments found</p>
                      <p className="text-slate-400 text-sm mt-1">Create your first assignment to get started</p>
                    </div>
                  </td>
                </tr>
              ) : currentItems.map(item => (
                <tr key={item.id} className="hover:bg-slate-50/80 transition-all duration-150">
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-500 font-mono">#{item.id}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-800 font-semibold">{item.vehicle_reg || '—'}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-700">{item.driver_name || '—'}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-600">{item.conductor_name || '—'}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-600">{item.cleaner_name || '—'}</span>
                  </td>
                  <td className="px-6 py-4">
                    <div className="flex justify-end items-center gap-2">
                      <button onClick={() => openViewModal(item)} className="px-3 py-1.5 text-xs font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-all duration-150">
                        View
                      </button>
                      <button onClick={() => openEditModal(item)} className="px-3 py-1.5 text-xs font-medium text-blue-600 hover:text-blue-700 hover:bg-blue-50 rounded-lg transition-all duration-150">
                        Edit
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Pagination */}
        {!loading && assignments.length > 0 && totalPages > 1 && (
          <div className="px-6 py-4 border-t border-slate-200 bg-slate-50/50">
            <div className="flex items-center justify-between">
              <div className="text-sm text-slate-600">
                Showing <span className="font-medium text-slate-900">{indexOfFirstItem + 1}</span> to{' '}
                <span className="font-medium text-slate-900">{Math.min(indexOfLastItem, assignments.length)}</span> of{' '}
                <span className="font-medium text-slate-900">{assignments.length}</span> results
              </div>
              
              <div className="flex items-center gap-2">
                <button
                  onClick={() => setCurrentPage(prev => prev - 1)}
                  disabled={currentPage === 1}
                  className="px-3 py-1.5 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-150"
                >
                  Previous
                </button>

                <div className="flex items-center gap-1">
                  {getPageNumbers().map(pageNum => (
                    <button
                      key={pageNum}
                      onClick={() => setCurrentPage(pageNum)}
                      className={`min-w-[2.5rem] px-3 py-1.5 text-sm font-medium rounded-lg transition-all duration-150 ${
                        currentPage === pageNum
                          ? 'bg-slate-800 text-white shadow-md'
                          : 'text-slate-700 bg-white border border-slate-300 hover:bg-slate-50'
                      }`}
                    >
                      {pageNum}
                    </button>
                  ))}
                </div>

                <button
                  onClick={() => setCurrentPage(prev => prev + 1)}
                  disabled={currentPage === totalPages}
                  className="px-3 py-1.5 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-150"
                >
                  Next
                </button>
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Modal */}
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={getModalTitle()}>
        <div className="space-y-5">

          {renderDropdown('driver',    'Driver',    drivers,    true)}
          {renderDropdown('conductor', 'Conductor', conductors, false)}
          {renderDropdown('cleaner',   'Cleaner',   cleaners,   false)}
          {renderDropdown('vehicle',   'Vehicle',   vehicles,   true)}

          <div className="flex items-center justify-end gap-3 pt-6 border-t border-slate-200">
            <button type="button" onClick={() => setIsModalOpen(false)} className="px-5 py-2.5 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-xl hover:bg-slate-50 transition-all">
              {isReadOnly ? 'Close' : 'Cancel'}
            </button>
            {!isReadOnly && (
              <button type="button" onClick={handleSubmit} disabled={submitting} className="px-5 py-2.5 text-sm font-medium text-white bg-slate-800 rounded-xl hover:bg-slate-700 shadow-md hover:shadow-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                {submitting ? 'Saving...' : modalMode === 'edit' ? 'Update' : 'Save'}
              </button>
            )}
          </div>
        </div>
      </Modal>

    </div>
  );
}
