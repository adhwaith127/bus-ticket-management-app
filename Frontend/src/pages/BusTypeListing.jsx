import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import SearchBar from '../components/SearchBar';
import { useFilteredList } from '../assets/js/useFilteredList';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function BusTypeListing() {

  // ── Section 1: State Management ──────────────────────────────────────────────
  // Core data and modal states
  const [busTypes, setBusTypes]       = useState([]);
  const [loading, setLoading]         = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode]     = useState('create');
  const [submitting, setSubmitting]   = useState(false);
  const [editingItem, setEditingItem] = useState(null);

  // Pagination states - tracks current page for UI-side pagination
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const emptyForm = { bustype_code: '', name: '', is_active: true };
  const [formData, setFormData] = useState(emptyForm);

  // ── Section 2: Search & Filter Logic ───────────────────────────────────────
  const { filteredItems, searchTerm, setSearchTerm, resetSearch } = useFilteredList(
    busTypes,
    ['bustype_code', 'name']
  );

  // ── Section 3a: Data Fetching ────────────────────────────────────────────────
  useEffect(() => { fetchBusTypes(); }, []);

  const fetchBusTypes = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/bus-types/`);
      setBusTypes(res.data?.data || []);
      setCurrentPage(1); // Reset to page 1 when data refreshes
    } catch (err) {
      console.error('Error fetching bus types:', err);
      setBusTypes([]);
    } finally {
      setLoading(false);
    }
  };

  const handleSubmit = async () => {
    setSubmitting(true);
    try {
      let response;
      if (modalMode === 'edit') {
        response = await api.put(`${BASE_URL}/masterdata/bus-types/update/${editingItem.id}/`, formData);
      } else {
        response = await api.post(`${BASE_URL}/masterdata/bus-types/create/`, formData);
      }
      if (response?.status === 200 || response?.status === 201) {
        window.alert(response.data.message || 'Success');
        setIsModalOpen(false);
        setFormData(emptyForm);
        fetchBusTypes();
      }
    } catch (err) {
      if (!err.response) return window.alert('Server unreachable. Try later.');
      const { data } = err.response;
      const firstError = data.errors ? Object.values(data.errors)[0][0] : data.message;
      window.alert(firstError || 'Validation failed');
    } finally {
      setSubmitting(false);
    }
  };

  // ── Section 3b: Pagination Logic ─────────────────────────────────────────────
  // Calculate which items to display on current page
  // Example: Page 2 with 10 items per page shows items 10-19
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentItems = filteredItems.slice(indexOfFirstItem, indexOfLastItem);
  const totalPages = Math.ceil(filteredItems.length / itemsPerPage);

  // Generate array of page numbers to display (max 3 at a time)
  // Logic: Show current page ± 1, but keep within bounds
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
    setFormData(item);
    setEditingItem(item);
    setModalMode('view');
    setIsModalOpen(true);
  };

  const openEditModal = (item) => {
    setFormData(item);
    setEditingItem(item);
    setModalMode('edit');
    setIsModalOpen(true);
  };

  const handleInputChange = (e) => {
    const { name, value, type, checked } = e.target;
    setFormData(prev => ({ ...prev, [name]: type === 'checkbox' ? checked : value }));
  };

  const isReadOnly    = modalMode === 'view';
  const getModalTitle = () => ({ view: 'Bus Type Details', edit: 'Edit Bus Type', create: 'Create Bus Type' }[modalMode]);

  const getStatusBadge = (active) => active
    ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
    : 'bg-rose-50 text-rose-700 border-rose-200';

  // ── Section 5: Render ────────────────────────────────────────────────────────
  return (
    <div className="p-6 md:p-10 min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">

      {/* Header with gradient background */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
        <div>
          <h1 className="text-3xl font-bold bg-gradient-to-r from-slate-800 to-slate-600 bg-clip-text text-transparent tracking-tight">
            Bus Types
          </h1>
          <p className="text-slate-600 mt-1.5">Manage bus categories for your company</p>
        </div>
        <button
          onClick={openCreateModal}
          className="flex items-center justify-center bg-gradient-to-r from-slate-800 to-slate-700 hover:from-slate-700 hover:to-slate-600 text-white px-6 py-3 rounded-xl transition-all duration-200 shadow-lg hover:shadow-xl transform hover:-translate-y-0.5"
        >
          <span className="mr-2 text-lg">+</span>
          <span className="font-medium">Create Bus Type</span>
        </button>
      </div>

      {/* Search Bar */}
      <SearchBar
        searchTerm={searchTerm}
        onSearchChange={setSearchTerm}
        onReset={resetSearch}
        placeholder="Search by code or name..."
      />

      {/* Enhanced Table Card */}
      <div className="bg-white rounded-2xl shadow-lg border border-slate-200/60 overflow-hidden backdrop-blur-sm">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead>
              <tr className="bg-gradient-to-r from-slate-50 to-slate-100/50 border-b border-slate-200">
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">ID</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Code</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Name</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Status</th>
                <th className="px-6 py-4 text-right text-xs font-semibold text-slate-600 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr>
                  <td colSpan="5" className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center">
                      <div className="animate-spin rounded-full h-10 w-10 border-b-2 border-slate-800"></div>
                      <p className="text-slate-500 mt-3">Loading bus types...</p>
                    </div>
                  </td>
                </tr>
              ) : currentItems.length === 0 ? (
                <tr>
                  <td colSpan="5" className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center">
                      <div className="rounded-full bg-slate-100 p-3 mb-3">
                        <svg className="w-6 h-6 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M20 13V6a2 2 0 00-2-2H6a2 2 0 00-2 2v7m16 0v5a2 2 0 01-2 2H6a2 2 0 01-2-2v-5m16 0h-2.586a1 1 0 00-.707.293l-2.414 2.414a1 1 0 01-.707.293h-3.172a1 1 0 01-.707-.293l-2.414-2.414A1 1 0 006.586 13H4" />
                        </svg>
                      </div>
                      <p className="text-slate-500 font-medium">No bus types found</p>
                      <p className="text-slate-400 text-sm mt-1">Create your first bus type to get started</p>
                    </div>
                  </td>
                </tr>
              ) : currentItems.map(item => (
                <tr key={item.id} className="hover:bg-slate-50/80 transition-all duration-150 group">
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-500 font-mono">#{item.id}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-800 font-semibold">{item.bustype_code}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-700">{item.name}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className={`inline-flex items-center px-3 py-1 rounded-full text-xs font-medium border transition-colors ${getStatusBadge(item.is_active)}`}>
                      <span className={`w-1.5 h-1.5 rounded-full mr-1.5 ${item.is_active ? 'bg-emerald-500' : 'bg-rose-500'}`}></span>
                      {item.is_active ? 'Active' : 'Inactive'}
                    </span>
                  </td>
                  <td className="px-6 py-4">
                    <div className="flex justify-end items-center gap-2">
                      <button 
                        onClick={() => openViewModal(item)} 
                        className="px-3 py-1.5 text-xs font-medium text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded-lg transition-all duration-150"
                      >
                        View
                      </button>
                      <button 
                        onClick={() => openEditModal(item)} 
                        className="px-3 py-1.5 text-xs font-medium text-blue-600 hover:text-blue-700 hover:bg-blue-50 rounded-lg transition-all duration-150"
                      >
                        Edit
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Pagination Controls - Only show if there's data and multiple pages */}
        {!loading && filteredItems.length > 0 && totalPages > 1 && (
          <div className="px-6 py-4 border-t border-slate-200 bg-slate-50/50">
            <div className="flex items-center justify-between">
              <div className="text-sm text-slate-600">
                Showing <span className="font-medium text-slate-900">{indexOfFirstItem + 1}</span> to{' '}
                <span className="font-medium text-slate-900">{Math.min(indexOfLastItem, busTypes.length)}</span> of{' '}
                <span className="font-medium text-slate-900">{busTypes.length}</span> results
              </div>
              
              <div className="flex items-center gap-2">
                {/* Previous Button */}
                <button
                  onClick={() => setCurrentPage(prev => prev - 1)}
                  disabled={currentPage === 1}
                  className="px-3 py-1.5 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50 disabled:opacity-50 disabled:cursor-not-allowed transition-all duration-150"
                >
                  Previous
                </button>

                {/* Page Numbers */}
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

                {/* Next Button */}
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

          <div className="space-y-2">
            <label className="text-sm font-semibold text-slate-700">Bus Type Code *</label>
            <input
              type="text" name="bustype_code" value={formData.bustype_code}
              onChange={handleInputChange} readOnly={isReadOnly}
              className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent read-only:bg-slate-50 read-only:text-slate-600 transition-all"
              placeholder="e.g., BT001"
            />
          </div>

          <div className="space-y-2">
            <label className="text-sm font-semibold text-slate-700">Name *</label>
            <input
              type="text" name="name" value={formData.name}
              onChange={handleInputChange} readOnly={isReadOnly}
              className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent read-only:bg-slate-50 read-only:text-slate-600 transition-all"
              placeholder="e.g., Luxury Coach"
            />
          </div>

          {modalMode !== 'create' && (
            <div className="flex items-center gap-3 p-4 bg-slate-50 rounded-xl border border-slate-200">
              <input
                type="checkbox" name="is_active" id="is_active"
                checked={formData.is_active} onChange={handleInputChange}
                disabled={isReadOnly}
                className="w-4 h-4 rounded border-slate-300 text-slate-800 focus:ring-slate-500"
              />
              <label htmlFor="is_active" className="text-sm font-medium text-slate-700">Active Status</label>
            </div>
          )}

          <div className="flex items-center justify-end gap-3 pt-6 border-t border-slate-200">
            <button
              type="button" onClick={() => setIsModalOpen(false)}
              className="px-5 py-2.5 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-xl hover:bg-slate-50 transition-all"
            >
              {isReadOnly ? 'Close' : 'Cancel'}
            </button>
            {!isReadOnly && (
              <button
                type="button" onClick={handleSubmit} disabled={submitting}
                className="px-5 py-2.5 text-sm font-medium text-white bg-slate-800 rounded-xl hover:bg-slate-700 shadow-md hover:shadow-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all"
              >
                {submitting ? 'Saving...' : modalMode === 'edit' ? 'Update' : 'Save'}
              </button>
            )}
          </div>
        </div>
      </Modal>

    </div>
  );
}
