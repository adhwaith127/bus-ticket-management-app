import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import api, { BASE_URL } from '../assets/js/axiosConfig';

// Boolean toggle flags for route configuration
// Stored as array for easy iteration when rendering checkboxes
const ROUTE_FLAGS = [
  { name: 'half',       label: 'Half Fare'           },
  { name: 'luggage',    label: 'Luggage'              },
  { name: 'student',    label: 'Student Concession'   },
  { name: 'adjust',     label: 'Fare Adjustment'      },
  { name: 'conc',       label: 'General Concession'   },
  { name: 'ph',         label: 'PH Concession'        },
  { name: 'pass_allow', label: 'Pass Holders'         },
  { name: 'use_stop',   label: 'Stop-based Fare'      },
];

export default function RouteListing() {

  // ── Section 1: State Management ──────────────────────────────────────────────
  const [routes, setRoutes]           = useState([]);
  const [busTypes, setBusTypes]       = useState([]);
  const [loading, setLoading]         = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode]     = useState('create');
  const [submitting, setSubmitting]   = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [showDeleted, setShowDeleted] = useState(false);

  // Pagination state
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 10;

  const emptyForm = {
    route_code: '', route_name: '', min_fare: '', fare_type: '',
    bus_type: '', start_from: 0, is_deleted: false,
    half: false, luggage: false, student: false, adjust: false,
    conc: false, ph: false, pass_allow: false, use_stop: false,
  };
  const [formData, setFormData] = useState(emptyForm);

  // ── Section 2: Data Fetching ─────────────────────────────────────────────────
  useEffect(() => { fetchBusTypes(); }, []);
  useEffect(() => { fetchRoutes(); }, [showDeleted]);

  const fetchRoutes = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/routes/`, {
        params: { show_deleted: showDeleted }
      });
      setRoutes(res.data?.data || []);
      setCurrentPage(1);
    } catch (err) {
      console.error('Error fetching routes:', err);
      setRoutes([]);
    } finally {
      setLoading(false);
    }
  };

  const fetchBusTypes = async () => {
    try {
      const res = await api.get(`${BASE_URL}/masterdata/dropdowns/bus-types/`);
      setBusTypes(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching bus types:', err);
    }
  };

  const handleSubmit = async () => {
    setSubmitting(true);
    try {
      let response;
      if (modalMode === 'edit') {
        response = await api.put(`${BASE_URL}/masterdata/routes/update/${editingItem.id}/`, formData);
      } else {
        response = await api.post(`${BASE_URL}/masterdata/routes/create/`, formData);
      }
      if (response?.status === 200 || response?.status === 201) {
        window.alert(response.data.message || 'Success');
        setIsModalOpen(false);
        setFormData(emptyForm);
        fetchRoutes();
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

  // ── Section 3: Pagination Logic ──────────────────────────────────────────────
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentItems = routes.slice(indexOfFirstItem, indexOfLastItem);
  const totalPages = Math.ceil(routes.length / itemsPerPage);

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
  const openCreateModal = () => { setFormData(emptyForm); setEditingItem(null); setModalMode('create'); setIsModalOpen(true); };

  const openViewModal = (item) => {
    setFormData({ ...emptyForm, ...item, bus_type: item.bus_type });
    setEditingItem(item);
    setModalMode('view');
    setIsModalOpen(true);
  };

  const openEditModal = (item) => {
    setFormData({ ...emptyForm, ...item, bus_type: item.bus_type });
    setEditingItem(item);
    setModalMode('edit');
    setIsModalOpen(true);
  };

  const handleInputChange = (e) => {
    const { name, value, type, checked } = e.target;
    setFormData(prev => ({ ...prev, [name]: type === 'checkbox' ? checked : value }));
  };

  const isReadOnly    = modalMode === 'view';
  const getModalTitle = () => ({ view: 'Route Details', edit: 'Edit Route', create: 'Create Route' }[modalMode]);

  // ── Section 5: Render ────────────────────────────────────────────────────────
  return (
    <div className="p-6 md:p-10 min-h-screen bg-gradient-to-br from-slate-50 to-slate-100">

      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
        <div>
          <h1 className="text-3xl font-bold bg-gradient-to-r from-slate-800 to-slate-600 bg-clip-text text-transparent tracking-tight">
            Routes
          </h1>
          <p className="text-slate-600 mt-1.5">Manage bus routes for your company</p>
        </div>
        <div className="flex items-center gap-3">
          <label className="flex items-center gap-2 px-4 py-2 text-sm text-slate-700 bg-white border border-slate-300 rounded-xl cursor-pointer hover:bg-slate-50 transition-all">
            <input type="checkbox" checked={showDeleted} onChange={() => setShowDeleted(p => !p)} className="w-4 h-4 rounded border-slate-300 text-slate-800 focus:ring-slate-500" />
            <span className="font-medium">Show deleted</span>
          </label>
          <button onClick={openCreateModal} className="flex items-center justify-center bg-gradient-to-r from-slate-800 to-slate-700 hover:from-slate-700 hover:to-slate-600 text-white px-6 py-3 rounded-xl transition-all duration-200 shadow-lg hover:shadow-xl transform hover:-translate-y-0.5">
            <span className="mr-2 text-lg">+</span>
            <span className="font-medium">Create Route</span>
          </button>
        </div>
      </div>

      {/* Enhanced Table */}
      <div className="bg-white rounded-2xl shadow-lg border border-slate-200/60 overflow-hidden backdrop-blur-sm">
        <div className="overflow-x-auto">
          <table className="w-full">
            <thead>
              <tr className="bg-gradient-to-r from-slate-50 to-slate-100/50 border-b border-slate-200">
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">ID</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Code</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Name</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Bus Type</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Min Fare</th>
                <th className="px-6 py-4 text-left text-xs font-semibold text-slate-600 uppercase tracking-wider">Status</th>
                <th className="px-6 py-4 text-right text-xs font-semibold text-slate-600 uppercase tracking-wider">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr>
                  <td colSpan="7" className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center">
                      <div className="animate-spin rounded-full h-10 w-10 border-b-2 border-slate-800"></div>
                      <p className="text-slate-500 mt-3">Loading routes...</p>
                    </div>
                  </td>
                </tr>
              ) : currentItems.length === 0 ? (
                <tr>
                  <td colSpan="7" className="px-6 py-12 text-center">
                    <div className="flex flex-col items-center justify-center">
                      <div className="rounded-full bg-slate-100 p-3 mb-3">
                        <svg className="w-6 h-6 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                          <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M9 20l-5.447-2.724A1 1 0 013 16.382V5.618a1 1 0 011.447-.894L9 7m0 13l6-3m-6 3V7m6 10l4.553 2.276A1 1 0 0021 18.382V7.618a1 1 0 00-.553-.894L15 4m0 13V4m0 0L9 7" />
                        </svg>
                      </div>
                      <p className="text-slate-500 font-medium">No routes found</p>
                      <p className="text-slate-400 text-sm mt-1">Create your first route to get started</p>
                    </div>
                  </td>
                </tr>
              ) : currentItems.map(item => (
                <tr key={item.id} className="hover:bg-slate-50/80 transition-all duration-150">
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-500 font-mono">#{item.id}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-800 font-semibold">{item.route_code}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-700">{item.route_name}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="inline-flex items-center px-2.5 py-1 rounded-full text-xs font-medium bg-purple-50 text-purple-700 border border-purple-200">
                      {item.bus_type_name || '—'}
                    </span>
                  </td>
                  <td className="px-6 py-4">
                    <span className="text-sm text-slate-800 font-semibold">₹{item.min_fare}</span>
                  </td>
                  <td className="px-6 py-4">
                    <span className={`inline-flex items-center px-3 py-1 rounded-full text-xs font-medium border ${item.is_deleted ? 'bg-rose-50 text-rose-700 border-rose-200' : 'bg-emerald-50 text-emerald-700 border-emerald-200'}`}>
                      <span className={`w-1.5 h-1.5 rounded-full mr-1.5 ${item.is_deleted ? 'bg-rose-500' : 'bg-emerald-500'}`}></span>
                      {item.is_deleted ? 'Deleted' : 'Active'}
                    </span>
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
        {!loading && routes.length > 0 && totalPages > 1 && (
          <div className="px-6 py-4 border-t border-slate-200 bg-slate-50/50">
            <div className="flex items-center justify-between">
              <div className="text-sm text-slate-600">
                Showing <span className="font-medium text-slate-900">{indexOfFirstItem + 1}</span> to{' '}
                <span className="font-medium text-slate-900">{Math.min(indexOfLastItem, routes.length)}</span> of{' '}
                <span className="font-medium text-slate-900">{routes.length}</span> results
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

          {/* Basic Info */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-700">Route Code *</label>
              <input type="text" name="route_code" value={formData.route_code} onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent read-only:bg-slate-50 read-only:text-slate-600 transition-all"
                placeholder="e.g., R001" />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-700">Route Name *</label>
              <input type="text" name="route_name" value={formData.route_name} onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent read-only:bg-slate-50 read-only:text-slate-600 transition-all"
                placeholder="e.g., Central to Airport" />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-700">Min Fare (₹) *</label>
              <input type="number" name="min_fare" value={formData.min_fare} onChange={handleInputChange} readOnly={isReadOnly} min="0"
                className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent read-only:bg-slate-50 read-only:text-slate-600 transition-all"
                placeholder="e.g., 10" />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-700">Fare Type *</label>
              <input type="number" name="fare_type" value={formData.fare_type} onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent read-only:bg-slate-50 read-only:text-slate-600 transition-all"
                placeholder="e.g., 1" />
            </div>
            <div className="space-y-2">
              <label className="text-sm font-semibold text-slate-700">Start From (Stage)</label>
              <input type="number" name="start_from" value={formData.start_from} onChange={handleInputChange} readOnly={isReadOnly} min="0"
                className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent read-only:bg-slate-50 read-only:text-slate-600 transition-all"
                placeholder="0" />
            </div>
          </div>

          <div className="space-y-2">
            <label className="text-sm font-semibold text-slate-700">Bus Type *</label>
            {isReadOnly ? (
              <input type="text" value={formData.bus_type_name || '—'} readOnly className="w-full px-4 py-2.5 border border-slate-300 rounded-xl bg-slate-50 text-slate-600" />
            ) : (
              <select name="bus_type" value={formData.bus_type} onChange={handleInputChange}
                className="w-full px-4 py-2.5 border border-slate-300 rounded-xl focus:ring-2 focus:ring-slate-500 focus:border-transparent bg-white transition-all">
                <option value="">-- Select Bus Type --</option>
                {busTypes.map(bt => (
                  <option key={bt.id} value={bt.id}>{bt.name} ({bt.bustype_code})</option>
                ))}
              </select>
            )}
          </div>

          {/* Boolean flags rendered as toggle pills */}
          <div>
            <p className="text-sm font-semibold text-slate-700 mb-3">Allowed Options</p>
            <div className="grid grid-cols-2 md:grid-cols-4 gap-3">
              {ROUTE_FLAGS.map(flag => (
                <label key={flag.name} className={`flex items-center justify-center gap-2 text-sm cursor-pointer px-3 py-2.5 rounded-xl border-2 font-medium transition-all duration-150 ${
                  formData[flag.name] 
                    ? 'bg-slate-800 text-white border-slate-800 shadow-md' 
                    : 'bg-white text-slate-700 border-slate-200 hover:border-slate-300'
                } ${isReadOnly ? 'cursor-default' : ''}`}>
                  <input
                    type="checkbox" name={flag.name}
                    checked={formData[flag.name] || false}
                    onChange={handleInputChange}
                    disabled={isReadOnly}
                    className="sr-only"
                  />
                  {flag.label}
                </label>
              ))}
            </div>
          </div>

          {/* Soft delete toggle */}
          {modalMode === 'edit' && (
            <div className="flex items-center gap-3 p-4 bg-rose-50 rounded-xl border border-rose-200">
              <input type="checkbox" name="is_deleted" id="route_is_deleted" checked={formData.is_deleted || false} onChange={handleInputChange} className="w-4 h-4 rounded border-slate-300 text-rose-600 focus:ring-rose-500" />
              <label htmlFor="route_is_deleted" className="text-sm font-medium text-rose-700">Mark as deleted</label>
            </div>
          )}

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
