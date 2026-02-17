import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function EmployeeListing() {

  // ── Section 1: State ────────────────────────────────────────────────────────
  const [employees, setEmployees]     = useState([]);
  const [empTypes, setEmpTypes]       = useState([]);   // for the dropdown in the form
  const [loading, setLoading]         = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode]     = useState('create');
  const [submitting, setSubmitting]   = useState(false);
  const [editingItem, setEditingItem] = useState(null);
  const [showDeleted, setShowDeleted] = useState(false);

  const emptyForm = { employee_code: '', employee_name: '', emp_type: '', phone_no: '', password: '', is_deleted: false };
  const [formData, setFormData] = useState(emptyForm);

  // ── Section 2: Fetch on mount ────────────────────────────────────────────────
  // We fetch employee types ONCE (they don't change while user is on the page).
  // We re-fetch employees whenever showDeleted toggle changes.
  useEffect(() => {
    fetchEmpTypes();
  }, []);

  useEffect(() => {
    fetchEmployees();
  }, [showDeleted]);

  // ── Section 3: API calls ─────────────────────────────────────────────────────
  const fetchEmployees = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/employees/`, {
        params: { show_deleted: showDeleted }
      });
      setEmployees(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching employees:', err);
      setEmployees([]);
    } finally {
      setLoading(false);
    }
  };

  // Fetches the lightweight dropdown list of employee types
  const fetchEmpTypes = async () => {
    try {
      const res = await api.get(`${BASE_URL}/masterdata/dropdowns/employee-types/`);
      setEmpTypes(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching employee types:', err);
    }
  };

  const handleSubmit = async () => {
    setSubmitting(true);
    try {
      let response;
      if (modalMode === 'edit') {
        response = await api.put(`${BASE_URL}/masterdata/employees/update/${editingItem.id}/`, formData);
      } else {
        response = await api.post(`${BASE_URL}/masterdata/employees/create/`, formData);
      }
      if (response?.status === 200 || response?.status === 201) {
        window.alert(response.data.message || 'Success');
        setIsModalOpen(false);
        setFormData(emptyForm);
        fetchEmployees();
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

  // ── Section 4: Modal helpers ──────────────────────────────────────────────────
  const openCreateModal = () => {
    setFormData(emptyForm);
    setEditingItem(null);
    setModalMode('create');
    setIsModalOpen(true);
  };

  const openViewModal = (item) => {
    // emp_type comes back as an ID from the list — keep it consistent
    setFormData({ ...item, emp_type: item.emp_type });
    setEditingItem(item);
    setModalMode('view');
    setIsModalOpen(true);
  };

  const openEditModal = (item) => {
    setFormData({ ...item, emp_type: item.emp_type });
    setEditingItem(item);
    setModalMode('edit');
    setIsModalOpen(true);
  };

  const handleInputChange = (e) => {
    const { name, value, type, checked } = e.target;
    setFormData(prev => ({ ...prev, [name]: type === 'checkbox' ? checked : value }));
  };

  const isReadOnly    = modalMode === 'view';
  const getModalTitle = () => ({ view: 'Employee Details', edit: 'Edit Employee', create: 'Create Employee' }[modalMode]);

  // ── Section 5: Render ─────────────────────────────────────────────────────────
  return (
    <div className="p-6 md:p-10 min-h-screen bg-slate-50 animate-fade-in">

      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
        <div>
          <h1 className="text-3xl font-bold text-slate-800 tracking-tight">Employees</h1>
          <p className="text-slate-500 mt-1">Manage drivers, conductors, and other staff</p>
        </div>
        <div className="flex items-center gap-3">
          <label className="flex items-center gap-2 text-sm text-slate-600 cursor-pointer">
            <input type="checkbox" checked={showDeleted} onChange={() => setShowDeleted(p => !p)} className="w-4 h-4 rounded border-slate-300" />
            Show deleted
          </label>
          <button onClick={openCreateModal} className="flex items-center justify-center bg-slate-800 hover:bg-slate-700 text-white px-5 py-2.5 rounded-xl transition-all shadow-lg hover:shadow-xl transform hover:-translate-y-0.5">
            <span className="font-medium">+ Create Employee</span>
          </button>
        </div>
      </div>

      {/* Table */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50/50 border-b border-slate-200">
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">ID</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Code</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Name</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Type</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Phone</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Status</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr><td colSpan="7" className="px-6 py-8 text-center text-slate-500">Loading...</td></tr>
              ) : employees.length === 0 ? (
                <tr><td colSpan="7" className="px-6 py-8 text-center text-slate-500">No employees found.</td></tr>
              ) : employees.map(item => (
                <tr key={item.id} className="hover:bg-slate-50/80 transition-colors">
                  <td className="px-6 py-4 text-sm text-slate-500 font-mono">#{item.id}</td>
                  <td className="px-6 py-4 text-sm text-slate-800 font-medium">{item.employee_code}</td>
                  <td className="px-6 py-4 text-sm text-slate-800">{item.employee_name}</td>
                  <td className="px-6 py-4 text-sm text-slate-600">{item.emp_type_name || '—'}</td>
                  <td className="px-6 py-4 text-sm text-slate-600">{item.phone_no || '—'}</td>
                  <td className="px-6 py-4">
                    <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium border ${item.is_deleted ? 'bg-red-100 text-red-700 border-red-200' : 'bg-emerald-100 text-emerald-700 border-emerald-200'}`}>
                      {item.is_deleted ? 'Deleted' : 'Active'}
                    </span>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <div className="flex justify-end items-center space-x-2">
                      <button onClick={() => openViewModal(item)} className="p-1.5 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-md transition-colors">View</button>
                      <button onClick={() => openEditModal(item)} className="p-1.5 text-slate-500 hover:text-blue-600 hover:bg-blue-50 rounded-md transition-colors">Edit</button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Modal */}
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={getModalTitle()}>
        <div className="space-y-4">

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Employee Code *</label>
              <input
                type="text" name="employee_code" value={formData.employee_code}
                onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50"
              />
            </div>

            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Employee Name *</label>
              <input
                type="text" name="employee_name" value={formData.employee_name}
                onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50"
              />
            </div>
          </div>

          <div className="space-y-1">
            <label className="text-sm font-medium text-slate-700">Employee Type *</label>
            {isReadOnly ? (
              // In view mode show the name text, not a dropdown
              <input type="text" value={formData.emp_type_name || '—'} readOnly className="w-full px-3 py-2 border border-slate-300 rounded-lg bg-slate-50" />
            ) : (
              <select
                name="emp_type" value={formData.emp_type}
                onChange={handleInputChange}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 bg-white"
              >
                <option value="">-- Select Type --</option>
                {empTypes.map(t => (
                  <option key={t.id} value={t.id}>{t.emp_type_name} ({t.emp_type_code})</option>
                ))}
              </select>
            )}
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Phone Number</label>
              <input
                type="text" name="phone_no" value={formData.phone_no || ''}
                onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50"
              />
            </div>

            <div className="space-y-1">
              <label className="text-sm font-medium text-slate-700">Device PIN</label>
              <input
                type="text" name="password" value={formData.password || ''}
                onChange={handleInputChange} readOnly={isReadOnly}
                className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 read-only:bg-slate-50"
              />
            </div>
          </div>

          {/* Soft delete toggle — only on edit */}
          {modalMode === 'edit' && (
            <div className="flex items-center space-x-3 p-3 bg-red-50 rounded-lg border border-red-100">
              <input
                type="checkbox" name="is_deleted" id="emp_is_deleted"
                checked={formData.is_deleted || false} onChange={handleInputChange}
                className="w-4 h-4 rounded border-slate-300"
              />
              <label htmlFor="emp_is_deleted" className="text-sm font-medium text-red-700">Mark as deleted</label>
            </div>
          )}

          <div className="flex items-center justify-end space-x-3 pt-6 border-t border-slate-100 mt-6">
            <button type="button" onClick={() => setIsModalOpen(false)} className="px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-lg hover:bg-slate-50">
              {isReadOnly ? 'Close' : 'Cancel'}
            </button>
            {!isReadOnly && (
              <button type="button" onClick={handleSubmit} disabled={submitting} className="px-4 py-2 text-sm font-medium text-white bg-slate-800 rounded-lg hover:bg-slate-700 shadow-md disabled:opacity-50">
                {submitting ? 'Saving...' : modalMode === 'edit' ? 'Update' : 'Save'}
              </button>
            )}
          </div>
        </div>
      </Modal>

    </div>
  );
}
