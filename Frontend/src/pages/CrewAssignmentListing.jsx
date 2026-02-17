import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function CrewAssignmentListing() {

  // ── Section 1: State ────────────────────────────────────────────────────────
  const [assignments, setAssignments] = useState([]);
  const [loading, setLoading]         = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode]     = useState('create');
  const [submitting, setSubmitting]   = useState(false);
  const [editingItem, setEditingItem] = useState(null);

  // Separate dropdown lists for each role.
  // These are fetched once when the component mounts.
  const [drivers, setDrivers]       = useState([]);
  const [conductors, setConductors] = useState([]);
  const [cleaners, setCleaners]     = useState([]);
  const [vehicles, setVehicles]     = useState([]);

  const emptyForm = { driver: '', conductor: '', cleaner: '', vehicle: '' };
  const [formData, setFormData] = useState(emptyForm);

  // ── Section 2: Fetch on mount ────────────────────────────────────────────────
  useEffect(() => {
    fetchAssignments();
    fetchDropdowns();
  }, []);

  // ── Section 3: API calls ─────────────────────────────────────────────────────
  const fetchAssignments = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/masterdata/crew-assignments/`);
      setAssignments(res.data?.data || []);
    } catch (err) {
      console.error('Error fetching crew assignments:', err);
      setAssignments([]);
    } finally {
      setLoading(false);
    }
  };

  // All 4 dropdown lists are fetched in parallel using Promise.all.
  // This is faster than fetching them one by one sequentially.
  // The ?type= query param filters employees by their emp_type_code.
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
      // Send only non-empty values — conductor and cleaner are optional
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

  // ── Section 4: Modal helpers ──────────────────────────────────────────────────
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

  // Helper: render a dropdown OR a read-only text field depending on modalMode
  const renderDropdown = (name, label, options, required = false) => (
    <div className="space-y-1">
      <label className="text-sm font-medium text-slate-700">{label}{required ? ' *' : ' (optional)'}</label>
      {isReadOnly ? (
        // In view mode, show the name of the selected option
        <input type="text"
          value={options.find(o => String(o.id) === String(formData[name]))?.employee_name
              || options.find(o => String(o.id) === String(formData[name]))?.bus_reg_num
              || '—'}
          readOnly className="w-full px-3 py-2 border border-slate-300 rounded-lg bg-slate-50"
        />
      ) : (
        <select name={name} value={formData[name]} onChange={handleInputChange}
          className="w-full px-3 py-2 border border-slate-300 rounded-lg focus:ring-slate-500 bg-white">
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

  // ── Section 5: Render ─────────────────────────────────────────────────────────
  return (
    <div className="p-6 md:p-10 min-h-screen bg-slate-50 animate-fade-in">

      {/* Header */}
      <div className="flex flex-col md:flex-row md:items-center justify-between mb-8 gap-4">
        <div>
          <h1 className="text-3xl font-bold text-slate-800 tracking-tight">Crew Assignments</h1>
          <p className="text-slate-500 mt-1">Assign drivers, conductors and cleaners to vehicles</p>
        </div>
        <button onClick={openCreateModal} className="flex items-center justify-center bg-slate-800 hover:bg-slate-700 text-white px-5 py-2.5 rounded-xl transition-all shadow-lg hover:shadow-xl transform hover:-translate-y-0.5">
          <span className="font-medium">+ Create Assignment</span>
        </button>
      </div>

      {/* Table */}
      <div className="bg-white rounded-2xl shadow-sm border border-slate-200 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-50/50 border-b border-slate-200">
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">ID</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Vehicle</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Driver</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Conductor</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider">Cleaner</th>
                <th className="px-6 py-4 text-xs font-semibold text-slate-500 uppercase tracking-wider text-right">Actions</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {loading ? (
                <tr><td colSpan="6" className="px-6 py-8 text-center text-slate-500">Loading...</td></tr>
              ) : assignments.length === 0 ? (
                <tr><td colSpan="6" className="px-6 py-8 text-center text-slate-500">No crew assignments found.</td></tr>
              ) : assignments.map(item => (
                <tr key={item.id} className="hover:bg-slate-50/80 transition-colors">
                  <td className="px-6 py-4 text-sm text-slate-500 font-mono">#{item.id}</td>
                  <td className="px-6 py-4 text-sm text-slate-800 font-medium">{item.vehicle_reg || '—'}</td>
                  <td className="px-6 py-4 text-sm text-slate-800">{item.driver_name || '—'}</td>
                  <td className="px-6 py-4 text-sm text-slate-600">{item.conductor_name || '—'}</td>
                  <td className="px-6 py-4 text-sm text-slate-600">{item.cleaner_name || '—'}</td>
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

          {renderDropdown('driver',    'Driver',    drivers,    true)}
          {renderDropdown('conductor', 'Conductor', conductors, false)}
          {renderDropdown('cleaner',   'Cleaner',   cleaners,   false)}
          {renderDropdown('vehicle',   'Vehicle',   vehicles,   true)}

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
