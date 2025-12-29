import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import api, { BASE_URL } from '../assets/js/axiosConfig';
import '../styles/BranchListing.css';

export default function BranchListing() {
  const [branches, setBranches] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [modalMode, setModalMode] = useState('create');
  const [submitting, setSubmitting] = useState(false);
  const [editingBranch, setEditingBranch] = useState(null);

  const [formData, setFormData] = useState({
    branch_code: '',
    branch_name: '',
    address: '',
    city: '',
    state: '',
    zip_code: ''
  });

  useEffect(() => {
    fetchBranches();
  }, []);

  const fetchBranches = async () => {
    setLoading(true);
    try {
      const res = await api.get(`${BASE_URL}/branches/`);
      setBranches(res.data?.data || []);
    } catch (err) {
      console.error("Error fetching branches:", err);
      setBranches([]);
    } finally {
      setLoading(false);
    }
  };

  const resetFormData = () => {
    setFormData({ branch_code:'', branch_name:'', address:'', city:'', state:'', zip_code:'' });
  };

  const openCreateModal = () => {
    resetFormData();
    setEditingBranch(null);
    setModalMode('create');
    setIsModalOpen(true);
  };

  const openViewModal = (branch) => {
    setEditingBranch(branch);
    setFormData(branch);
    setModalMode('view');
    setIsModalOpen(true);
  };

  const openEditModal = (branch) => {
    setEditingBranch(branch);
    setFormData(branch);
    setModalMode('edit');
    setIsModalOpen(true);
  };

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setSubmitting(true);
    try {
      let response;
      if (modalMode === 'edit') {
        response = await api.put(`${BASE_URL}/update-branch-details/${editingBranch.id}/`, formData);
      } else {
        response = await api.post(`${BASE_URL}/create-branch/`, formData);
      }

      if (response?.status === 200 || response?.status === 201) {
        window.alert(response.data.message || 'Success');
        setIsModalOpen(false);
        resetFormData();
        fetchBranches();
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

  const isReadOnly = modalMode === 'view';

  const getStatusBadge = (active) => {
    return active ? 'branch-status-active' : 'branch-status-inactive';
  };

  const getStatusLabel = (active) => {
    return active ? 'Active' : 'Inactive';
  };

  const getModalTitle = () => {
    if (modalMode === 'view') return 'Branch Details';
    if (modalMode === 'edit') return 'Edit Branch';
    return 'Create Branch';
  };

  return (
    <div className="branch-page">
      <div className="branch-header">
        <h1>Branch Management</h1>
        <button className="branch-btn-add" onClick={openCreateModal}>+ Create Branch</button>
      </div>

      <div className="branch-table-container">
        <table className="branch-table">
          <thead>
            <tr>
              <th>ID</th>
              <th>Branch Code</th>
              <th>Name</th>
              <th>Address</th>
              <th>Status</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan="6" className="text-center">Loading...</td></tr>
            ) : branches.length === 0 ? (
              <tr><td colSpan="6" className="text-center">No branches found.</td></tr>
            ) : (
              branches.map(branch => (
                <tr key={branch.id}>
                  <td>{branch.id}</td>
                  <td>{branch.branch_code}</td>
                  <td>{branch.branch_name}</td>
                  <td>{branch.address}, {branch.city}, {branch.state}, {branch.zip_code}</td>
                  <td>
                    <span className={`branch-status-badge ${getStatusBadge(branch.is_active)}`}>
                      {getStatusLabel(branch.is_active)}
                    </span>
                  </td>
                  <td>
                    <div className="branch-actions">
                      <button className="branch-btn-view" onClick={() => openViewModal(branch)}>View</button>
                      <button className="branch-btn-edit" onClick={() => openEditModal(branch)}>Edit</button>
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title={getModalTitle()}>
        <div className="branch-modal-form">
          <div className="branch-form-group">
            <label>Branch Code *</label>
            <input type="text" name="branch_code" value={formData.branch_code} onChange={handleInputChange} readOnly={isReadOnly} />
          </div>
          <div className="branch-form-group">
            <label>Branch Name *</label>
            <input type="text" name="branch_name" value={formData.branch_name} onChange={handleInputChange} readOnly={isReadOnly} />
          </div>
          <div className="branch-form-group">
            <label>Address *</label>
            <textarea name="address" value={formData.address} onChange={handleInputChange} readOnly={isReadOnly}></textarea>
          </div>
          <div className="branch-form-row">
            <div className="branch-form-group">
              <label>City *</label>
              <input type="text" name="city" value={formData.city} onChange={handleInputChange} readOnly={isReadOnly} />
            </div>
            <div className="branch-form-group">
              <label>State *</label>
              <input type="text" name="state" value={formData.state} onChange={handleInputChange} readOnly={isReadOnly} />
            </div>
            <div className="branch-form-group">
              <label>Zip Code *</label>
              <input type="text" name="zip_code" value={formData.zip_code} onChange={handleInputChange} readOnly={isReadOnly} />
            </div>
          </div>

          <div className="branch-form-actions">
            <button className="branch-btn-cancel" onClick={() => setIsModalOpen(false)}>
              {modalMode === 'view' ? 'Close' : 'Cancel'}
            </button>
            {modalMode !== 'view' && (
              <button className="branch-btn-submit" onClick={handleSubmit} disabled={submitting}>
                {submitting ? 'Saving...' : modalMode === 'edit' ? 'Update Branch' : 'Save Branch'}
              </button>
            )}
          </div>
        </div>
      </Modal>
    </div>
  );
}