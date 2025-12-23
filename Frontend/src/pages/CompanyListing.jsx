import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import '../styles/CompanyListing.css';
// import axios from 'axios';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function CompanyListing() {
  const [companies, setCompanies] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  
  // ==================== NEW STATE FOR VIEW/EDIT ====================
  const [modalMode, setModalMode] = useState('create'); // 'create', 'view', 'edit'
  const [editingCompany, setEditingCompany] = useState(null);

  // Form State
  const [formData, setFormData] = useState({
    company_name: '',
    company_email: '',
    contact_person: '',
    contact_number: '',
    address: '',
    city: '',
    state: '',
    zip_code: '',
    number_of_licence: 1
  });

  useEffect(() => {
    fetchCompanies();
  }, []);

  const fetchCompanies = async () => {
    setLoading(true);
    try {
      const response = await api.get(`${BASE_URL}/customer-data/`);
      const companyData = response.data?.data || [];
      setCompanies(companyData);
    } catch (err) {
      console.error("Error fetching companies:", err);
      setCompanies([]);
    } finally {
      setLoading(false);
    }
  };

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  // ==================== OPEN MODAL FOR CREATE ====================
  const openCreateModal = () => {
    setModalMode('create');
    setEditingCompany(null);
    setFormData({
      company_name: '',
      company_email: '',
      contact_person: '',
      contact_number: '',
      address: '',
      city: '',
      state: '',
      zip_code: '',
      number_of_licence: 1
    });
    setIsModalOpen(true);
  };

  // ==================== OPEN MODAL FOR VIEW ====================
  const openViewModal = (company) => {
    setModalMode('view');
    setEditingCompany(company);
    setFormData({
      company_name: company.company_name || '',
      company_email: company.company_email || '',
      contact_person: company.contact_person || '',
      contact_number: company.contact_number || '',
      address: company.address || '',
      city: company.city || '',
      state: company.state || '',
      zip_code: company.zip_code || '',
      number_of_licence: company.number_of_licence || 1
    });
    setIsModalOpen(true);
  };

  // ==================== OPEN MODAL FOR EDIT ====================
  const openEditModal = (company) => {
    setModalMode('edit');
    setEditingCompany(company);
    setFormData({
      company_name: company.company_name || '',
      company_email: company.company_email || '',
      contact_person: company.contact_person || '',
      contact_number: company.contact_number || '',
      address: company.address || '',
      city: company.city || '',
      state: company.state || '',
      zip_code: company.zip_code || '',
      number_of_licence: company.number_of_licence || 1
    });
    setIsModalOpen(true);
  };

  // ==================== HANDLE CREATE/UPDATE SUBMIT ====================
  const handleSubmit = async (e) => {
    e.preventDefault();
    setSubmitting(true);

    try {
      if (modalMode === 'edit') {
        // UPDATE COMPANY
        const response = await api.put(
          `${BASE_URL}/update-company-details/${editingCompany.id}/`,
          formData
        );

        if (response.status === 200) {
          window.alert(response.data.message || 'Company updated successfully!');
          setIsModalOpen(false);
          fetchCompanies();
        }
      } else if (modalMode === 'create') {
        // CREATE COMPANY
        const response = await api.post(
          `${BASE_URL}/create-company/`,
          formData
        );

        if (response.status === 201) {
          window.alert(response.data.message);
          setIsModalOpen(false);
          fetchCompanies();
        }
      }

      // Reset form
      setFormData({
        company_name: '',
        company_email: '',
        contact_person: '',
        contact_number: '',
        address: '',
        city: '',
        state: '',
        zip_code: '',
        number_of_licence: 1
      });

    } catch (err) {
      if (!err.response) {
        window.alert('Server unreachable. Please try again later.');
        return;
      }

      const { status, data } = err.response;

      if (status === 400) {
        if (data.errors) {
          const firstError = Object.values(data.errors)[0]?.[0];
          window.alert(firstError || data.message);
        } else {
          window.alert(data.message || 'Invalid input');
        }
        return;
      }

      window.alert(data?.message || 'Something went wrong');

    } finally {
      setSubmitting(false);
    }
  };

  // ==================== CLOSE MODAL ====================
  const closeModal = () => {
    setIsModalOpen(false);
    setEditingCompany(null);
    setModalMode('create');
  };

  // ==================== GET MODAL TITLE ====================
  const getModalTitle = () => {
    if (modalMode === 'view') return 'Company Details';
    if (modalMode === 'edit') return 'Edit Company';
    return 'Register Company';
  };

  // Check if field should be read-only
  const isReadOnly = modalMode === 'view';

  return (
    <div className="company-list-page">
      <div className="company-list-header">
        <h1>Company Management</h1>
        <button className="btn-add" onClick={openCreateModal}>
          + Register New Company
        </button>
      </div>

      <div className="table-container">
        <table className="data-table">
          <thead>
            <tr>
              <th>ID</th>
              <th>Company Name</th>
              <th>Email</th>
              <th>Contact Person</th>
              <th>Status</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan="6" className="text-center">Loading...</td></tr>
            ) : companies.length === 0 ? (
              <tr><td colSpan="6" className="text-center">No companies found.</td></tr>
            ) : (
              companies.map((company) => (
                <tr key={company.id}>
                  <td>{company.id}</td>
                  <td>{company.company_name}</td>
                  <td>{company.company_email}</td>
                  <td>{company.contact_person}</td>
                  <td>
                    <span className={`status-badge ${company.verification_status}`}>
                      {company.verification_status || 'Active'}
                    </span>
                  </td>
                  <td>
                    <div className="action-buttons">
                      <button 
                        className="btn-view" 
                        onClick={() => openViewModal(company)}
                        title="View Details"
                      >
                        View
                      </button>
                      <button 
                        className="btn-edit" 
                        onClick={() => openEditModal(company)}
                        title="Edit Company"
                      >
                        Edit
                      </button>
                    </div>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Modal for Create/View/Edit */}
      <Modal isOpen={isModalOpen} onClose={closeModal} title={getModalTitle()}>
        <form onSubmit={handleSubmit} className="modal-form">
          <div className="form-row">
            <div className="form-group">
              <label>Company Name *</label>
              <input 
                type="text" 
                name="company_name" 
                value={formData.company_name} 
                onChange={handleInputChange} 
                required 
                readOnly={isReadOnly}
              />
            </div>
            <div className="form-group">
              <label>Email *</label>
              <input 
                type="email" 
                name="company_email" 
                value={formData.company_email} 
                onChange={handleInputChange} 
                required 
                readOnly={isReadOnly}
              />
            </div>
          </div>
          
          <div className="form-row">
            <div className="form-group">
              <label>Contact Person *</label>
              <input 
                type="text" 
                name="contact_person" 
                value={formData.contact_person} 
                onChange={handleInputChange} 
                required 
                readOnly={isReadOnly}
              />
            </div>
            <div className="form-group">
              <label>Contact Number *</label>
              <input 
                type="text" 
                name="contact_number" 
                value={formData.contact_number} 
                onChange={handleInputChange} 
                required 
                readOnly={isReadOnly}
              />
            </div>
          </div>

          <div className="form-group">
            <label>Address *</label>
            <textarea 
              name="address" 
              value={formData.address} 
              onChange={handleInputChange} 
              required 
              rows="3"
              readOnly={isReadOnly}
            ></textarea>
          </div>

          {/* FIXED ROW: City 40%, State 35%, Zip 25% */}
          <div className="form-row form-row--location">
            <div className="form-group form-group--city">
              <label>City *</label>
              <input 
                type="text" 
                name="city" 
                value={formData.city} 
                onChange={handleInputChange} 
                required 
                readOnly={isReadOnly}
              />
            </div>
            <div className="form-group form-group--state">
              <label>State *</label>
              <input 
                type="text" 
                name="state" 
                value={formData.state} 
                onChange={handleInputChange} 
                required 
                readOnly={isReadOnly}
              />
            </div>
            <div className="form-group form-group--zip">
              <label>Zip Code *</label>
              <input 
                type="text" 
                name="zip_code" 
                value={formData.zip_code} 
                onChange={handleInputChange} 
                required 
                readOnly={isReadOnly}
              />
            </div>
          </div>

          <div className="form-group">
            <label>Number of Licenses *</label>
            <input 
              type="number" 
              name="number_of_licence" 
              value={formData.number_of_licence} 
              onChange={handleInputChange} 
              min="1" 
              required 
              readOnly={isReadOnly}
            />
          </div>

          <div className="form-actions">
            <button type="button" className="btn-cancel" onClick={closeModal}>
              {modalMode === 'view' ? 'Close' : 'Cancel'}
            </button>
            {modalMode !== 'view' && (
              <button type="submit" className="btn-submit" disabled={submitting}>
                {submitting ? 'Saving...' : modalMode === 'edit' ? 'Update Company' : 'Save Company'}
              </button>
            )}
          </div>
        </form>
      </Modal>
    </div>
  );
}