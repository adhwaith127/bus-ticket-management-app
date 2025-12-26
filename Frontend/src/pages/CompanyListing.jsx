import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import '../styles/CompanyListing.css';
import api, { BASE_URL } from '../assets/js/axiosConfig';

export default function CompanyListing() {
  const [companies, setCompanies] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [submitting, setSubmitting] = useState(false);
  
  // Modal Mode Management
  const [modalMode, setModalMode] = useState('create'); // 'create', 'view', 'edit'
  const [editingCompany, setEditingCompany] = useState(null);

  // License State
  const [registeringLicense, setRegisteringLicense] = useState({});
  // const [validatingLicense, setValidatingLicense] = useState({});

  // Form State
  const [formData, setFormData] = useState({
    company_name: '',
    company_email: '',
    gst_number: '',
    contact_person: '',
    contact_number: '',
    address: '',
    address_2: '',
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

  const resetFormData = () => {
    setFormData({
      company_name: '',
      company_email: '',
      gst_number: '',
      contact_person: '',
      contact_number: '',
      address: '',
      address_2: '',
      city: '',
      state: '',
      zip_code: '',
      number_of_licence: 1
    });
  };

  const populateFormData = (company) => {
    setFormData({
      company_name: company.company_name || '',
      company_email: company.company_email || '',
      gst_number: company.gst_number || '',
      contact_person: company.contact_person || '',
      contact_number: company.contact_number || '',
      address: company.address || '',
      address_2: company.address_2 || '',
      city: company.city || '',
      state: company.state || '',
      zip_code: company.zip_code || '',
      number_of_licence: company.number_of_licence || 1
    });
  };

  // ==================== MODAL HANDLERS ====================
  
  const openCreateModal = () => {
    setModalMode('create');
    setEditingCompany(null);
    resetFormData();
    setIsModalOpen(true);
  };

  const openViewModal = (company) => {
    setModalMode('view');
    setEditingCompany(company);
    populateFormData(company);
    setIsModalOpen(true);
  };

  const openEditModal = (company) => {
    setModalMode('edit');
    setEditingCompany(company);
    populateFormData(company);
    setIsModalOpen(true);
  };

  const closeModal = () => {
    setIsModalOpen(false);
    setEditingCompany(null);
    setModalMode('create');
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setSubmitting(true);

    try {
      let response;
      
      if (modalMode === 'edit') {
        response = await api.put(
          `${BASE_URL}/update-company-details/${editingCompany.id}/`,
          formData
        );
      } else if (modalMode === 'create') {
        response = await api.post(
          `${BASE_URL}/create-company/`,
          formData
        );
      }

      if (response?.status === 200 || response?.status === 201) {
        window.alert(response.data.message || 'Operation successful!');
        setIsModalOpen(false);
        resetFormData();
        fetchCompanies();
      }

    } catch (err) {
      if (!err.response) {
        window.alert('Server unreachable. Please try again later.');
        return;
      }

      const { status, data } = err.response;

      if (status === 400 && data.errors) {
        const firstError = Object.values(data.errors)[0]?.[0];
        window.alert(firstError || data.message);
      } else {
        window.alert(data?.message || 'Something went wrong');
      }

    } finally {
      setSubmitting(false);
    }
  };

  // ==================== LICENSE REGISTRATION HANDLER ====================
  
  const handleRegisterLicense = async (companyId) => {
    setRegisteringLicense(prev => ({ ...prev, [companyId]: true }));

    try {
      const response = await api.post(
        `${BASE_URL}/register-company-license/${companyId}/`
      );

      if (response.status === 200) {
        window.alert(response.data.message || 'Company registered successfully!');
        fetchCompanies();
      }
    } catch (err) {
      console.error("License registration error:", err);
      
      if (!err.response) {
        window.alert('Server unreachable. Please try again later.');
        return;
      }

      const { data } = err.response;
      window.alert(data?.message || data?.error || 'License registration failed');
    } finally {
      setRegisteringLicense(prev => ({ ...prev, [companyId]: false }));
    }
  };
 
  const handleValidateLicense = async (companyId) => {
    try {
      const response = await api.post(
        `${BASE_URL}/validate-company-license/${companyId}/`
      );

      if (response.status === 200) {
        window.alert(
          response.data.message || 
          'License validation started! This may take up to 2 minutes. Refresh the page to see updated status.'
        );
        fetchCompanies(); // Refresh to show "Validating" status
      }
    } catch (err) {
      console.error("License validation error:", err);
      
      if (!err.response) {
        window.alert('Server unreachable. Please try again later.');
        return;
      }

      const { data } = err.response;
      window.alert(data?.message || data?.error || 'License validation failed');
    }
  };

  
  const getModalTitle = () => {
    if (modalMode === 'view') return 'Company Details';
    if (modalMode === 'edit') return 'Edit Company';
    return 'Register Company';
  };

  const getStatusBadgeClass = (status) => {
    switch (status) {
      case 'Approve':
        return 'status-approved';
      case 'Pending':
        return 'status-pending';
      case 'Validating':
        return 'status-validating';
      case 'Expired':
        return 'status-expired';
      case 'Block':
        return 'status-blocked';
      default:
        return 'status-pending';
    }
  };

  const getStatusLabel = (status) => {
    switch (status) {
      case 'Approve':
        return 'Approved';
      case 'Pending':
        return 'Pending';
      case 'Validating':
        return 'Validating...';
      case 'Expired':
        return 'Expired';
      case 'Block':
        return 'Blocked';
      default:
        return 'Pending';
    }
  };

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
              <th>License</th>
              <th>Actions</th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan="7" className="text-center">Loading...</td></tr>
            ) : companies.length === 0 ? (
              <tr><td colSpan="7" className="text-center">No companies found.</td></tr>
            ) : (
              companies.map((company) => {
                const isPending = company.authentication_status === 'Pending';
                const isValidating = company.authentication_status === 'Validating'; // ← Changed: read from DB
                const hasCompanyId = company.company_id !== null && company.company_id !== undefined;
                const isRegistering = registeringLicense[company.id];
                
                return (
                  <tr key={company.id}>
                    <td>{company.id}</td>
                    <td>{company.company_name}</td>
                    <td>{company.company_email}</td>
                    <td>{company.contact_person}</td>
                    <td>
                      <span className={`status-badge ${getStatusBadgeClass(company.authentication_status)}`}>
                        {getStatusLabel(company.authentication_status)}
                      </span>
                    </td>
                    <td>
                      {!hasCompanyId ? (
                        <button 
                          className="btn-register" 
                          onClick={() => handleRegisterLicense(company.id)}
                          disabled={isRegistering}
                          title="Register with License Server"
                        >
                          {isRegistering ? 'Registering...' : 'Register Company'}
                        </button>
                      ) : isPending ? (
                      <button 
                        className="btn-validate" 
                        onClick={() => handleValidateLicense(company.id)}
                        title="Validate License"
                      >
                        Validate License
                      </button>
                    ) : isValidating ? (
                      <span className="license-status status-validating" title="Validation in progress">
                        ⏳ Validating...
                      </span>
                    ) : (
                      <span className={`license-status ${getStatusBadgeClass(company.authentication_status)}`}>
                        {getStatusLabel(company.authentication_status)}
                      </span>
                    )}
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
                );
              })
            )}
          </tbody>
        </table>
      </div>

      {/* Modal for Create/View/Edit */}
      <Modal isOpen={isModalOpen} onClose={closeModal} title={getModalTitle()}>
        <div className="modal-form">
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
              <label>GST Number</label>
              <input 
                type="text" 
                name="gst_number" 
                value={formData.gst_number} 
                onChange={handleInputChange} 
                placeholder="Optional"
                readOnly={isReadOnly}
              />
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

          <div className="form-group">
            <label>Address 2</label>
            <textarea 
              name="address_2" 
              value={formData.address_2} 
              onChange={handleInputChange} 
              rows="2"
              placeholder="Optional - Additional address details"
              readOnly={isReadOnly}
            ></textarea>
          </div>

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

          <div className="form-actions">
            <button type="button" className="btn-cancel" onClick={closeModal}>
              {modalMode === 'view' ? 'Close' : 'Cancel'}
            </button>
            {modalMode !== 'view' && (
              <button type="button" className="btn-submit" onClick={handleSubmit} disabled={submitting}>
                {submitting ? 'Saving...' : modalMode === 'edit' ? 'Update Company' : 'Save Company'}
              </button>
            )}
          </div>
        </div>
      </Modal>
    </div>
  );
}