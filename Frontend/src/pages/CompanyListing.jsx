import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
// import api, { BASE_URL } from '../assets/js/axiosConfig';
import '../styles/CompanyListing.css';
import axios from 'axios';


export default function CompanyListing() {
  const [companies, setCompanies] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  const BASE_URL = import.meta.env.VITE_API_BASE_URL;

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
    const response = await axios.get(`${BASE_URL}/customer-data/`);
    
    // Access response.data.data because backend wraps the array in a 'data' key
    // use optional chaining and a fallback to an empty array
    const companyData = response.data?.data || [];
    
    setCompanies(companyData);
  } catch (err) {
    console.error("Error fetching companies:", err);
    setCompanies([]); // Ensure state remains an array on error
  } finally {
    setLoading(false);
  }
};

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleRegister = async (e) => {
    e.preventDefault();
    setSubmitting(true);

    try {
      const response = await axios.post(
        `${BASE_URL}/create-company/`,
        formData
      );

      if (response.status === 201) {
        window.alert(response.data.message);

        setIsModalOpen(false);
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

        fetchCompanies();
      }

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



  return (
    <div className="company-list-page">
      <div className="company-list-header">
        <h1>Company Management</h1>
        <button className="btn-add" onClick={() => setIsModalOpen(true)}>
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
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan="5" className="text-center">Loading...</td></tr>
            ) : companies.length === 0 ? (
              <tr><td colSpan="5" className="text-center">No companies found.</td></tr>
            ) : (
              companies.map((company) => (
                <tr key={company.id}>
                  <td>{company.id}</td>
                  <td>{company.company_name}</td>
                  <td>{company.company_email}</td>
                  <td>{company.contact_person}</td>
                  <td>
                    <span className={`status-badge ${company.verification_status}`}>
                      {company.verification_status}
                    </span>
                  </td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Register Modal */}
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title="Register Company">
        <form onSubmit={handleRegister} className="modal-form">
          <div className="form-row">
            <div className="form-group">
              <label>Company Name *</label>
              <input type="text" name="company_name" value={formData.company_name} onChange={handleInputChange} required />
            </div>
            <div className="form-group">
              <label>Email *</label>
              <input type="email" name="company_email" value={formData.company_email} onChange={handleInputChange} required />
            </div>
          </div>
          
          <div className="form-row">
            <div className="form-group">
              <label>Contact Person</label>
              <input type="text" name="contact_person" value={formData.contact_person} onChange={handleInputChange} required />
            </div>
            <div className="form-group">
              <label>Contact Number</label>
              <input type="text" name="contact_number" value={formData.contact_number} onChange={handleInputChange} required />
            </div>
          </div>

          <div className="form-group">
            <label>Address</label>
            <textarea name="address" value={formData.address} onChange={handleInputChange} required rows="3"></textarea>
          </div>

          <div className="form-row">
            <div className="form-group">
              <label>City</label>
              <input type="text" name="city" value={formData.city} onChange={handleInputChange} required />
            </div>
            <div className="form-group">
              <label>State</label>
              <input type="text" name="state" value={formData.state} onChange={handleInputChange} required />
            </div>
            <div className="form-group">
              <label>Zip Code</label>
              <input type="text" name="zip_code" value={formData.zip_code} onChange={handleInputChange} required />
            </div>
          </div>

          <div className="form-group">
            <label>Number of Licenses</label>
            <input type="number" name="number_of_licence" value={formData.number_of_licence} onChange={handleInputChange} min="1" required />
          </div>

          <div className="form-actions">
            <button type="button" className="btn-cancel" onClick={() => setIsModalOpen(false)}>Cancel</button>
            <button type="submit" className="btn-submit" disabled={submitting}>
              {submitting ? 'Saving...' : 'Save Company'}
            </button>
          </div>
        </form>
      </Modal>
    </div>
  );
}