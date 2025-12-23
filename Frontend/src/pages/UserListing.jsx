import { useState, useEffect } from 'react';
import Modal from '../components/Modal';
import '../styles/UserListing.css'; 
import axios from 'axios';

export default function UserListing() {
  const [users, setUsers] = useState([]);
  const [companies, setCompanies] = useState([]);
  const [loading, setLoading] = useState(true);
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [submitting, setSubmitting] = useState(false);

  const BASE_URL = import.meta.env.VITE_API_BASE_URL;

  // Form State
  const [formData, setFormData] = useState({
    username: '',
    email: '',
    role: 'user',
    company_id: '',
    password: ''
  });

  useEffect(() => {
    fetchUsers();
    fetchCompanies();
  }, []);

  const fetchUsers = async () => {
    setLoading(true);
    try {
      const response = await axios.get(`${BASE_URL}/get_users/`);
      setUsers(response.data.data || []);
    } catch (err) {
      console.error("Error fetching users:", err);
    } finally {
      setLoading(false);
    }
  };

  const fetchCompanies = async () => {
    try {
      const response = await axios.get(`${BASE_URL}/customer-data/`);
      setCompanies(response.data?.data || []); 
    } catch (err) {
      console.error("Error fetching companies for dropdown:", err);
      setCompanies([]);
    }
  };

  // ==================== HELPER FUNCTION TO MAP COMPANY ID TO NAME ====================
  const getCompanyNameById = (companyId) => {
    // If companyId is null, undefined, or empty, return N/A
    if (!companyId) {
      return 'N/A';
    }

    // Find the company in the companies array
    const company = companies.find(comp => comp.id === companyId);
    
    // If company found, return name, otherwise return N/A
    return company ? company.company_name : 'N/A';
  };

  const handleInputChange = (e) => {
    const { name, value } = e.target;
    setFormData(prev => ({ ...prev, [name]: value }));
  };

  const handleAddUser = async (e) => {
    e.preventDefault();
    setSubmitting(true);
    try {
      await axios.post(`${BASE_URL}/create_user/`, formData);
      window.alert('User created successfully!');
      setIsModalOpen(false);
      setFormData({ username: '', email: '', role: 'user', company_id: '', password: '' });
      fetchUsers();
    } catch (err) {
      window.alert(err.response?.data?.message || 'Failed to create user');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <div className="user-list-page">
      <div className="user-list-header">
        <h1>User Management</h1>
        <button className="btn-add" onClick={() => setIsModalOpen(true)}>
          + Add New User
        </button>
      </div>

      <div className="table-container">
        <table className="data-table">
          <thead>
            <tr>
              <th>ID</th>
              <th>Username</th>
              <th>Email</th>
              <th>Role</th>
              <th>Company</th>
              <th>Joined Date</th>
            </tr>
          </thead>
          <tbody>
            {loading ? (
              <tr><td colSpan="6" className="text-center">Loading...</td></tr>
            ) : users.length === 0 ? (
              <tr><td colSpan="6" className="text-center">No users found.</td></tr>
            ) : (
              users.map((user) => (
                <tr key={user.id}>
                  <td>{user.id}</td>
                  <td>{user.username}</td>
                  <td>{user.email}</td>
                  <td>
                    <span className={`role-badge ${user.role}`}>
                      {user.role}
                    </span>
                  </td>
                  <td>{getCompanyNameById(user.company)}</td>
                  <td>{user.date_joined ? new Date(user.date_joined).toLocaleDateString() : '-'}</td>
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>

      {/* Add User Modal */}
      <Modal isOpen={isModalOpen} onClose={() => setIsModalOpen(false)} title="Create User Account">
        <form onSubmit={handleAddUser} className="modal-form">
          <div className="form-group">
            <label>Username</label>
            <input type="text" name="username" value={formData.username} onChange={handleInputChange} required />
          </div>
          
          <div className="form-group">
            <label>Email Address</label>
            <input type="email" name="email" value={formData.email} onChange={handleInputChange} required />
          </div>

          <div className="form-group">
            <label>Role</label>
            <select name="role" value={formData.role} onChange={handleInputChange} required>
              <option value="user">User</option>
              <option value="admin">Admin</option>
            </select>
          </div>

          <div className="form-group">
            <label>Assign Company</label>
            <select name="company_id" value={formData.company_id} onChange={handleInputChange} required>
              <option value="">-- Select Company --</option>
              {companies.map(company => (
                <option key={company.id} value={company.id}>
                  {company.company_name}
                </option>
              ))}
            </select>
          </div>

          <div className="form-group">
            <label>Password</label>
            <input type="password" name="password" value={formData.password} onChange={handleInputChange} required />
          </div>

          <div className="form-actions">
            <button type="button" className="btn-cancel" onClick={() => setIsModalOpen(false)}>Cancel</button>
            <button type="submit" className="btn-submit" disabled={submitting}>
              {submitting ? 'Creating...' : 'Create User'}
            </button>
          </div>
        </form>
      </Modal>
    </div>
  );
}