import '../styles/AdminHome.css';

const AdminHome = () => {
  const storedUser = localStorage.getItem("user") ? JSON.parse(localStorage.getItem("user")) : null;
  const username = storedUser?.username || "User";
  return (
    <div className="adminhome">
      <div className="adminhome__card">
        <div className="adminhome__header">
          <h1 className="adminhome__title">Welcome, {username}</h1>
        </div>
      </div>
    </div>
  );
};

export default AdminHome;