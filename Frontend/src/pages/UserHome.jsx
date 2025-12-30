import "../styles/UserHome.css";

export default function UserHome() {
  const storedUser = localStorage.getItem("user") ? JSON.parse(localStorage.getItem("user")) : null;
  const username = storedUser?.username || "User";

  return (
    <div className="user-home-page">
      <header className="user-home-header">
        <h1>Welcome, {username}</h1>
      </header>
    </div>
  );
}