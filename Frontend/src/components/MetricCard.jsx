export default function MetricCard({ title, value, iconClass, color, loading }) {
  return (
    <div className={`metric-card ${loading ? 'loading' : ''}`}>
      <div className="card-content">
        <div className="card-text">
          <h3 className="card-title">{title}</h3>
          <span className="card-value" style={{ color: color }}>
            {value}
          </span>
        </div>
        <div 
          className="icon-container" 
          style={{ backgroundColor: color }}
        >
          <i className={iconClass}></i>
        </div>
      </div>
    </div>
  );
};