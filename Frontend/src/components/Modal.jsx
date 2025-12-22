import React from 'react';
import '../styles/Modal.css';

const Modal = ({ isOpen, onClose, title, children }) => {
  if (!isOpen) return null;

  return (
    <div className="custom-modal-overlay" onClick={onClose}>
      <div className="custom-modal-content" onClick={(e) => e.stopPropagation()}>
        <div className="custom-modal-header">
          <h3 className="custom-modal-title">{title}</h3>
          <button className="custom-modal-close" onClick={onClose}>&times;</button>
        </div>
        <div className="custom-modal-body">
          {children}
        </div>
      </div>
    </div>
  );
};

export default Modal;