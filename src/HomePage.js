import React from 'react';
import { useNavigate } from 'react-router-dom';
import "./HomePage.css";

const HomePage = () => {
  const navigate = useNavigate();

  const navigateToApp = () => {
    navigate('/app');
  };

  return (
    <div className="container">
      <header>
        <h1>Zxcel</h1>
      </header>

      <div className="center">
        <div className="button-wrapper">
          <button className="btn btn-one" onClick={navigateToApp}>
            <span>Начать работу</span>
            <svg width="360" height="90">
              <rect x="0" y="0" width="360" height="90" />
            </svg>
          </button>
        </div>
      </div>
    </div>
  );
};

export default HomePage;