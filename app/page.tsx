"use client"

import React, { useState } from 'react';
import { useRouter } from 'next/navigation';

const LandingPage: React.FC = () => {
  const [isVisible, setIsVisible] = useState<boolean>(true);
  const router = useRouter();
  
  const handleGetStarted = (): void => {
    setIsVisible(false); // Hide the intro section
    setTimeout(() => {
      router.push('/loginpage');
    }, 1000); // This should match your CSS transition duration
  };

  return (
    <div 
      className={`min-h-screen bg-gradient-to-br from-blue-600 via-blue-700 to-blue-800 flex items-center justify-center p-4 ${isVisible ? 'opacity-100' : 'opacity-0 transition-opacity duration-1000'}`}
    >
      {/* Left Column - Text Content */}
      <div className="w-full md:w-1/2 mb-10 md:mb-0 pr-0 md:pr-8">
        <h1 className="text-4xl md:text-5xl font-bold text-white mb-6">
          TrackOpz: Supply Chain Management System
        </h1>
        
        <p className="text-white mb-8">
          TrackOpz is a comprehensive supply chain management system designed to streamline operations, enhance visibility, and improve efficiency across your supply chain. From inventory management to order tracking, TrackOpz provides the tools you need to optimize your logistics and drive business success.
        </p>
        
        <button
          className="bg-white text-blue-700 px-8 py-3 rounded-full shadow hover:bg-blue-50 transition-colors"
          onClick={handleGetStarted}
        >
          Go to Login
        </button>
      </div>
      <footer className="absolute bottom-4 left-0 w-full text-white text-xs text-center">
        © 2025 All rights reserved | TrackOpz
      </footer>
    </div>
  );
}

export default LandingPage;

