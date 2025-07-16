'use client';
import React, { useState, useEffect } from 'react';
import { Menu, ChevronDown } from 'lucide-react';
import Sidebar from '../../components/sidebar';

// Type definitions
interface JobData {
  machine: string;
  product: string;
  state: string;
  stage: string;
  quantity: number;
}

interface CustomDropdownProps {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
}

// Custom Dropdown Component
const CustomDropdown: React.FC<CustomDropdownProps> = ({ 
  label, 
  value, 
  options, 
  onChange 
}) => {
  const [isOpen, setIsOpen] = useState<boolean>(false);

  const handleOptionClick = (option: string): void => {
    onChange(option);
    setIsOpen(false);
  };

  const handleToggle = (): void => {
    setIsOpen(!isOpen);
  };

  return (
    <div className="mb-4 md:mb-6">
      <label className="block text-gray-700 font-medium mb-2 text-sm md:text-base">
        {label}
      </label>
      <div className="relative">
        <button
          onClick={handleToggle}
          className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-left flex items-center justify-between hover:border-blue-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
        >
          <span className="text-gray-700 truncate max-w-[calc(100%-30px)]">{value}</span>
          <ChevronDown 
            className={`w-5 h-5 text-gray-400 transition-transform duration-200 ${
              isOpen ? 'rotate-180' : ''
            }`} 
          />
        </button>
        
        {isOpen && (
          <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-gray-300 rounded-lg shadow-lg z-10 max-h-48 overflow-y-auto">
            {options.map((option: string, index: number) => (
              <button
                key={index}
                onClick={() => handleOptionClick(option)}
                className="w-full px-4 py-3 text-left hover:bg-blue-50 text-gray-700 first:rounded-t-lg last:rounded-b-lg transition-colors text-sm md:text-base"
              >
                {option}
              </button>
            ))}
          </div>
        )}
      </div>
    </div>
  );
};

export default function AddJobsForm() {
  const [selectedMachine, setSelectedMachine] = useState<string>('Cutting');
  const [selectedProduct, setSelectedProduct] = useState<string>('');
  const [selectedState, setSelectedState] = useState<string>('ON');
  const [selectedStage, setSelectedStage] = useState<string>('Milling');
  const [quantity, setQuantity] = useState<number>(1);
  const [quantityInput, setQuantityInput] = useState<string>('1');
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [validationMessage, setValidationMessage] = useState<string>('');
  const [isValidating, setIsValidating] = useState<boolean>(false);
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' | 'info' } | null>(null);

  // Auto-hide messages after 5 seconds
  useEffect(() => {
    if (message) {
      const timer = setTimeout(() => {
        setMessage(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [message]);

  const machines: string[] = [
    'Cutting',
    'CNC Turning Soft-1',
    'CNC Turning Soft-2',
    'Gun Drilling',
    'Heat Treatment',
    'CNC Turning Hard-1',
    'CNC Turning hard-2',
    'Grinding',
    'Flat Milling VMC',
    'Dummy Bottom CNC',
    'Profile CNC',
    'PDI',
    'RFD'
  ];

  const states: string[] = [
    'ON',
    'OFF',
  ];




  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  const handleCloseSidebar = (): void => {
    setSidebarOpen(false);
  };

  // Validate if product can be set to OFF status
  const validateProductStatus = async (productName: string, machine: string, state: string): Promise<boolean> => {
    if (state !== 'OFF') {
      setValidationMessage('');
      return true; // No validation needed for ON status
    }

    setIsValidating(true);
    setValidationMessage('Validating product status...');

    try {
      const response = await fetch('/api/validate-product-status', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ productName, machine, quantity })
      });

      const data = await response.json();

      if (data.canSetOff) {
        setValidationMessage(`✅ ${data.reason}`);
        setIsValidating(false);
        return true;
      } else {
        setValidationMessage(`❌ ${data.reason}`);
        setIsValidating(false);
        return false;
      }
    } catch {
      setValidationMessage('❌ Error validating product status');
      setIsValidating(false);
      return false;
    }
  };

  const handleAddJob = async (): Promise<void> => {
    // Validate product status before adding job
    const isValid = await validateProductStatus(selectedProduct, selectedMachine, selectedState);
    
    if (!isValid) {
      return; // Don't proceed if validation fails
    }

    // Validate quantity
    if (quantity <= 0 || quantity > 1000) {
      setMessage({ text: 'Please enter a valid quantity (1-1000)', type: 'error' });
      return;
    }

    const jobData: JobData = {
      machine: selectedMachine,
      product: selectedProduct,
      state: selectedState,
      stage: selectedStage,
      quantity: quantity
    };
    
    try {
      const res = await fetch('/api/jobs', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(jobData)
      });
      if (res.ok) {
        await res.json();
        // Optionally show a success message or reset form
        setSelectedProduct('');
        setSelectedStage('Milling');
        setSelectedState('ON');
        setSelectedMachine('Cutting');
        setQuantity(1);
        setQuantityInput('1');
        setValidationMessage('');
        setMessage({ text: `Successfully added ${quantity} job(s)!`, type: 'success' });
      } else {
        const errorData = await res.json();
        setMessage({ text: `Failed to add job: ${errorData.error || 'Unknown error'}`, type: 'error' });
      }
    } catch {
      setMessage({ text: 'Failed to add job.', type: 'error' });
    }
  };

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Sidebar Component */}
      <Sidebar 
        isOpen={sidebarOpen} 
        onClose={handleCloseSidebar} 
        username={null}
      />

      {/* Header */}
      <header className="bg-blue-700 shadow-sm px-4 py-4 flex items-center justify-between sticky top-0 z-20">
        <button 
          onClick={handleMenuClick}
          className="p-2 hover:bg-blue-600 rounded-lg transition-colors"
        >
          <Menu className="w-6 h-6 text-white" />
        </button>
        
        <h1 className="text-lg md:text-xl font-semibold text-white text-center flex-grow">Add Jobs</h1>
        
        <div className="w-8 h-8 md:w-10 md:h-10 bg-blue-600 rounded-full flex items-center justify-center">
          <span className="text-white font-medium text-base md:text-lg">A</span>
        </div>
      </header>

      {/* Main Content */}
      <main className="px-4 md:px-6 py-6 flex-grow">
        <div className="w-full max-w-md mx-auto">
          {/* Message Display */}
          {message && (
            <div className={`mb-4 p-3 rounded-lg text-sm ${
              message.type === 'success' 
                ? 'bg-green-100 text-green-700 border border-green-200' 
                : message.type === 'error' 
                ? 'bg-red-100 text-red-700 border border-red-200'
                : 'bg-blue-100 text-blue-700 border border-blue-200'
            }`}>
              {message.text}
            </div>
          )}

          {/* Select Machine/Process */}
          <CustomDropdown
            label="Select Machine/Process"
            value={selectedMachine}
            options={machines}
            onChange={(value) => {
              setSelectedMachine(value);
              setValidationMessage('');
              // Validate if state is OFF
              if (selectedState === 'OFF' && selectedProduct && value) {
                validateProductStatus(selectedProduct, value, selectedState);
              }
            }}
          />

          {/* Product */}
          <div className="mb-4 md:mb-6">
            <label className="block text-gray-700 font-medium mb-2 text-sm md:text-base">
              Product ID
            </label>
            <input
              type="text"
              className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
              value={selectedProduct}
              onChange={e => {
                setSelectedProduct(e.target.value);
                setValidationMessage('');
                // Validate if state is OFF
                if (selectedState === 'OFF' && e.target.value && selectedMachine) {
                  validateProductStatus(e.target.value, selectedMachine, selectedState);
                }
              }}
              placeholder="A-1825"
            />
          </div>

          {/* Quantity */}
          <div className="mb-4 md:mb-6">
            <label className="block text-gray-700 font-medium mb-2 text-sm md:text-base">
              Quantity
            </label>
            <input
              type="number"
              min="1"
              max="1000"
              className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
              value={quantityInput}
              onChange={e => {
                const inputValue = e.target.value;
                setQuantityInput(inputValue);
                
                // Update quantity if it's a valid number
                const newQuantity = parseInt(inputValue);
                if (!isNaN(newQuantity)) {
                  setQuantity(newQuantity);
                }
              }}
              onBlur={e => {
                // Apply constraints when user finishes editing
                const inputValue = parseInt(e.target.value) || 1;
                const constrainedValue = Math.max(1, Math.min(1000, inputValue));
                setQuantity(constrainedValue);
                setQuantityInput(constrainedValue.toString());
                
                // Re-validate if state is OFF
                if (selectedState === 'OFF' && selectedProduct && selectedMachine) {
                  validateProductStatus(selectedProduct, selectedMachine, selectedState);
                }
              }}
              placeholder="1"
            />
            <p className="text-sm text-gray-500 mt-1">
              Enter quantity (1-1000)
            </p>
          </div>


          {/* State */}
          <CustomDropdown
            label="State"
            value={selectedState}
            options={states}
            onChange={(value) => {
              setSelectedState(value);
              // Clear validation message when state changes
              setValidationMessage('');
              // Validate immediately if state is OFF
              if (value === 'OFF' && selectedProduct && selectedMachine) {
                validateProductStatus(selectedProduct, selectedMachine, value);
              }
            }}
          />

          {/* Validation Message */}
          {validationMessage && (
            <div className={`p-3 rounded-lg text-sm ${
              validationMessage.startsWith('✅') 
                ? 'bg-green-100 text-green-700 border border-green-200' 
                : validationMessage.startsWith('❌') 
                ? 'bg-red-100 text-red-700 border border-red-200'
                : 'bg-blue-100 text-blue-700 border border-blue-200'
            }`}>
              {validationMessage}
            </div>
          )}

          {/* Add Button */}
          <div className="flex justify-center mt-6 md:mt-8">
            <button
              onClick={handleAddJob}
              disabled={isValidating || (selectedState === 'OFF' && validationMessage.startsWith('❌'))}
              className={`w-full md:w-auto px-6 md:px-8 py-3 rounded-full font-medium transition-colors shadow-lg ${
                isValidating || (selectedState === 'OFF' && validationMessage.startsWith('❌'))
                  ? 'bg-gray-300 text-gray-500 cursor-not-allowed'
                  : 'bg-white text-blue-700 hover:bg-blue-700 hover:text-white'
              }`}
            >
              {isValidating ? 'Validating...' : 'Add'}
            </button>
          </div>
        </div>
      </main>
    </div>
  );
}