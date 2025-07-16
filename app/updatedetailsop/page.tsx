'use client';
import React, { useState, useEffect, useRef } from 'react';
import { Menu, ChevronDown, Check } from 'lucide-react';
import Sidebar from '../../components/sidebar';

// Type definitions
type DispatchStatus = 'Dispatched' | 'Pending' | 'In Transit';
type ProcessStepKey = 'deburring' | 'finalInspect' | 'oiling';

interface ProcountProduct {
  id: number;
  product: {
    id: number;
    name: string;
  };
  count: number;
  status: string;
  machine?: string;
  stage?: string;
  state?: string;
  updatedAt?: string;
}

interface ProcessSteps {
  deburring: boolean;
  finalInspect: boolean;
  oiling: boolean;
}

interface CustomDropdownProps {
  label: string;
  value: string;
  options: string[];
  onChange: (value: string) => void;
}

interface ProcessCheckboxProps {
  label: string;
  checked: boolean;
  onChange: () => void;
}

interface UpdateData {
  id: string | number;
  processSteps: ProcessSteps;
  dispatchStatus: DispatchStatus;
  quantity: number;
}

export default function UpdateDetailsPage(): React.ReactElement {
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);

  const [procountProducts, setProcountProducts] = useState<ProcountProduct[]>([]);
  const [processSteps, setProcessSteps] = useState<ProcessSteps>({
    deburring: true,
    finalInspect: true,
    oiling: true
  });
  const [selectedProductId, setSelectedProductId] = useState<string | number | null>(null);
  const [quantity, setQuantity] = useState<number>(1);
  const [quantityInput, setQuantityInput] = useState<string>('1');

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

  // Fetch procount products from backend
  useEffect(() => {
    async function fetchProcountProducts() {
      const res = await fetch('/api/procount');
      if (res.ok) {
        const data = await res.json();
        console.log('Procount products loaded:', data.productCounts);
        setProcountProducts(data.productCounts);
        // Set the first product as selected if available
        if (data.productCounts.length > 0 && !selectedProductId) {
          setSelectedProductId(data.productCounts[0].id);
        }
      }
    }
    fetchProcountProducts();
  }, [selectedProductId]);



  // Polling for procount products updates
  useEffect(() => {
    const pollInterval = setInterval(async () => {
      try {
        const res = await fetch('/api/procount');
        if (res.ok) {
          const data = await res.json();
          setProcountProducts(data.productCounts);
          // Set the first product as selected if available and no product is currently selected
          if (data.productCounts.length > 0 && !selectedProductId) {
            setSelectedProductId(data.productCounts[0].id);
          }
        }
      } catch (error) {
        console.warn('Failed to fetch procount products:', error);
      }
    }, 5000); // Poll every 5 seconds

    return () => {
      clearInterval(pollInterval);
    };
  }, [selectedProductId]);

  // Real-time updates using Server-Sent Events for procount
  useEffect(() => {
    let eventSource: EventSource | null = null;
    
    const connectToStream = () => {
      try {
        eventSource = new EventSource('/api/procount/stream');
        
        eventSource.onmessage = (event) => {
          try {
            const data = JSON.parse(event.data);
            
            if (data.type === 'productCountUpdated') {
              console.log('Product count updated:', data.productCount);
              // Refresh the procount products list
              fetch('/api/procount')
                .then(res => res.json())
                .then(data => {
                  setProcountProducts(data.productCounts);
                  // Set the first product as selected if available and no product is currently selected
                  if (data.productCounts.length > 0 && !selectedProductId) {
                    setSelectedProductId(data.productCounts[0].id);
                  }
                })
                .catch(error => console.warn('Failed to refresh procount products:', error));
            }
          } catch (error) {
            console.error('Error parsing SSE data:', error);
          }
        };

        eventSource.onopen = () => {
          console.log('SSE connection opened successfully for procount');
        };

        eventSource.onerror = (error) => {
          console.log('SSE connection error (this is normal when page is closed):', error);
          if (eventSource) {
            eventSource.close();
            eventSource = null;
          }
          // Only try to reconnect if the page is still active
          if (!document.hidden) {
            setTimeout(connectToStream, 10000); // Wait 10 seconds before reconnecting
          }
        };

      } catch (error) {
        console.error('Failed to create SSE connection:', error);
        // Try to reconnect after 5 seconds
        setTimeout(connectToStream, 5000);
      }
    };

    connectToStream();

    // Handle page visibility changes
    const handleVisibilityChange = () => {
      if (document.hidden) {
        // Page is hidden, close the connection
        if (eventSource) {
          eventSource.close();
          eventSource = null;
        }
      } else {
        // Page is visible again, try to reconnect
        setTimeout(connectToStream, 2000);
      }
    };

    document.addEventListener('visibilitychange', handleVisibilityChange);

    // Cleanup on unmount
    return () => {
      document.removeEventListener('visibilitychange', handleVisibilityChange);
      if (eventSource) {
        eventSource.close();
      }
    };
  }, [selectedProductId]);

  // Debug selectedProductId changes
  useEffect(() => {
    console.log('Selected product ID changed to:', selectedProductId);
    if (selectedProductId) {
      const selectedProduct = procountProducts.find(p => p.id === selectedProductId);
      console.log('Selected product details:', selectedProduct);
    }
  }, [selectedProductId, procountProducts]);

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };



  const handleProcessStepChange = (step: ProcessStepKey): void => {
    setProcessSteps(prev => ({
      ...prev,
      [step]: !prev[step]
    }));
  };

  const handleUpdate = async (): Promise<void> => {
    if (!selectedProductId) {
      setMessage({ text: 'Please select a product first', type: 'info' });
      return;
    }

    if (quantity <= 0) {
      setMessage({ text: 'Please enter a valid quantity (greater than 0)', type: 'info' });
      return;
    }

    // Validate that the requested quantity doesn't exceed available quantity
    const selectedProduct = procountProducts.find(p => p.id === selectedProductId);
    if (!selectedProduct) {
      setMessage({ text: 'Selected product not found', type: 'error' });
      return;
    }

    if (quantity > selectedProduct.count) {
      setMessage({ text: `Cannot dispatch ${quantity} items. Only ${selectedProduct.count} items are available in procount.`, type: 'error' });
      return;
    }

    const updateData: UpdateData = {
      id: selectedProductId,
      processSteps: processSteps,
      dispatchStatus: 'Pending', // Default to Pending since we removed the dropdown
      quantity: quantity
    };

    try {
      console.log('Sending update request with data:', updateData);
      
      const response = await fetch('/api/operator/update', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
        },
        body: JSON.stringify(updateData),
      });

      console.log('Update response status:', response.status);
      
      if (!response.ok) {
        const errorText = await response.text();
        console.error('Update failed with status:', response.status, 'Error:', errorText);
        setMessage({ text: `❌ Update failed: Server returned ${response.status}. Please try again.`, type: 'error' });
        return;
      }

      const result = await response.json();
      console.log('Update response:', result);

      if (result.success) {
        // Show success message
        const selectedProduct = procountProducts.find(p => p.id === selectedProductId);
        const productName = selectedProduct?.product?.name || 'Unknown Product';
        const successMessage = `✅ Product "${productName}" updated successfully!\n\nQuantity: ${quantity}\nProcess Steps: ${Object.entries(processSteps).filter(([_, completed]) => completed).map(([step, _]) => step).join(', ')}`;
        setMessage({ text: successMessage, type: 'success' });
        console.log('Updated details:', updateData);
        
        // Reset form after successful update
        setProcessSteps({
          deburring: true,
          finalInspect: true,
          oiling: true
        });
        setSelectedProductId(null);
        setQuantity(1);
      } else {
        console.error('Update failed:', result.error);
        setMessage({ text: `❌ Update failed: ${result.error}`, type: 'error' });
      }
    } catch (error) {
      console.error('Error updating product:', error);
      setMessage({ text: 'Failed to update product. Please check your connection and try again.', type: 'error' });
    }
  };

  const CustomDropdown: React.FC<CustomDropdownProps> = ({ 
    label, 
    value, 
    options, 
    onChange 
  }) => {
    const [isOpen, setIsOpen] = useState<boolean>(false);
    const dropdownRef = useRef<HTMLDivElement>(null);

    // Handle click outside dropdown to close it
    useEffect(() => {
      function handleClickOutside(event: MouseEvent) {
        if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
          console.log('Click outside detected, closing dropdown');
          setIsOpen(false);
        }
      }

      if (isOpen) {
        document.addEventListener('mousedown', handleClickOutside);
      }

      return () => {
        document.removeEventListener('mousedown', handleClickOutside);
      };
    }, [isOpen]);

    const handleOptionSelect = (option: string, event: React.MouseEvent) => {
      event.preventDefault();
      event.stopPropagation();
      console.log('Updatedetailsop dropdown: Option selected:', option);
      console.log('Current value:', value);
      console.log('Calling onChange with:', option);
      
      onChange(option);
      setIsOpen(false);
    };

    // Helper function to check if an option is selected
    const isOptionSelected = (option: string) => {
      // For the "Recently Finished Products" dropdown, the option format is "id|name - process (date)"
      // and the value is "name - process (date)"
      if (option.includes('|')) {
        const optionDisplayText = option.split('|')[1]; // "name - process (date)"
        return value === optionDisplayText;
      }
      return value === option;
    };

    // Helper function to get display text from option
    const getDisplayText = (option: string) => {
      if (option.includes('|')) {
        return option.split('|')[1]; // "name - process (date)"
      }
      return option;
    };

    return (
      <div className="mb-6">
        <label className="block text-gray-700 font-medium mb-2">
          {label}
        </label>
        <div className="relative" ref={dropdownRef}>
          <button
            type="button"
            onClick={() => {
              console.log('Updatedetailsop dropdown clicked, current state:', isOpen);
              console.log('Available options:', options);
              console.log('Current value:', value);
              setIsOpen(!isOpen);
            }}
            className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-left flex items-center justify-between hover:border-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
          >
            <span className="text-gray-700 text-sm">{value}</span>
            <div className="flex items-center space-x-2">
              <ChevronDown 
                className={`w-4 h-4 text-gray-400 transition-transform duration-200 ${
                  isOpen ? 'rotate-180' : ''
                }`} 
              />
            </div>
          </button>
          
          {isOpen && (
            <div className="absolute top-full left-0 right-0 mt-1 bg-white border border-gray-300 rounded-lg shadow-xl z-50 max-h-48 overflow-y-auto">
              {options.length === 0 ? (
                <div className="px-4 py-3 text-gray-500 text-center">
                  <div className="text-sm">No finished products available</div>
                  <div className="text-xs text-gray-400 mt-1">Products will appear here when they are added to live category</div>
                </div>
              ) : (
                options.map((option: string, index: number) => {
                  const selected = isOptionSelected(option);
                  const displayText = getDisplayText(option);
                  return (
                    <div
                      key={index}
                      onClick={(e) => {
                        e.preventDefault();
                        e.stopPropagation();
                        console.log('Dropdown item clicked directly:', option);
                        handleOptionSelect(option, e as any);
                      }}
                      className={`w-full px-4 py-3 text-left hover:bg-blue-50 text-gray-700 first:rounded-t-lg last:rounded-b-lg transition-colors cursor-pointer ${
                        selected ? 'bg-blue-50 border-blue-200' : ''
                      }`}
                    >
                      {displayText}
                      {selected && (
                        <span className="text-xs text-blue-600 ml-2">✓ Selected</span>
                      )}
                    </div>
                  );
                })
              )}
            </div>
          )}
        </div>
      </div>
    );
  };

  const ProcessCheckbox: React.FC<ProcessCheckboxProps> = ({ 
    label, 
    checked, 
    onChange 
  }) => {
    return (
      <div className="bg-blue-50 rounded-lg p-4 mb-3">
        <label className="flex items-center justify-between cursor-pointer">
          <span className="text-gray-700 font-medium">{label}</span>
          <div className="relative">
            <input
              type="checkbox"
              checked={checked}
              onChange={onChange}
              className="sr-only"
            />
            <div className={`w-6 h-6 rounded border-2 flex items-center justify-center transition-colors ${
              checked 
                ? 'bg-blue-700 border-blue-700' 
                : 'bg-white border-gray-300 hover:border-gray-400'
            }`}>
              {checked && (
                <Check className="w-4 h-4 text-white" />
              )}
            </div>
          </div>
        </label>
      </div>
    );
  };

  const allStepsChecked = processSteps.deburring && processSteps.oiling && processSteps.finalInspect;

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Sidebar Component */}
      <Sidebar 
        isOpen={sidebarOpen} 
        onClose={() => setSidebarOpen(false)} 
        username={null}
      />

      {/* Header */}
      <header className="bg-white shadow-sm px-4 py-4 flex items-center justify-between">
        <button 
          onClick={handleMenuClick}
          className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
        >
          <Menu className="w-6 h-6 text-blue-700" />
        </button>
        
        <h1 className="text-xl font-semibold text-blue-700"> Dispatched Menu </h1>
        
        <div className="w-10 h-10 bg-blue-600 rounded-full flex items-center justify-center">
          <span className="text-white font-medium text-lg">A</span>
        </div>
      </header>

      {/* Main Content */}
      <main className="px-6 py-8">
        <div className="max-w-md mx-auto">
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

          {/* Procount Products */}
          <div className="mb-4">
            <div className="flex items-center justify-between mb-2">
              <label className="block text-sm font-medium text-gray-700">
                Recently Finished Products
              </label>
            </div>
            <CustomDropdown
              label=""
              value={selectedProductId ? 
                (() => {
                  const selectedProduct = procountProducts.find(p => p.id === selectedProductId);
                  return selectedProduct ? `${selectedProduct.product?.name} (Count: ${selectedProduct.count})` : '';
                })() 
                : ''
              }
              options={procountProducts.map((p: ProcountProduct) => `${p.id}|${p.product?.name} (Count: ${p.count})`)}
              onChange={(value: string) => {
                console.log('Recently Finished Products dropdown onChange called with:', value);
                if (value && value.includes('|')) {
                  const id = value.split('|')[0];
                  console.log('Extracted ID:', id);
                  const productId = parseInt(id);
                  setSelectedProductId(productId);
                  
                  // Update quantity to not exceed available quantity
                  const selectedProduct = procountProducts.find(p => p.id === productId);
                  if (selectedProduct && quantity > selectedProduct.count) {
                    setQuantity(selectedProduct.count);
                    setQuantityInput(selectedProduct.count.toString());
                  }
                } else {
                  console.log('No valid ID found in value:', value);
                  setSelectedProductId(null);
                }
              }}
            />
          </div>

          {/* Quantity Input */}
          <div className="mb-6">
            <label className="block text-gray-700 font-medium mb-2">
              Quantity
            </label>
            <input
              type="number"
              min="1"
              max={selectedProductId ? (() => {
                const selectedProduct = procountProducts.find(p => p.id === selectedProductId);
                return selectedProduct?.count || 1;
              })() : 1}
              value={quantityInput}
              onChange={(e) => {
                const inputValue = e.target.value;
                setQuantityInput(inputValue);
                
                // Update quantity if it's a valid number
                const newQuantity = parseInt(inputValue);
                if (!isNaN(newQuantity)) {
                  setQuantity(newQuantity);
                }
              }}
              onBlur={(e) => {
                // Apply constraints when user finishes editing
                const inputValue = parseInt(e.target.value) || 1;
                const selectedProduct = procountProducts.find(p => p.id === selectedProductId);
                const maxQuantity = selectedProduct?.count || 1;
                const constrainedValue = Math.max(1, Math.min(maxQuantity, inputValue));
                setQuantity(constrainedValue);
                setQuantityInput(constrainedValue.toString());
              }}
              className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-blue-500 transition-colors"
              placeholder="Enter quantity"
            />
          </div>

          {/* Process Steps */}
          <div className="mb-6">
            <ProcessCheckbox
              label="Deburring"
              checked={processSteps.deburring}
              onChange={() => handleProcessStepChange('deburring')}
            />
            <ProcessCheckbox
              label="Oiling"
              checked={processSteps.oiling}
              onChange={() => handleProcessStepChange('oiling')}
            />
            <ProcessCheckbox
              label="Final Inspect"
              checked={processSteps.finalInspect}
              onChange={() => handleProcessStepChange('finalInspect')}
            />
          </div>
          


          {/* Dispatch Button */}
          <div className="flex justify-center mt-8">
            <button
              onClick={handleUpdate}
              className={`bg-white text-blue-700 px-8 py-3 rounded-full font-medium transition-colors shadow-lg ${allStepsChecked ? 'hover:bg-blue-700 hover:text-white' : 'opacity-50 cursor-not-allowed'}`}
              disabled={!allStepsChecked}
            >
              Dispatch
            </button>
          </div>
        </div>
      </main>
    </div>
  );
}