'use client';
import React, { useState, useEffect, useRef } from 'react';
import { Menu, ChevronDown } from 'lucide-react';
import Sidebar from '../../components/sidebar';

interface LiveProduct {
  id: string;
  name: string;
  process: string;
  status: string;
  jobId: number;
}

interface DispatchForm {
  product: string;
  quantity: number;
  destination: string;
  notes: string;
}

export default function DispatchPage() {
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [liveProducts, setLiveProducts] = useState<LiveProduct[]>([]);
  const [selectedProduct, setSelectedProduct] = useState<string>('');
  const [isDropdownOpen, setIsDropdownOpen] = useState<boolean>(false);
  const [dispatchForm, setDispatchForm] = useState<DispatchForm>({
    product: '',
    quantity: 1,
    destination: '',
    notes: ''
  });
  const [quantityInput, setQuantityInput] = useState<string>('1');
  const dropdownRef = useRef<HTMLDivElement>(null);
  const [isLoading, setIsLoading] = useState<boolean>(true);
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

  // Fetch live products (state ON) from backend
  useEffect(() => {
    async function fetchLiveProducts() {
      setIsLoading(true);
      try {
        const res = await fetch('/api/live-products/dispatchable');
        if (res.ok) {
          const data = await res.json();
          const liveProductsArr = data.dispatchableProducts.map((product: any) => ({
            id: product.id,
            name: product.name,
            process: product.process,
            status: product.status,
            jobId: product.jobId,
          }));
          console.log('Live products loaded:', liveProductsArr); // Debug log
          console.log('Raw API response:', data); // Debug raw response
          setLiveProducts(liveProductsArr);
          if (liveProductsArr.length > 0 && !selectedProduct) {
            setSelectedProduct(liveProductsArr[0].name);
            setDispatchForm(prev => ({ ...prev, product: liveProductsArr[0].name }));
          }
        }
      } catch (error) {
        console.error('Error fetching live products:', error);
      } finally {
        setIsLoading(false);
      }
    }
    fetchLiveProducts();

    // Set up real-time updates using Server-Sent Events
    const eventSource = new EventSource('/api/live-products/dispatchable/stream');
    
    eventSource.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        
        if (data.type === 'product_moved_to_past') {
          // Remove the product from the live products list
          setLiveProducts(prev => prev.filter(product => product.id !== data.data.productId));
          
          // If the removed product was selected, clear the selection
          if (selectedProduct === data.data.productName) {
            setSelectedProduct('');
            setDispatchForm(prev => ({ ...prev, product: '' }));
          }
          
          console.log(`Product ${data.data.productName} moved to past products`);
        }
      } catch (error) {
        console.error('Error parsing SSE data:', error);
      }
    };

    eventSource.onerror = (error) => {
      console.error('SSE connection error:', error);
      eventSource.close();
    };

    // Cleanup on unmount
    return () => {
      eventSource.close();
    };
  }, [selectedProduct]);

  // Handle click outside dropdown to close it
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
        console.log('Click outside detected, closing dropdown');
        setIsDropdownOpen(false);
      }
    }

    if (isDropdownOpen) {
      document.addEventListener('mousedown', handleClickOutside);
    }

    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, [isDropdownOpen]);

  // Debug selectedProduct changes
  useEffect(() => {
    console.log('Selected product changed to:', selectedProduct);
    console.log('Dispatch form product field:', dispatchForm.product);
  }, [selectedProduct, dispatchForm.product]);

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  const handleProductSelect = (productName: string, event?: React.MouseEvent): void => {
    event?.preventDefault();
    event?.stopPropagation();
    
    console.log('Product selected:', productName); // Debug log
    console.log('Previous selected product:', selectedProduct);
    
    setSelectedProduct(productName);
    setDispatchForm(prev => ({ ...prev, product: productName }));
    setIsDropdownOpen(false);
    
    console.log('Product selection completed');
  };

  const handleFormChange = (field: keyof DispatchForm, value: string | number): void => {
    setDispatchForm(prev => ({ ...prev, [field]: value }));
  };

  const handleDispatch = async (): Promise<void> => {
    if (!dispatchForm.product || !dispatchForm.destination) {
      setMessage({ text: 'Please fill in all required fields', type: 'info' });
      return;
    }

    // Find the selected product to get its jobId
    const selectedProductData = liveProducts.find(p => p.name === selectedProduct);
    if (!selectedProductData) {
      setMessage({ text: 'Selected product not found', type: 'error' });
      return;
    }

    try {
      // First, dispatch the product
      const dispatchRes = await fetch('/api/admin/dispatched-items', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(dispatchForm)
      });

      if (dispatchRes.ok) {
        // Then, automatically move the product from live to past
        const moveRes = await fetch('/api/products/lifecycle', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            productId: selectedProductData.id,
            jobId: selectedProductData.jobId,
            action: 'move_to_past',
            reason: 'dispatched'
          })
        });

        if (moveRes.ok) {
          setMessage({ text: 'Product dispatched successfully and moved to past products!', type: 'success' });
          // Reset form
          setDispatchForm({
            product: '',
            quantity: 1,
            destination: '',
            notes: ''
          });
          setQuantityInput('1');
          setSelectedProduct('');
          // Refresh the live products list
          const refreshRes = await fetch('/api/live-products/dispatchable');
          if (refreshRes.ok) {
            const data = await refreshRes.json();
            const liveProductsArr = data.dispatchableProducts.map((product: any) => ({
              id: product.id,
              name: product.name,
              process: product.process,
              status: product.status,
              jobId: product.jobId,
            }));
            setLiveProducts(liveProductsArr);
          }
        } else {
          setMessage({ text: 'Product dispatched but failed to move to past products', type: 'error' });
        }
      } else {
        setMessage({ text: 'Failed to dispatch product', type: 'error' });
      }
    } catch (error) {
      console.error('Error dispatching product:', error);
      setMessage({ text: 'Failed to dispatch product', type: 'error' });
    }
  };

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
        
        <h1 className="text-xl font-semibold text-blue-700">Dispatch Products</h1>
        
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

          {/* Product Selection Dropdown */}
          <div className="mb-6">
            <div className="flex items-center justify-between mb-2">
              <label className="block text-sm font-medium text-gray-700">
                Select Product to Dispatch
              </label>
              <button
                type="button"
                onClick={() => {
                  setIsLoading(true);
                  fetch('/api/live-products/dispatchable')
                    .then(res => res.json())
                    .then(data => {
                      const liveProductsArr = data.dispatchableProducts.map((product: any) => ({
                        id: product.id,
                        name: product.name,
                        process: product.process,
                        status: product.status,
                        jobId: product.jobId,
                      }));
                      setLiveProducts(liveProductsArr);
                      setIsLoading(false);
                    })
                    .catch(error => {
                      console.error('Error refreshing products:', error);
                      setIsLoading(false);
                    });
                }}
                className="text-xs text-blue-600 hover:text-blue-800 underline"
              >
                Refresh
              </button>
            </div>
            <div className="relative" ref={dropdownRef}>
              <button
                type="button"
                onClick={() => {
                  console.log('Dropdown clicked, current state:', isDropdownOpen);
                  console.log('Available products:', liveProducts);
                  setIsDropdownOpen(!isDropdownOpen);
                }}
                className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-left flex items-center justify-between hover:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500"
              >
                <span className={selectedProduct ? 'text-gray-900' : 'text-gray-500'}>
                  {selectedProduct || 'Select a product'}
                </span>
                <div className="flex items-center space-x-2">
                  <ChevronDown className={`w-5 h-5 text-gray-400 transition-transform ${isDropdownOpen ? 'rotate-180' : ''}`} />
                </div>
              </button>
              
              {isDropdownOpen && (
                <div className="absolute z-50 w-full mt-1 bg-white border border-gray-300 rounded-lg shadow-xl max-h-60 overflow-y-auto">
                  {isLoading ? (
                    <div className="px-4 py-3 text-gray-500 text-center">
                      <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600 mx-auto mb-2"></div>
                      Loading products...
                    </div>
                  ) : liveProducts.length === 0 ? (
                    <div className="px-4 py-3 text-gray-500 text-center">
                      <div className="text-sm">No live products available</div>
                      <div className="text-xs text-gray-400 mt-1">Products will appear here when they are in production</div>
                    </div>
                  ) : (
                    liveProducts.map((product) => (
                      <button
                        key={product.id}
                        type="button"
                        onClick={(e) => {
                          e.preventDefault();
                          e.stopPropagation();
                          console.log('Dropdown item clicked:', product.name);
                          handleProductSelect(product.name, e);
                        }}
                        className={`w-full px-4 py-3 text-left hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-b-0 transition-colors focus:outline-none focus:bg-blue-50 ${
                          selectedProduct === product.name ? 'bg-blue-50 border-blue-200' : ''
                        }`}
                      >
                        <div className="font-medium text-gray-900">{product.name}</div>
                        <div className="text-sm text-gray-500">{product.process}</div>
                        {selectedProduct === product.name && (
                          <div className="text-xs text-blue-600 mt-1">✓ Selected</div>
                        )}
                      </button>
                    ))
                  )}
                </div>
              )}
            </div>
          </div>

          {/* Dispatch Form */}
          <div className="bg-white rounded-lg shadow-sm p-6 space-y-4">
            <h2 className="text-lg font-semibold text-gray-900 mb-4">Dispatch Details</h2>
            
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Quantity
              </label>
              <input
                type="number"
                min="1"
                value={quantityInput}
                onChange={(e) => {
                  const inputValue = e.target.value;
                  setQuantityInput(inputValue);
                  
                  // Update quantity if it's a valid number
                  const newQuantity = parseInt(inputValue);
                  if (!isNaN(newQuantity)) {
                    handleFormChange('quantity', newQuantity);
                  }
                }}
                onBlur={(e) => {
                  // Apply constraints when user finishes editing
                  const inputValue = parseInt(e.target.value) || 1;
                  const constrainedValue = Math.max(1, inputValue);
                  handleFormChange('quantity', constrainedValue);
                  setQuantityInput(constrainedValue.toString());
                }}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Destination *
              </label>
              <input
                type="text"
                value={dispatchForm.destination}
                onChange={(e) => handleFormChange('destination', e.target.value)}
                placeholder="Enter destination"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">
                Notes
              </label>
              <textarea
                value={dispatchForm.notes}
                onChange={(e) => handleFormChange('notes', e.target.value)}
                placeholder="Additional notes (optional)"
                rows={3}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500"
              />
            </div>

            <button
              onClick={handleDispatch}
              className="w-full bg-blue-600 text-white py-3 rounded-lg font-medium hover:bg-blue-700 transition-colors"
            >
              Dispatch Product
            </button>
          </div>
        </div>
      </main>
    </div>
  );
} 