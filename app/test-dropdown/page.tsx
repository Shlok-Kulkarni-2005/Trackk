'use client';
import React, { useState, useEffect } from 'react';

interface TestProduct {
  id: string;
  name: string;
  process: string;
}

export default function TestDropdownPage() {
  const [products, setProducts] = useState<TestProduct[]>([]);
  const [selectedProduct, setSelectedProduct] = useState<string>('');
  const [isDropdownOpen, setIsDropdownOpen] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);

  // Test data
  const testProducts: TestProduct[] = [
    { id: '1', name: 'Product A', process: 'Cutting' },
    { id: '2', name: 'Product B', process: 'Milling' },
    { id: '3', name: 'Product C', process: 'Drilling' },
  ];

  useEffect(() => {
    // Simulate API call
    setTimeout(() => {
      setProducts(testProducts);
      setIsLoading(false);
    }, 1000);
  }, []);

  const handleProductSelect = (productName: string) => {
    console.log('Test: Product selected:', productName);
    setSelectedProduct(productName);
    setIsDropdownOpen(false);
  };

  return (
    <div className="min-h-screen bg-gray-50 p-8">
      <div className="max-w-md mx-auto">
        <h1 className="text-2xl font-bold mb-6">Dropdown Test</h1>
        
        <div className="mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Test Product Selection
          </label>
          <div className="relative">
            <button
              type="button"
              onClick={() => {
                console.log('Test: Dropdown clicked');
                setIsDropdownOpen(!isDropdownOpen);
              }}
              className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-left flex items-center justify-between hover:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <span className={selectedProduct ? 'text-gray-900' : 'text-gray-500'}>
                {selectedProduct || 'Select a product'}
              </span>
              <span className="text-xs text-gray-400">({products.length} available)</span>
            </button>
            
            {isDropdownOpen && (
              <div className="absolute z-50 w-full mt-1 bg-white border border-gray-300 rounded-lg shadow-xl max-h-60 overflow-y-auto">
                {isLoading ? (
                  <div className="px-4 py-3 text-gray-500 text-center">Loading...</div>
                ) : products.length === 0 ? (
                  <div className="px-4 py-3 text-gray-500 text-center">No products available</div>
                ) : (
                  products.map((product) => (
                    <button
                      key={product.id}
                      type="button"
                      onClick={() => handleProductSelect(product.name)}
                      className="w-full px-4 py-3 text-left hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-b-0 transition-colors focus:outline-none focus:bg-blue-50"
                    >
                      <div className="font-medium text-gray-900">{product.name}</div>
                      <div className="text-sm text-gray-500">{product.process}</div>
                    </button>
                  ))
                )}
              </div>
            )}
          </div>
        </div>

        <div className="bg-white p-4 rounded-lg">
          <h2 className="font-semibold mb-2">Debug Info:</h2>
          <p><strong>Selected Product:</strong> {selectedProduct || 'None'}</p>
          <p><strong>Dropdown Open:</strong> {isDropdownOpen ? 'Yes' : 'No'}</p>
          <p><strong>Products Count:</strong> {products.length}</p>
          <p><strong>Loading:</strong> {isLoading ? 'Yes' : 'No'}</p>
        </div>
      </div>
    </div>
  );
} 