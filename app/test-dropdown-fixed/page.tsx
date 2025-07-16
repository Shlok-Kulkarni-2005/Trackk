'use client';
import React, { useState, useEffect, useRef } from 'react';

interface TestProduct {
  id: number;
  name: string;
  process: string;
  date: string;
}

export default function TestDropdownFixedPage() {
  const [products, setProducts] = useState<TestProduct[]>([]);
  const [selectedProductId, setSelectedProductId] = useState<number | null>(null);
  const [isDropdownOpen, setIsDropdownOpen] = useState<boolean>(false);
  const [isLoading, setIsLoading] = useState<boolean>(true);
  const dropdownRef = useRef<HTMLDivElement>(null);

  // Simulate API call
  useEffect(() => {
    setTimeout(() => {
      const mockProducts: TestProduct[] = [
        { id: 1, name: 'Product A', process: 'Cutting', date: '01/01/2024' },
        { id: 2, name: 'Product B', process: 'Milling', date: '02/01/2024' },
        { id: 3, name: 'Product C', process: 'Drilling', date: '03/01/2024' },
        { id: 4, name: 'Product D', process: 'Finishing', date: '04/01/2024' },
        { id: 5, name: 'Product E', process: 'Assembly', date: '05/01/2024' },
      ];
      setProducts(mockProducts);
      setIsLoading(false);
    }, 1000);
  }, []);

  // Handle click outside
  useEffect(() => {
    function handleClickOutside(event: MouseEvent) {
      if (dropdownRef.current && !dropdownRef.current.contains(event.target as Node)) {
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

  const handleProductSelect = (productId: number) => {
    console.log('Product selected:', productId);
    setSelectedProductId(productId);
    setIsDropdownOpen(false);
  };

  const selectedProduct = products.find(p => p.id === selectedProductId);

  return (
    <div className="min-h-screen bg-gray-50 p-8">
      <div className="max-w-md mx-auto">
        <h1 className="text-2xl font-bold mb-6">Fixed Dropdown Test</h1>
        
        <div className="mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Recently Finished Products
          </label>
          <div className="relative" ref={dropdownRef}>
            <button
              type="button"
              onClick={() => {
                console.log('Dropdown clicked, current state:', isDropdownOpen);
                setIsDropdownOpen(!isDropdownOpen);
              }}
              className="w-full bg-white border border-gray-300 rounded-lg px-4 py-3 text-left flex items-center justify-between hover:border-blue-500 focus:outline-none focus:ring-2 focus:ring-blue-500"
            >
              <span className={selectedProduct ? 'text-gray-900' : 'text-gray-500'}>
                {selectedProduct ? selectedProduct.name : 'Select a product...'}
              </span>
              <div className="flex items-center space-x-2">
                <span className="text-xs text-gray-400">({products.length} available)</span>
                <svg className={`w-5 h-5 text-gray-400 transition-transform ${isDropdownOpen ? 'rotate-180' : ''}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                </svg>
              </div>
            </button>
            
            {isDropdownOpen && (
              <div className="absolute z-50 w-full mt-1 bg-white border border-gray-300 rounded-lg shadow-xl max-h-60 overflow-y-auto">
                {isLoading ? (
                  <div className="px-4 py-3 text-gray-500 text-center">
                    <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-blue-600 mx-auto mb-2"></div>
                    Loading products...
                  </div>
                ) : products.length === 0 ? (
                  <div className="px-4 py-3 text-gray-500 text-center">
                    <div className="text-sm">No products available</div>
                  </div>
                ) : (
                  products.map((product) => (
                    <div
                      key={product.id}
                      onClick={() => handleProductSelect(product.id)}
                      className={`px-4 py-3 hover:bg-blue-50 cursor-pointer border-b border-gray-100 last:border-b-0 transition-colors ${
                        selectedProductId === product.id ? 'bg-blue-50' : ''
                      }`}
                    >
                      <div className="font-medium text-gray-900">{product.name}</div>
                      <div className="text-sm text-gray-500">{product.process} ({product.date})</div>
                      {selectedProductId === product.id && (
                        <div className="text-xs text-blue-600 mt-1">✓ Selected</div>
                      )}
                    </div>
                  ))
                )}
              </div>
            )}
          </div>
        </div>

        <div className="bg-white p-4 rounded-lg">
          <h2 className="font-semibold mb-2">Debug Info:</h2>
          <p><strong>Selected Product ID:</strong> {selectedProductId || 'None'}</p>
          <p><strong>Selected Product Name:</strong> {selectedProduct?.name || 'None'}</p>
          <p><strong>Dropdown Open:</strong> {isDropdownOpen ? 'Yes' : 'No'}</p>
          <p><strong>Products Count:</strong> {products.length}</p>
          <p><strong>Loading:</strong> {isLoading ? 'Yes' : 'No'}</p>
        </div>

        <div className="mt-4">
          <button
            onClick={() => {
              if (products.length > 0) {
                const firstProduct = products[0];
                console.log('Test: Setting first product as selected:', firstProduct);
                setSelectedProductId(firstProduct.id);
              }
            }}
            className="bg-blue-600 text-white px-4 py-2 rounded hover:bg-blue-700"
          >
            Test: Select First Product
          </button>
        </div>
      </div>
    </div>
  );
} 