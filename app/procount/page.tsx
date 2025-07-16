'use client';
import React, { useState, useEffect } from 'react';
import { Menu, Clock } from 'lucide-react';
import Sidebar from '../../components/sidebar';

interface Product {
  id: number;
  name: string;
  count: number;
  status: string;
  machine?: string;
  stage?: string;
  state?: string;
  updatedAt?: string;
}

export default function ProductListPage(): React.ReactElement {
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [products, setProducts] = useState<Product[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  // Fetch product counts from API
  const fetchProductCounts = async (): Promise<void> => {
    try {
      setLoading(true);
      console.log('Fetching product counts...');
      const response = await fetch('/api/procount');
      console.log('Response status:', response.status);
      if (response.ok) {
        const data = await response.json();
        console.log('Raw API data:', data);
        const formattedProducts = data.productCounts.map((item: any) => ({
          id: item.id,
          name: item.product.name,
          count: item.count,
          status: item.state,
          machine: item.machine,
          stage: item.stage,
          state: item.state,
          updatedAt: item.updatedAt
        }));
        console.log('Formatted products:', formattedProducts);
        setProducts(formattedProducts);
      } else {
        console.error('API response not ok:', response.status, response.statusText);
        setError('Failed to fetch product counts');
      }
    } catch (err) {
      console.error('Error fetching product counts:', err);
      setError('Failed to fetch product counts');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    fetchProductCounts();
    
    // Set up real-time updates for procount
    const procountEventSource = new EventSource('/api/procount/stream');
    
    procountEventSource.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        if (data.type === 'productCountUpdated') {
          // Update the products list with the new data
          setProducts(prevProducts => {
            const existingIndex = prevProducts.findIndex(p => p.id === data.productCount.id);
            if (existingIndex >= 0) {
              // Update existing product
              const updated = [...prevProducts];
              updated[existingIndex] = data.productCount;
              return updated;
            } else {
              // Add new product
              return [data.productCount, ...prevProducts];
            }
          });
        }
      } catch (error) {
        console.error('Error parsing procount SSE message:', error);
      }
    };
    
    procountEventSource.onerror = (error) => {
      console.error('Procount SSE connection error:', error);
    };

    // Set up real-time updates for RFD status
    const rfdEventSource = new EventSource('/api/rfd-status/stream');
    
    rfdEventSource.onmessage = (event) => {
      try {
        const data = JSON.parse(event.data);
        if (data.type === 'rfd_status_update') {
          console.log('RFD status update received:', data.data);
          // Refresh the product counts when RFD status is updated
          fetchProductCounts();
        }
      } catch (error) {
        console.error('Error parsing RFD status SSE message:', error);
      }
    };
    
    rfdEventSource.onerror = (error) => {
      console.error('RFD status SSE connection error:', error);
    };
    
    return () => {
      procountEventSource.close();
      rfdEventSource.close();
    };
  }, []);

  const handleProductClick = (product: Product): void => {
    console.log('Product clicked:', product);
    // Navigate to product details or perform action
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
        
        <h1 className="text-xl font-semibold text-blue-700">Product History</h1>
        
        <div className="w-10 h-10 bg-blue-600 rounded-full flex items-center justify-center">
          <span className="text-white font-medium text-lg">A</span>
        </div>
      </header>

      {/* Main Content */}
      <main className="px-4 py-6">
        <div className="max-w-md mx-auto">
          {/* Loading State */}
          {loading && (
            <div className="text-center py-12">
              <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <div className="animate-spin rounded-full h-8 w-8 border-b-2 border-blue-600"></div>
              </div>
              <h3 className="text-gray-500 font-medium mb-2">Loading products...</h3>
            </div>
          )}

          {/* Error State */}
          {error && !loading && (
            <div className="text-center py-12">
              <div className="w-16 h-16 bg-red-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <span className="text-red-400 text-2xl">⚠️</span>
              </div>
              <h3 className="text-red-500 font-medium mb-2">Error loading products</h3>
              <p className="text-red-400 text-sm mb-4">{error}</p>
              <button
                onClick={fetchProductCounts}
                className="bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors"
              >
                Try Again
              </button>
            </div>
          )}

          {/* Product List */}
          {!loading && !error && (
            <div className="space-y-3">
              {products.map((product: Product) => (
                <button
                  key={product.id}
                  onClick={() => handleProductClick(product)}
                  className="w-full bg-white rounded-lg p-4 shadow-sm border border-gray-100 hover:shadow-md hover:border-gray-200 transition-all text-left"
                >
                  <div className="flex items-center space-x-3">
                    {/* Status Icon */}
                    <div className="w-6 h-6 border-2 border-gray-400 rounded-full flex items-center justify-center flex-shrink-0">
                      <div className="w-2 h-2 bg-gray-400 rounded-full"></div>
                    </div>
                    
                    {/* Product Info */}
                    <div className="flex-1 min-w-0">
                      <h3 className="text-gray-900 font-medium text-base mb-1">
                        {product.name}
                      </h3>
                      <div className="flex items-center space-x-1">
                        <span className="text-gray-500 text-sm">Quantity: </span>
                        <span className="text-gray-600 text-sm font-medium">
                          {product.count}
                        </span>
                      </div>
                    </div>

                    {/* count Icon */}
                    <div className="flex-shrink-0">
                      <Clock className="w-4 h-4 text-gray-400" />
                    </div>
                  </div>
                </button>
              ))}
            </div>
          )}

          {/* Empty State (if no products) */}
          {!loading && !error && products.length === 0 && (
            <div className="text-center py-12">
              <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <span className="text-gray-400 text-2xl">📦</span>
              </div>
              <h3 className="text-gray-500 font-medium mb-2">No products found</h3>
              <p className="text-gray-400 text-sm">
                Products will appear here when RFD machine is selected with OFF status in Add Jobs
              </p>
            </div>
          )}

        </div>
      </main>
    </div>
  );
}