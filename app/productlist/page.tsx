'use client';
import React, { useState, useEffect } from 'react';
import { Menu, Clock, Trash2, Check, X } from 'lucide-react';
import Sidebar from '../../components/sidebar';

interface Product {
  id: number;
  name: string;
  process: string;
  status: string;
  date: string;
}

export default function ProductListPage(): React.ReactElement {
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [products, setProducts] = useState<Product[]>([]);
  const [isConfirmingClear, setIsConfirmingClear] = useState<boolean>(false);

  // Initial fetch on mount (if not cleared)
  useEffect(() => {
    async function fetchJobs() {
      const res = await fetch('/api/jobs');
      if (res.ok) {
        const data = await res.json();
        // Get hidden IDs from localStorage
        const hiddenIdsRaw = localStorage.getItem('productlist_hidden_ids');
        let hiddenIds: number[] = [];
        if (hiddenIdsRaw) {
          try {
            hiddenIds = JSON.parse(hiddenIdsRaw);
          } catch {}
        }
        // Group jobs by productId and keep only the latest job for each product
        const latestJobByProduct: { [productId: string]: any } = {};
        data.jobs.forEach((job: any) => {
          const productId = job.product.id;
          if (
            !latestJobByProduct[productId] ||
            new Date(job.createdAt) > new Date(latestJobByProduct[productId].createdAt)
          ) {
            latestJobByProduct[productId] = job;
          }
        });
        // Map to Product[] and sort by date descending, filter out hidden
        const productsArr = Object.values(latestJobByProduct)
          .map((job: any) => ({
            id: job.product.id,
            name: job.product.name,
            process: job.machine.name,
            status: job.state,
            date: new Date(job.createdAt).toLocaleDateString(),
          }))
          .filter((product: any) => !hiddenIds.includes(product.id))
          .sort((a: any, b: any) => new Date(b.date).getTime() - new Date(a.date).getTime());
        setProducts(productsArr);
      }
    }
    fetchJobs();
    // Poll every 5 seconds for new jobs
    const interval = setInterval(fetchJobs, 5000);
    return () => clearInterval(interval);
  }, []);

  const handleMenuClick = (): void => {
    setSidebarOpen(true);
  };

  const handleProductClick = (product: Product): void => {
    console.log('Product clicked:', product);
    // Navigate to product details or perform action
  };

  const handleClearProducts = (): void => {
    setIsConfirmingClear(true);
  };

  const confirmClearProducts = async (): Promise<void> => {
    try {
      // Fetch all jobs from backend
      let allIds: number[] = [];
      try {
        const res = await fetch('/api/jobs');
        if (res.ok) {
          const data = await res.json();
          const idsSet = new Set<number>();
          data.jobs.forEach((job: any) => {
            idsSet.add(job.product.id);
          });
          allIds = Array.from(idsSet);
        } else {
          throw new Error('API returned error');
        }
      } catch (fetchError) {
        // Fallback: hide only currently visible products
        allIds = products.map((product) => product.id);
        alert('Failed to fetch all jobs from backend. Only currently visible products will be hidden.');
      }
      // Merge with any already hidden
      const hiddenIdsRaw = localStorage.getItem('productlist_hidden_ids');
      let hiddenIds: number[] = [];
      if (hiddenIdsRaw) {
        try {
          hiddenIds = JSON.parse(hiddenIdsRaw);
        } catch {}
      }
      const merged = Array.from(new Set([...hiddenIds, ...allIds]));
      localStorage.setItem('productlist_hidden_ids', JSON.stringify(merged));
      setProducts([]);
      setIsConfirmingClear(false);
      console.log('Product list cleared (hidden locally, all IDs or fallback)');
    } catch (error) {
      console.error('Error clearing product list:', error);
      alert('Failed to clear product list.');
    }
  };

  const cancelClearProducts = (): void => {
    setIsConfirmingClear(false);
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
        
        <h1 className="text-xl font-semibold text-blue-700">Recently Added</h1>
        
        <div className="w-10 h-10 bg-blue-600 rounded-full flex items-center justify-center">
          <span className="text-white font-medium text-lg">A</span>
        </div>
      </header>

      {/* Main Content */}
      <main className="px-4 py-6">
        <div className="max-w-md mx-auto">
          {/* Clear Button with Confirmation */}
          {products.length > 0 && (
            <div className="mb-4 flex justify-end items-center space-x-3">
              {isConfirmingClear ? (
                <>
                  <span className="text-sm text-gray-600">Confirm clear all?</span>
                  <button
                    onClick={confirmClearProducts}
                    className="text-green-600 hover:text-green-700 transition-colors"
                  >
                    <Check className="w-5 h-5" />
                  </button>
                  <button
                    onClick={cancelClearProducts}
                    className="text-red-600 hover:text-red-700 transition-colors"
                  >
                    <X className="w-5 h-5" />
                  </button>
                </>
              ) : (
                <button
                  onClick={handleClearProducts}
                  className="flex items-center text-red-600 hover:text-red-700 text-sm font-medium transition-colors"
                >
                  <Trash2 className="w-4 h-4 mr-2" />
                  Clear All
                </button>
              )}
            </div>
          )}

          {/* Product List */}
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
                      <span className="text-gray-500 text-sm">Operation:</span>
                      <span className="text-gray-600 text-sm font-medium">
                        {product.process}
                      </span>
                    </div>
                    <div className="flex items-center space-x-1">
                      <span className="text-gray-500 text-xs">Added:</span>
                      <span className="text-gray-400 text-xs font-medium">
                        {product.date}
                      </span>
                    </div>
                  </div>
                  {/* Process Icon */}
                  <div className="flex-shrink-0">
                    <Clock className="w-4 h-4 text-gray-400" />
                  </div>
                </div>
              </button>
            ))}
          </div>

          {/* Empty State (if no products) */}
          {products.length === 0 && (
            <div className="text-center py-12">
              <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                <span className="text-gray-400 text-2xl">📦</span>
              </div>
              <h3 className="text-gray-500 font-medium mb-2">No products found</h3>
              <p className="text-gray-400 text-sm">
                Add products to see them listed here
              </p>
            </div>
          )}

          {/* Summary Stats */}
          {products.length > 0 && (
            <div className="mt-8 bg-white rounded-lg p-4 shadow-sm border border-gray-100">
              <div className="text-center">
                <div className="text-2xl font-bold text-gray-900 mb-1">
                  {products.length}
                </div>
                <div className="text-gray-500 text-sm">
                  Total Products
                </div>
              </div>
              {/* You can add more summary stats here if needed */}
            </div>
          )}
        </div>
      </main>
    </div>
  );
}