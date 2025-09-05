'use client';
import React, { useState, useEffect } from 'react';
import { Menu, X } from 'lucide-react';
import Sidebar from '../../components/sidebar';

interface Product {
  id: string;
  name: string;
  operation: string;
  date: string;
  state: 'ON' | 'OFF';
  quantity: number;
}

export default function WorkPanelInterface() {
  const [sidebarOpen, setSidebarOpen] = useState<boolean>(false);
  const [currentView, setCurrentView] = useState<'live' | 'past'>('live');
  const [selectedProduct, setSelectedProduct] = useState<Product | null>(null);
  const [productData, setProductData] = useState<Product[]>([]);
  const [allJobs, setAllJobs] = useState<any[]>([]);
  const [now, setNow] = useState(Date.now());
  const [message, setMessage] = useState<{ text: string; type: 'success' | 'error' | 'info' } | null>(null);

  // Message auto-clear
  useEffect(() => {
    if (message) {
      const timer = setTimeout(() => {
        setMessage(null);
      }, 5000);
      return () => clearTimeout(timer);
    }
  }, [message]);

  useEffect(() => {
    const interval = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(interval);
  }, []);

  useEffect(() => {
    async function fetchJobs() {
      try {
        const res = await fetch('/api/jobs');
        if (res.ok) {
          const data = await res.json();
          setAllJobs(data.jobs);
          
          // Group jobs by product and machine
          const jobsByKey: { [key: string]: any[] } = {};
          data.jobs.forEach((job: any) => {
            const key = `${job.product.id}__${job.machine.name}__${job.state}`;
            if (!jobsByKey[key]) jobsByKey[key] = [];
            jobsByKey[key].push(job);
          });
          
          // Build productData with ON and OFF quantities for each (product, machine) pair
          const products: Product[] = [];
          Object.entries(jobsByKey).forEach(([key, jobs]) => {
            const [productId, machineName, state] = key.split('__');
            const productName = jobs[0].product.name;
            const date = new Date(jobs[0].createdAt).toLocaleDateString();
            products.push({
              id: `${productId}__${machineName}__${state}`,
              name: productName,
              operation: machineName,
              date,
              state: state as 'ON' | 'OFF',
              quantity: jobs.length,
            });
          });
          setProductData(products);
        }
      } catch (error) {
        console.error('Error fetching jobs:', error);
      }
    }
    fetchJobs();
  }, []);

  // Helper to format duration as HH:mm:ss
  const formatDuration = (ms: number) => {
    const totalSeconds = Math.floor(ms / 1000);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${seconds.toString().padStart(2, '0')}`;
  };

  const getLiveProducts = (): Product[] => {
    return productData.filter(product => product.state === 'ON' && product.quantity > 0);
  };

  const getPastProducts = (): Product[] => {
    return productData.filter(product => product.state === 'OFF' && product.quantity > 0);
  };

  const getStatusColor = (status: string): string => {
    switch (status) {
      case 'ON': return 'text-green-600';
      case 'OFF': return 'text-gray-500';
      default: return 'text-gray-500';
    }
  };

  const getStatusDotColor = (status: string): string => {
    switch (status) {
      case 'ON': return 'bg-green-500';
      case 'OFF': return 'bg-gray-400';
      default: return 'bg-gray-400';
    }
  };

  const handleSeeDetails = (product: Product): void => {
    setSelectedProduct(product);
  };

  const handleClose = (): void => {
    setSelectedProduct(null);
  };

  const renderProductList = (products: Product[]) => (
    <div className="space-y-4">
      {products.map((product, idx) => (
        <div key={idx} className="bg-white rounded-lg p-4 shadow-sm border border-gray-200 hover:shadow-md transition-shadow">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center space-x-3">
              <div className={`w-3 h-3 rounded-full ${getStatusDotColor(product.state)}`}></div>
              <h3 className="text-gray-900 font-medium text-base">{product.name}</h3>
            </div>
            <div className="flex items-center space-x-2">
              <span className={`px-2 py-1 rounded-full text-xs font-medium ${getStatusColor(product.state)}`}>
                {product.state}
              </span>
              <button
                onClick={() => handleSeeDetails(product)}
                className="text-blue-600 hover:text-blue-800 text-sm font-medium"
              >
                See Details
              </button>
            </div>
          </div>
          
          <div className="grid grid-cols-2 gap-4 text-sm">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Operation</label>
              <div className="text-gray-900">{product.operation}</div>
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Date</label>
              <div className="text-gray-900">{product.date}</div>
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Timing</label>
              <div className="text-gray-900">N/A</div>
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Quantity</label>
              <div className="text-gray-900">{product.quantity}</div>
            </div>
          </div>
        </div>
      ))}
    </div>
  );

  const renderDetailsView = () => {
    if (!selectedProduct) return null;
    
    const jobs = allJobs.filter(job => 
      job.product.id.toString() === selectedProduct.id.split('__')[0] && 
      job.machine.name === selectedProduct.operation
    );
    
    let timeInfo = '';
    if (selectedProduct.state === 'ON') {
      const lastOnJob = jobs.filter(j => j.state === 'ON').sort((a, b) => new Date(b.createdAt).getTime() - new Date(a.createdAt).getTime())[0];
      if (lastOnJob) {
        const start = new Date(lastOnJob.createdAt).getTime();
        timeInfo = `Running: ${formatDuration(now - start)}`;
      }
    } else {
      // For past products: use earliest createdAt and latest updatedAt
      const sortedJobs = jobs.sort((a, b) => new Date(a.createdAt).getTime() - new Date(b.createdAt).getTime());
      if (sortedJobs.length > 0) {
        const firstJob = sortedJobs[0];
        const lastJob = sortedJobs[sortedJobs.length - 1];
        const duration = new Date(lastJob.updatedAt || lastJob.createdAt).getTime() - new Date(firstJob.createdAt).getTime();
        if (duration > 0) {
          timeInfo = `Completed in ${formatDuration(duration)}`;
        } else {
          timeInfo = `Completed`;
        }
      }
    }
    // Calculate quantity
    let quantity = 0;
    if (selectedProduct.state === 'ON') {
      // Live product: count ON jobs
      quantity = jobs.filter(j => j.state === 'ON').length;
    } else {
      // Past product: count OFF jobs (since all ON jobs were turned OFF at once)
      quantity = jobs.filter(j => j.state === 'OFF').length;
    }
    return (
      <div className="bg-white rounded-lg shadow-sm border border-gray-200 p-6">
        <div className="flex items-center justify-between mb-6">
          <h2 className="text-xl font-semibold text-gray-900">{selectedProduct.name}</h2>
          <button
            onClick={handleClose}
            className="text-gray-400 hover:text-gray-600"
          >
            <X className="w-6 h-6" />
          </button>
        </div>
        
        <div className="grid grid-cols-2 gap-6">
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Operation</label>
            <div className="text-gray-900 text-lg">{selectedProduct.operation}</div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Status</label>
            <div className="text-gray-900 text-lg">{selectedProduct.state}</div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Date</label>
            <div className="text-gray-900 text-lg">{selectedProduct.date}</div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Timing</label>
            <div className="text-gray-900 text-lg">{timeInfo}</div>
          </div>
          <div>
            <label className="block text-sm font-medium text-gray-700 mb-2">Quantity</label>
            <div className="text-gray-900 text-lg">{quantity}</div>
          </div>
        </div>
      </div>
    );
  };

  return (
    <div className="min-h-screen bg-gray-100">
      {/* Sidebar Component */}
      <Sidebar 
        isOpen={sidebarOpen} 
        onClose={() => setSidebarOpen(false)} 
        username={null}
      />

      {/* Header */}
      <header className="bg-white px-6 py-4 flex items-center justify-between sticky top-0 z-20">
        <button
          onClick={() => setSidebarOpen(true)}
          className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
        >
          <Menu className="w-6 h-6 text-blue-700" />
        </button>

        <h1 className="text-xl font-semibold text-blue-700">Work Panel</h1>

        <div className="w-10 h-10 bg-blue-600 rounded-full flex items-center justify-center"> 
          <span className="text-white font-medium text-lg">A</span>
        </div>
      </header>

      {/* Section Tabs */}
      <div className="flex border-b">
        <button
          onClick={() => setCurrentView('live')}
          className={`flex-1 py-3 flex items-center justify-center ${
            currentView === 'live'
              ? 'bg-blue-100 text-blue-700 border-b-2 border-blue-700'
              : 'text-gray-600 hover:bg-gray-100'
          }`}
        >
          Live Products
        </button>
        <button
          onClick={() => setCurrentView('past')}
          className={`flex-1 py-3 flex items-center justify-center ${
            currentView === 'past'
              ? 'bg-blue-100 text-blue-700 border-b-2 border-blue-700'
              : 'text-gray-600 hover:bg-gray-100'
          }`}
        >
          Past Products
        </button>
      </div>

      {/* Main Content */}
      <main className="px-6 py-8">
        <div className="max-w-2xl mx-auto">
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

          {selectedProduct ? (
            renderDetailsView()
          ) : (
            <>
              {currentView === 'live' ? (
                getLiveProducts().length > 0 ? (
                  renderProductList(getLiveProducts())
                ) : (
                  <div className="text-center py-12 bg-white rounded-lg">
                    <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                      <span className="text-gray-400 text-2xl">🏁</span>
                    </div>
                    <h3 className="text-gray-500 font-medium mb-2">No Live Products</h3>
                    <p className="text-gray-400 text-sm">
                      All products have completed their lifecycle
                    </p>
                  </div>
                )
              ) : (
                getPastProducts().length > 0 ? (
                  renderProductList(getPastProducts())
                ) : (
                  <div className="text-center py-12 bg-white rounded-lg">
                    <div className="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center mx-auto mb-4">
                      <span className="text-gray-400 text-2xl">📦</span>
                    </div>
                    <h3 className="text-gray-500 font-medium mb-2">No Past Products</h3>
                    <p className="text-gray-400 text-sm">
                      No completed products found
                    </p>
                  </div>
                )
              )}
            </>
          )}
        </div>
      </main>
    </div>
  );
}